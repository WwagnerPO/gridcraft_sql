//! Running VBA macros from the interface: Developer › Macros (Alt+F8), Enable Content for a
//! workbook's macros, `Workbook_Open`.
//!
//! A macro runs on a worker thread that owns the session meanwhile (the window shows only the
//! macro's progress, as Excel's does while a macro runs). `MsgBox` and `InputBox` come back to
//! this thread as questions; the worker waits for the answer. Stop interrupts it. On the web,
//! which has no threads, a macro runs at once and its message boxes answer their first button.

use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc;

use egui::{Color32, RichText, Ui, vec2};
use gridcraft_engine::Session;
use gridcraft_engine::vba_host::MacroUi;
use serde_json::{Value as Json, json};

use crate::SheetApp;

/// A question the macro waits on.
pub enum Ask {
    Msg { prompt: String, buttons: i64, title: String, reply: mpsc::Sender<i64> },
    Input { prompt: String, title: String, reply: mpsc::Sender<Option<String>> },
}

enum Event {
    Ask(Ask),
    Debug(String),
    Status(Option<String>),
}

/// The worker's UI: questions go to the window and block until answered.
struct ChannelUi {
    tx: mpsc::Sender<Event>,
}

impl MacroUi for ChannelUi {
    fn msgbox(&mut self, prompt: &str, buttons: i64, title: &str) -> i64 {
        let (reply, answer) = mpsc::channel();
        if self.tx.send(Event::Ask(Ask::Msg { prompt: prompt.into(), buttons, title: title.into(), reply })).is_err() {
            return 2;
        }
        answer.recv().unwrap_or(2)
    }
    fn inputbox(&mut self, prompt: &str, title: &str, default: &str) -> Option<String> {
        let (reply, answer) = mpsc::channel();
        let _ = default;
        if self.tx.send(Event::Ask(Ask::Input { prompt: prompt.into(), title: title.into(), reply })).is_err() {
            return None;
        }
        answer.recv().unwrap_or(None)
    }
    fn debug_print(&mut self, text: &str) {
        let _ = self.tx.send(Event::Debug(text.into()));
    }
    fn status(&mut self, text: Option<String>) {
        let _ = self.tx.send(Event::Status(text));
    }
}

type Outcome = (Session, Result<String, String>);

/// A macro running on its worker.
pub struct MacroRun {
    pub name: String,
    rx: mpsc::Receiver<Event>,
    handle: Option<std::thread::JoinHandle<Outcome>>,
    ask: Option<Ask>,
    input: String,
    cancel: Arc<AtomicBool>,
    status: Option<String>,
    pub output: Vec<String>,
}

/// Starts macro `name` of the active workbook.
pub fn start(app: &mut SheetApp, name: &str) {
    if app.ext.macro_run.is_some() {
        return;
    }
    if !app.commit_edit(0, 0, false, false) {
        return;
    }
    #[cfg(target_arch = "wasm32")]
    {
        let mut ui = gridcraft_engine::vba_host::Headless::default();
        let r = gridcraft_engine::vba_host::run_macro(&mut app.session, name, &mut ui);
        finish(app, name, r.map(|v| v.to_str().unwrap_or_default()).map_err(|e| e.to_string()), ui.output);
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        let mut session = std::mem::replace(&mut app.session, Session::new());
        let (tx, rx) = mpsc::channel();
        let cancel = Arc::new(AtomicBool::new(false));
        let stop = cancel.clone();
        let macro_name = name.to_string();
        let spawned = std::thread::Builder::new().name("gridcraft-macro".into()).spawn(move || {
            let mut ui = ChannelUi { tx };
            let r = gridcraft_engine::vba_host::run_macro_with(&mut session, &macro_name, &mut ui, Some(stop));
            (session, r.map(|v| v.to_str().unwrap_or_default()).map_err(|e| e.to_string()))
        });
        match spawned {
            Ok(handle) => {
                app.ext.macro_run = Some(MacroRun {
                    name: name.into(),
                    rx,
                    handle: Some(handle),
                    ask: None,
                    input: String::new(),
                    cancel,
                    status: None,
                    output: vec![],
                })
            }
            Err(e) => app.message = Some(("Macro".into(), format!("The macro couldn't start: {e}"))),
        }
    }
}

fn finish(app: &mut SheetApp, name: &str, r: Result<String, String>, output: Vec<String>) {
    match r {
        Ok(_) => {
            let msg = if output.is_empty() { format!("{} {name}", app.ui.language.tr("Macro finished:")) } else { output.join("\n") };
            app.toast = Some((msg, crate::now_ms()));
        }
        Err(e) => app.message = Some(("Microsoft Visual Basic for Applications".replace("Microsoft ", ""), e)),
    }
}

/// Draws the running macro (its questions and a Stop button) instead of the workbook, and
/// takes the session back when the macro ends. `true` while a macro runs.
pub fn frame(app: &mut SheetApp, ui: &mut Ui) -> bool {
    let Some(run) = app.ext.macro_run.as_mut() else { return false };
    while let Ok(ev) = run.rx.try_recv() {
        match ev {
            Event::Ask(a) => run.ask = Some(a),
            Event::Debug(t) => run.output.push(t),
            Event::Status(t) => run.status = t,
        }
    }
    if run.handle.as_ref().is_some_and(|h| h.is_finished())
        && let Some(h) = run.handle.take()
    {
        let name = run.name.clone();
        let output = std::mem::take(&mut run.output);
        app.ext.macro_run = None;
        match h.join() {
            Ok((session, r)) => {
                app.session = session;
                finish(app, &name, r, output);
            }
            Err(_) => app.message = Some(("Macro".into(), "The macro stopped unexpectedly; the workbook could not be recovered from it.".into())),
        }
        ui.ctx().request_repaint();
        return false;
    }
    let lang = app.ui.language;
    let t = crate::theme::Tokens::get(ui.ctx());
    let mut answer_msg: Option<i64> = None;
    let mut answer_input: Option<Option<String>> = None;
    egui::Frame::NONE.fill(t.window).show(ui, |ui| {
        ui.set_min_size(ui.available_size());
        ui.vertical_centered(|ui| {
            ui.add_space(ui.available_height() * 0.25);
            ui.label(RichText::new(format!("{} {}", lang.tr("Running macro"), run.name)).font(crate::theme::ui_bold(15.0)));
            if let Some(s) = &run.status {
                ui.label(RichText::new(s).weak());
            }
            ui.add_space(12.0);
            match &run.ask {
                Some(Ask::Msg { prompt, buttons, title, .. }) => {
                    egui::Frame::group(ui.style()).fill(t.ribbon).inner_margin(16).show(ui, |ui| {
                        ui.set_max_width(460.0);
                        ui.label(RichText::new(title).strong());
                        ui.add_space(6.0);
                        let icon = match buttons & 0x70 {
                            16 => "⛔ ",
                            32 => "❓ ",
                            48 => "⚠ ",
                            64 => "ℹ ",
                            _ => "",
                        };
                        ui.label(format!("{icon}{prompt}"));
                        ui.add_space(10.0);
                        let set: &[(&str, i64)] = match buttons & 7 {
                            1 => &[("OK", 1), ("Cancel", 2)],
                            2 => &[("Abort", 3), ("Retry", 4), ("Ignore", 5)],
                            3 => &[("Yes", 6), ("No", 7), ("Cancel", 2)],
                            4 => &[("Yes", 6), ("No", 7)],
                            5 => &[("Retry", 4), ("Cancel", 2)],
                            _ => &[("OK", 1)],
                        };
                        ui.horizontal(|ui| {
                            for (label, code) in set {
                                if ui.add(egui::Button::new(lang.tr(label)).min_size(vec2(80.0, 24.0))).clicked() {
                                    answer_msg = Some(*code);
                                }
                            }
                        });
                        if ui.input(|i| i.key_pressed(egui::Key::Enter)) {
                            answer_msg = set.first().map(|x| x.1);
                        }
                    });
                }
                Some(Ask::Input { prompt, title, .. }) => {
                    egui::Frame::group(ui.style()).fill(t.ribbon).inner_margin(16).show(ui, |ui| {
                        ui.set_max_width(460.0);
                        ui.label(RichText::new(title).strong());
                        ui.add_space(6.0);
                        ui.label(prompt);
                        let r = ui.add(egui::TextEdit::singleline(&mut run.input).desired_width(380.0));
                        r.request_focus();
                        ui.add_space(8.0);
                        ui.horizontal(|ui| {
                            if ui.add(egui::Button::new(lang.tr("OK")).min_size(vec2(80.0, 24.0))).clicked()
                                || ui.input(|i| i.key_pressed(egui::Key::Enter))
                            {
                                answer_input = Some(Some(std::mem::take(&mut run.input)));
                            }
                            if ui.add(egui::Button::new(lang.tr("Cancel")).min_size(vec2(80.0, 24.0))).clicked() {
                                answer_input = Some(None);
                            }
                        });
                    });
                }
                None => {
                    ui.spinner();
                }
            }
            ui.add_space(16.0);
            if ui.button(RichText::new(lang.tr("Stop macro")).color(Color32::from_rgb(0xC4, 0x2B, 0x1C))).clicked() {
                run.cancel.store(true, Ordering::Relaxed);
                // A waiting question is answered with Cancel so the macro can see the stop.
                match run.ask.take() {
                    Some(Ask::Msg { reply, .. }) => {
                        let _ = reply.send(2);
                    }
                    Some(Ask::Input { reply, .. }) => {
                        let _ = reply.send(None);
                    }
                    None => {}
                }
            }
        });
    });
    if let Some(code) = answer_msg
        && let Some(Ask::Msg { reply, .. }) = run.ask.take()
    {
        let _ = reply.send(code);
    }
    if let Some(v) = answer_input
        && let Some(Ask::Input { reply, .. }) = run.ask.take()
    {
        let _ = reply.send(v);
    }
    ui.ctx().request_repaint_after(std::time::Duration::from_millis(30));
    true
}

/// The runnable macros of the active workbook (`Module1.Main`…).
pub fn list(app: &SheetApp) -> Vec<String> {
    app.session.active().and_then(|d| d.wb.vba_project.as_ref()).and_then(|b| gridcraft_engine::vba_host::macro_names(b).ok()).unwrap_or_default()
}

/// Enable Content: the user trusts the workbook's macros. Runs `Workbook_Open` / `Auto_Open`.
pub fn enable(app: &mut SheetApp) {
    let Some(uid) = app.session.active().map(|d| d.uid) else { return };
    if !app.ext.macros_enabled.contains(&uid) {
        app.ext.macros_enabled.push(uid);
    }
    app.notice = None;
    let names = app
        .session
        .active()
        .and_then(|d| d.wb.vba_project.as_ref())
        .and_then(|b| gridcraft_engine::vba_host::all_procedures(b).ok())
        .unwrap_or_default();
    if let Some(open) = names
        .iter()
        .find(|n| n.eq_ignore_ascii_case("ThisWorkbook.Workbook_Open"))
        .or_else(|| names.iter().find(|n| n.to_ascii_lowercase().ends_with(".auto_open")))
    {
        start(app, &open.clone());
    }
}

/// The Macros dialog's body: pick a macro and run it.
pub fn dialog(app: &mut SheetApp, ui: &mut Ui, d: &mut crate::dialogs::Dialog, open: &mut bool) {
    let lang = app.ui.language;
    let names = list(app);
    if names.is_empty() {
        ui.weak(lang.tr("This workbook has no macros."));
    }
    let mut selected = d.values.get("selected").and_then(Json::as_str).map(str::to_string).or_else(|| names.first().cloned()).unwrap_or_default();
    egui::ScrollArea::vertical().max_height(260.0).show(ui, |ui| {
        for n in &names {
            let r = ui.add(egui::Button::selectable(*n == selected, n).min_size(vec2(440.0, 22.0)));
            if r.clicked() {
                selected = n.clone();
            }
            if r.double_clicked() {
                *open = false;
                start(app, n);
            }
        }
    });
    d.values.insert("selected".into(), json!(selected));
    // Scripts recorded in GridCraft (Automate › Record Actions) are listed too.
    let scripts = app.session.run("automate.allScripts", json!({})).ok().and_then(|v| v.as_array().cloned()).unwrap_or_default();
    let scripts: Vec<String> =
        scripts.iter().filter(|s| s["kind"].as_str() != Some("vba")).filter_map(|s| s["name"].as_str().map(str::to_string)).collect();
    if !scripts.is_empty() {
        ui.separator();
        ui.label(RichText::new(lang.tr("GridCraft scripts")).strong());
        for s in scripts {
            if ui.button(&s).clicked() {
                *open = false;
                app.run_or_alert("automate.runScript", json!({"name": s}));
            }
        }
    }
    ui.add_space(8.0);
    ui.horizontal(|ui| {
        if ui.add_enabled(!selected.is_empty(), egui::Button::new(lang.tr("Run"))).clicked() {
            *open = false;
            start(app, &selected);
        }
        if ui.button(lang.tr("Cancel")).clicked() {
            *open = false;
        }
    });
}

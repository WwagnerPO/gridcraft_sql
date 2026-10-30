//! Running VBA macros from the window: the Macros dialog, a MsgBox answered by clicking, Stop,
//! and Enable Content running Workbook_Open.

use egui::vec2;
use egui_kittest::kittest::Queryable;
use gridcraft_engine::Session;
use gridcraft_engine::core::{CellRef, Value};
use gridcraft_ui_egui::SheetApp;
use serde_json::json;

type Harness = egui_kittest::Harness<'static, SheetApp>;

fn harness(modules: &[(&str, bool, &str)]) -> Harness {
    let mut session = Session::new();
    session.new_workbook();
    let bin = gridcraft_xlsx::ovba::tests_support::sample_project(modules);
    std::sync::Arc::make_mut(&mut session.doc_mut().unwrap().wb).vba_project = Some(std::sync::Arc::new(bin));
    // As a macro-enabled workbook arrives: from a file, macros disabled until enabled.
    session.doc_mut().unwrap().trust = gridcraft_engine::Trust::FromFile;
    let b = egui_kittest::Harness::builder().with_step_dt(1.0 / 60.0).with_size(vec2(1000.0, 600.0));
    let b = if std::env::var_os("GRIDCRAFT_SNAPSHOT_DIR").is_some() { b.wgpu() } else { b };
    let mut h = b.build_ui_state(
        |ui, app: &mut SheetApp| {
            app.logic(ui.ctx());
            app.ui(ui);
        },
        SheetApp::new(session, Default::default()),
    );
    h.run_steps(4);
    h
}

/// Enable Content on the security bar.
fn enable(h: &mut Harness) {
    h.state_mut().open_dialog("enableMacros", json!({}));
    h.run_steps(2);
}

/// Frames until the running macro asks something or ends.
fn wait(h: &mut Harness) {
    for _ in 0..400 {
        h.run_steps(1);
        std::thread::sleep(std::time::Duration::from_millis(2));
        if h.state().ext.macro_run.is_none() {
            return;
        }
        if h.query_by_label("Yes").is_some() || h.query_by_label("OK").is_some() {
            return;
        }
    }
}

fn cell(h: &Harness, a1: &str) -> Value {
    h.state().session.active().unwrap().wb.active().unwrap().value(CellRef::parse(a1).unwrap())
}

#[test]
fn run_from_the_dialog_and_answer_a_message_box() {
    let mut h = harness(&[(
        "Module1",
        true,
        "Attribute VB_Name = \"Module1\"\r\nSub Saludo()\r\n    If MsgBox(\"¿Escribir?\", vbYesNo) = vbYes Then Range(\"A1\").Value = \"sí\" Else Range(\"A1\").Value = \"no\"\r\n    Range(\"A2\").Value = 42\r\nEnd Sub\r\n",
    )]);
    enable(&mut h);
    h.state_mut().open_dialog("macros", json!({}));
    h.run_steps(4);
    assert!(h.query_by_label("Module1.Saludo").is_some(), "the macro is listed");
    h.get_by_label("Run").click();
    h.run_steps(1);
    wait(&mut h);
    assert!(h.state().ext.macro_run.is_some(), "the macro waits on its message box");
    if let Some(dir) = std::env::var_os("GRIDCRAFT_SNAPSHOT_DIR") {
        match h.render() {
            Ok(img) => {
                let _ = img.save(std::path::Path::new(&dir).join("macro-msgbox.png"));
            }
            Err(e) => eprintln!("render: {e}"),
        }
    }
    h.get_by_label("Yes").click();
    for _ in 0..400 {
        h.run_steps(1);
        std::thread::sleep(std::time::Duration::from_millis(2));
        if h.state().ext.macro_run.is_none() {
            break;
        }
    }
    assert!(h.state().ext.macro_run.is_none(), "the macro finished and gave the workbook back");
    assert_eq!(cell(&h, "A1"), Value::text("sí"));
    assert_eq!(cell(&h, "A2"), Value::Number(42.0));
    // One undo step undoes the macro.
    h.state_mut().run("edit.undo", json!({})).unwrap();
    assert_eq!(cell(&h, "A2"), Value::Empty);
}

#[test]
fn stop_interrupts_a_runaway_macro() {
    let mut h = harness(&[("Module1", true, "Sub Forever()\r\n    Do\r\n        Range(\"A1\").Value = 1\r\n    Loop\r\nEnd Sub\r\n")]);
    enable(&mut h);
    gridcraft_ui_egui::macros::start(h.state_mut(), "Forever");
    h.run_steps(3);
    assert!(h.state().ext.macro_run.is_some());
    h.get_by_label("Stop macro").click();
    for _ in 0..400 {
        h.run_steps(1);
        std::thread::sleep(std::time::Duration::from_millis(2));
        if h.state().ext.macro_run.is_none() {
            break;
        }
    }
    assert!(h.state().ext.macro_run.is_none());
    let msg = h.state().message.clone().map(|m| m.1).unwrap_or_default();
    assert!(msg.contains("interrupted"), "{msg}");
    assert!(h.state().session.active().is_some(), "the workbook came back");
}

#[test]
fn enable_content_runs_workbook_open() {
    let mut h = harness(&[
        (
            "ThisWorkbook",
            false,
            "Attribute VB_Name = \"ThisWorkbook\"\r\nPrivate Sub Workbook_Open()\r\n    Range(\"C3\").Value = \"abierto\"\r\nEnd Sub\r\n",
        ),
        ("Module1", true, "Sub A()\r\nEnd Sub\r\n"),
    ]);
    h.state_mut().open_dialog("enableMacros", json!({}));
    for _ in 0..400 {
        h.run_steps(1);
        std::thread::sleep(std::time::Duration::from_millis(2));
        if h.state().ext.macro_run.is_none() {
            break;
        }
    }
    assert_eq!(cell(&h, "C3"), Value::text("abierto"));
}

#[test]
fn the_open_message_leaves_macros_to_the_security_bar() {
    let mut h = harness(&[("Module1", true, "Sub A()\r\nEnd Sub\r\n")]);
    h.state_mut().message = Some(("GridCraft".into(), "sheet \"X\": controls are kept\n\nmacros (VBA project) were kept but not executed".into()));
    h.state_mut().notice = Some(gridcraft_ui_egui::Notice {
        text: "This workbook contains VBA macros.".into(),
        kind: gridcraft_ui_egui::NoticeKind::Warning,
        action: Some(("What this means".into(), "macros".into())),
    });
    h.run_steps(2);
    let msg = h.state().message.clone().map(|m| m.1).unwrap_or_default();
    assert_eq!(msg, "sheet \"X\": controls are kept");
    assert_eq!(h.state().notice.as_ref().and_then(|n| n.action.clone()).map(|a| a.1), Some("enableMacros".into()));
    assert!(h.query_by_label("Macros disabled").is_some(), "the status bar says macros are off");
    // A message that only had the macro line goes away.
    h.state_mut().message = Some(("GridCraft".into(), "macros (VBA project) were kept but not executed".into()));
    h.run_steps(2);
    assert!(h.state().message.is_none());
    h.state_mut().open_dialog("enableMacros", json!({}));
    h.run_steps(3);
    assert!(h.query_by_label("Macros enabled").is_some());
}

/// A sheet with a form button running `Mark`: click it on screen.
fn button_harness() -> Harness {
    let mut h = harness(&[("Module1", true, "Sub Mark()\r\n    Range(\"A1\").Value = Application.Caller\r\nEnd Sub\r\n")]);
    let p = gridcraft_xlsx::form_controls::sample_preserved(&[(1025, "Button 1", "Run me", "Mark", [2, 0, 3, 0, 4, 0, 5, 0])]);
    {
        let s = &mut h.state_mut().session;
        let wb = std::sync::Arc::make_mut(&mut s.doc_mut().unwrap().wb);
        wb.sheet_mut(0).unwrap().preserved = Some(std::sync::Arc::new(p));
    }
    h.run_steps(3);
    h
}

fn click_button(h: &mut Harness) {
    let app = h.state();
    let sh = app.session.active().unwrap().wb.active().unwrap();
    let geo = gridcraft_ui_egui::grid::Geo::new(sh, app.grid.rect.unwrap(), app.view().scroll);
    // Inside the button (C4:E6), clear of its edges.
    let pos = geo.cell_rect(sh, CellRef::parse("D5").unwrap()).center();
    h.input_mut().events.push(egui::Event::PointerMoved(pos));
    for pressed in [true, false] {
        h.input_mut().events.push(egui::Event::PointerButton {
            pos,
            button: egui::PointerButton::Primary,
            pressed,
            modifiers: egui::Modifiers::NONE,
        });
        h.run_steps(1);
    }
}

#[test]
fn clicking_a_form_button_runs_its_macro() {
    let mut h = button_harness();
    // Macros disabled: the click explains instead of running.
    click_button(&mut h);
    h.run_steps(2);
    assert!(h.state().message.as_ref().is_some_and(|m| m.1.contains("Enable Content")));
    assert_eq!(cell(&h, "A1"), Value::Empty);
    h.state_mut().message = None;
    h.state_mut().open_dialog("enableMacros", json!({}));
    h.run_steps(2);
    click_button(&mut h);
    for _ in 0..400 {
        h.run_steps(1);
        std::thread::sleep(std::time::Duration::from_millis(2));
        if h.state().ext.macro_run.is_none() && cell(&h, "A1") != Value::Empty {
            break;
        }
    }
    assert_eq!(cell(&h, "A1"), Value::text("Button 1"), "Application.Caller is the button");
}

#[test]
fn editing_a_cell_runs_worksheet_change_with_a_real_message_box() {
    let mut h = harness(&[
        ("Module1", true, "Attribute VB_Name = \"Module1\"\r\n"),
        (
            "Hoja1",
            false,
            "Attribute VB_Name = \"Hoja1\"\r\nPrivate Sub Worksheet_Change(ByVal Target As Range)\r\n    If MsgBox(\"¿Guardar \" & Target.Address & \"?\", vbYesNo) = vbYes Then Range(\"Z1\").Value = Target.Value\r\nEnd Sub\r\n",
        ),
    ]);
    {
        let s = &mut h.state_mut().session;
        let wb = std::sync::Arc::make_mut(&mut s.doc_mut().unwrap().wb);
        std::sync::Arc::make_mut(&mut wb.sheets[0]).code_name = Some("Hoja1".into());
    }
    // Disabled: editing raises nothing.
    h.state_mut().run("cell.set", json!({"cell": "B2", "input": "7"})).unwrap();
    h.run_steps(3);
    assert!(h.state().ext.macro_run.is_none());
    enable(&mut h);
    h.state_mut().run("cell.set", json!({"cell": "B2", "input": "8"})).unwrap();
    h.run_steps(2);
    wait(&mut h);
    h.run_steps(2);
    assert!(h.query_by_label_contains("$B$2").is_some(), "the handler's message box shows");
    h.get_by_label("Yes").click();
    for _ in 0..400 {
        h.run_steps(1);
        std::thread::sleep(std::time::Duration::from_millis(2));
        if h.state().ext.macro_run.is_none() {
            break;
        }
    }
    assert_eq!(cell(&h, "Z1"), Value::Number(8.0));
    assert!(h.state().message.is_none(), "{:?}", h.state().message);
}

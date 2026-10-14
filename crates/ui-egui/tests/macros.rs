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
    std::sync::Arc::make_mut(&mut session.doc_mut().unwrap().wb).vba_project = Some(bin);
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

/// Frames until the running macro asks something or ends.
fn wait(h: &mut Harness) {
    for _ in 0..400 {
        h.run_steps(1);
        std::thread::sleep(std::time::Duration::from_millis(2));
        if h.state().macro_run.is_none() {
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
    h.state_mut().open_dialog("macros", json!({}));
    h.run_steps(4);
    assert!(h.query_by_label("Module1.Saludo").is_some(), "the macro is listed");
    h.get_by_label("Run").click();
    h.run_steps(1);
    wait(&mut h);
    assert!(h.state().macro_run.is_some(), "the macro waits on its message box");
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
        if h.state().macro_run.is_none() {
            break;
        }
    }
    assert!(h.state().macro_run.is_none(), "the macro finished and gave the workbook back");
    assert_eq!(cell(&h, "A1"), Value::text("sí"));
    assert_eq!(cell(&h, "A2"), Value::Number(42.0));
    // One undo step undoes the macro.
    h.state_mut().run("edit.undo", json!({})).unwrap();
    assert_eq!(cell(&h, "A2"), Value::Empty);
}

#[test]
fn stop_interrupts_a_runaway_macro() {
    let mut h = harness(&[("Module1", true, "Sub Forever()\r\n    Do\r\n        Range(\"A1\").Value = 1\r\n    Loop\r\nEnd Sub\r\n")]);
    gridcraft_ui_egui::macros::start(h.state_mut(), "Forever");
    h.run_steps(3);
    assert!(h.state().macro_run.is_some());
    h.get_by_label("Stop macro").click();
    for _ in 0..400 {
        h.run_steps(1);
        std::thread::sleep(std::time::Duration::from_millis(2));
        if h.state().macro_run.is_none() {
            break;
        }
    }
    assert!(h.state().macro_run.is_none());
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
        if h.state().macro_run.is_none() {
            break;
        }
    }
    assert_eq!(cell(&h, "C3"), Value::text("abierto"));
}

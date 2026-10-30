//! Macros (Developer tab): the VBA project kept in a macro-enabled workbook.

use serde_json::{Value as Json, json};

use super::{CommandSpec, EngineError, Result, Session, bad, bool_param, cmd, has_doc, str_param, target_sheet};

pub fn specs() -> Vec<CommandSpec> {
    vec![
        cmd!(
            noundo "vba.run",
            "Run Macro",
            ["Developer", "Code", "Macros"],
            None,
            "{macro: \"Module1.Main\" | \"Main\", answers?: [msgbox button codes], inputs?: [text], trusted?: true (enables the workbook's macros first, see vba.trust)} runs a VBA macro of the workbook; its changes undo as one step → {result, messages, output}",
            has_doc,
            run
        ),
        cmd!(
            noundo "vba.trust",
            "Enable Content",
            [],
            None,
            "{trusted?: true | false} enables (or disables again) the macros of the active workbook for this session; a workbook opened from a file starts with them disabled → {trusted}",
            has_doc,
            trust
        ),
        cmd!(
            query "vba.buttons",
            "Buttons",
            [],
            None,
            "{sheet?} → [{name, caption, macro, range}] the sheet's form-control buttons",
            has_doc,
            buttons
        ),
        cmd!(
            noundo "vba.click",
            "Click Button",
            [],
            None,
            "{button: \"Button 1\", sheet?, answers?, inputs?, trusted?} runs the button's macro (Application.Caller is the button) → {result, messages, output}",
            has_doc,
            click
        ),
        cmd!(
            query "vba.macros",
            "Macros",
            ["Developer", "Code"],
            Some("Alt+F8"),
            "{} → [\"Module1.Main\", …] the macros you can run (public Subs without required arguments)",
            has_doc,
            macros
        ),
        cmd!(
            query "vba.modules",
            "VBA Modules",
            [],
            None,
            "{source?: true} → [{name, kind: standard|class|document|form, lines, source?}] the workbook's VBA modules",
            has_doc,
            modules
        ),
    ]
}

fn sheet_buttons(s: &Session, p: &Json) -> Result<(usize, Vec<gridcraft_xlsx::form_controls::FormButton>)> {
    let si = target_sheet(s, p)?;
    let d = s.doc()?;
    let list = d.wb.sheet(si).and_then(|sh| sh.preserved.as_deref()).map(gridcraft_xlsx::form_controls::buttons).unwrap_or_default();
    Ok((si, list))
}

fn buttons(s: &mut Session, p: &Json) -> Result<Json> {
    let (_, list) = sheet_buttons(s, p)?;
    Ok(Json::Array(
        list.iter()
            .map(|b| {
                let [c0, _, r0, _, c1, _, r1, _] = b.anchor;
                let range = gridcraft_core::RangeRef::new(gridcraft_core::CellRef::new(r0, c0), gridcraft_core::CellRef::new(r1, c1));
                json!({"name": b.name, "caption": b.caption, "macro": b.macro_name, "range": range.a1()})
            })
            .collect(),
    ))
}

fn ui_from(p: &Json) -> crate::vba_host::Headless {
    let mut ui = crate::vba_host::Headless::default();
    if let Some(a) = p.get("answers").and_then(Json::as_array) {
        ui.answers = a.iter().filter_map(Json::as_i64).collect();
    }
    if let Some(a) = p.get("inputs").and_then(Json::as_array) {
        ui.inputs = a.iter().map(|v| v.as_str().map(str::to_string)).collect();
    }
    ui
}

fn click(s: &mut Session, p: &Json) -> Result<Json> {
    trust_from_call(s, p)?;
    let name = str_param(p, "button").ok_or_else(|| bad("vba.click", "missing `button`"))?.to_string();
    let (si, list) = sheet_buttons(s, p)?;
    let b = list.iter().find(|b| b.name.eq_ignore_ascii_case(&name)).ok_or_else(|| bad("vba.click", format!("no button named `{name}`")))?;
    if b.macro_name.is_empty() {
        return Err(EngineError::Other(format!("No macro is assigned to {name}.")));
    }
    let macro_name = b.macro_name.clone();
    let caller = b.name.clone();
    // The button's sheet is active while its macro runs, as when it is clicked.
    s.doc_mut()?.wb_switch_sheet(si);
    let mut ui = ui_from(p);
    let r = crate::vba_host::run_macro_from(s, &macro_name, &mut ui, None, Some(caller)).map_err(|e| EngineError::Other(e.to_string()))?;
    Ok(json!({"result": r.to_str().unwrap_or_default(), "messages": ui.messages, "output": ui.output}))
}

fn macros(s: &mut Session, _: &Json) -> Result<Json> {
    let Some(proj) = project(s)? else { return Ok(json!([])) };
    let sources: Vec<(String, String)> = proj.modules.iter().map(|m| (m.name.clone(), m.source.clone())).collect();
    let program = gridcraft_vba::Program::parse(&sources).map_err(|e| EngineError::Other(e.to_string()))?;
    let mut ui = crate::vba_host::Headless::default();
    let mut host = crate::vba_host::ExcelHost::new(s, &mut ui);
    let names = gridcraft_vba::Interp::new(program, &mut host).macros();
    Ok(json!(names))
}

fn trust(s: &mut Session, p: &Json) -> Result<Json> {
    // Only a missing `trusted` means "trust"; `"false"`, `0` or `null` are mistakes, not consent.
    let trusted = match p.get("trusted") {
        None => true,
        Some(Json::Bool(b)) => *b,
        Some(_) => return Err(bad("vba.trust", "`trusted` must be true or false")),
    };
    let uid = s.doc()?.uid;
    crate::vba_host::trust_macros(s, uid, trusted);
    Ok(json!({"trusted": trusted}))
}

/// `trusted: true` on a run call enables the workbook's macros first, as `vba.trust` does.
fn trust_from_call(s: &mut Session, p: &Json) -> Result<()> {
    if p.get("trusted") == Some(&Json::Bool(true)) {
        let uid = s.doc()?.uid;
        crate::vba_host::trust_macros(s, uid, true);
    }
    Ok(())
}

fn run(s: &mut Session, p: &Json) -> Result<Json> {
    trust_from_call(s, p)?;
    let name = str_param(p, "macro").or_else(|| str_param(p, "name")).ok_or_else(|| bad("vba.run", "missing `macro`"))?.to_string();
    let mut ui = crate::vba_host::Headless::default();
    if let Some(a) = p.get("answers").and_then(Json::as_array) {
        ui.answers = a.iter().filter_map(Json::as_i64).collect();
    }
    if let Some(a) = p.get("inputs").and_then(Json::as_array) {
        ui.inputs = a.iter().map(|v| v.as_str().map(str::to_string)).collect();
    }
    let r = crate::vba_host::run_macro(s, &name, &mut ui).map_err(|e| EngineError::Other(e.to_string()))?;
    Ok(json!({"result": r.to_str().unwrap_or_default(), "messages": ui.messages, "output": ui.output}))
}

pub(crate) fn kind_name(k: gridcraft_xlsx::ovba::ModuleKind) -> &'static str {
    use gridcraft_xlsx::ovba::ModuleKind as K;
    match k {
        K::Standard => "standard",
        K::Class => "class",
        K::Document => "document",
        K::Form => "form",
    }
}

/// The active workbook's VBA project, read from its kept `vbaProject.bin`.
pub(crate) fn project(s: &Session) -> Result<Option<gridcraft_xlsx::ovba::VbaProject>> {
    let Some(bin) = s.doc()?.wb.vba_project.as_ref() else { return Ok(None) };
    gridcraft_xlsx::ovba::read_project(bin).map(Some).map_err(EngineError::Other)
}

fn modules(s: &mut Session, p: &Json) -> Result<Json> {
    let with_source = bool_param(p, "source").unwrap_or(false);
    let Some(proj) = project(s)? else { return Ok(json!([])) };
    Ok(Json::Array(
        proj.modules
            .iter()
            .map(|m| {
                let mut v = json!({"name": m.name, "kind": kind_name(m.kind), "lines": m.source.lines().count()});
                if with_source {
                    v["source"] = json!(m.source);
                }
                v
            })
            .collect(),
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lists_the_modules_of_a_macro_workbook() {
        let mut s = Session::new();
        s.execute("file.new", json!({})).unwrap();
        assert_eq!(s.execute("vba.modules", json!({})).unwrap(), json!([]));
        let bin =
            gridcraft_xlsx::ovba::tests_support::sample_project(&[("Module1", true, "Attribute VB_Name = \"Module1\"\r\nSub A()\r\nEnd Sub\r\n")]);
        Arc_mut(&mut s).vba_project = Some(std::sync::Arc::new(bin));
        let r = s.execute("vba.modules", json!({"source": true})).unwrap();
        assert_eq!(r[0]["name"], "Module1");
        assert_eq!(r[0]["kind"], "standard");
        assert!(r[0]["source"].as_str().unwrap().contains("Sub A()"));
    }

    #[allow(non_snake_case)]
    fn Arc_mut(s: &mut Session) -> &mut gridcraft_model::Workbook {
        std::sync::Arc::make_mut(&mut s.doc_mut().unwrap().wb)
    }
}

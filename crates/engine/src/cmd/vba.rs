//! Macros (Developer tab): the VBA project kept in a macro-enabled workbook.

use serde_json::{Value as Json, json};

use super::*;

pub fn specs() -> Vec<CommandSpec> {
    vec![
        cmd!(
            noundo "vba.run",
            "Run Macro",
            ["Developer", "Code", "Macros"],
            None,
            "{macro: \"Module1.Main\" | \"Main\", answers?: [msgbox button codes], inputs?: [text]} runs a VBA macro of the workbook; its changes undo as one step → {result, messages, output}",
            has_doc,
            run
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

fn macros(s: &mut Session, _: &Json) -> Result<Json> {
    let Some(proj) = project(s)? else { return Ok(json!([])) };
    let sources: Vec<(String, String)> = proj.modules.iter().map(|m| (m.name.clone(), m.source.clone())).collect();
    let program = gridcraft_vba::Program::parse(&sources).map_err(|e| EngineError::Other(e.to_string()))?;
    let mut ui = crate::vba_host::Headless::default();
    let mut host = crate::vba_host::ExcelHost::new(s, &mut ui);
    let names = gridcraft_vba::Interp::new(program, &mut host).macros();
    Ok(json!(names))
}

fn run(s: &mut Session, p: &Json) -> Result<Json> {
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
        Arc_mut(&mut s).vba_project = Some(bin);
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

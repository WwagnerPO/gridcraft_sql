//! The features this build adds to GridCraft — database connections (`get_data`) and VBA macros
//! (`macros`) — meet the shared app code here. `lib.rs` calls each hook with one line, so updating
//! from the upstream project touches as few shared lines as possible.

use serde_json::Value as Json;

use crate::{SheetApp, get_data, macros};

/// What the app keeps for these features.
#[derive(Default)]
pub struct State {
    /// When each query with a refresh interval last refreshed: (document uid, query name, ms).
    pub query_timers: Vec<(u64, String, f64)>,
    /// A VBA macro running on its worker (it holds the session meanwhile).
    pub macro_run: Option<macros::MacroRun>,
    /// Workbooks (by uid) whose macros the user enabled.
    pub macros_enabled: Vec<u64>,
}

/// Menu-style commands handled here; true when `id` was.
pub fn run_or_alert(app: &mut SheetApp, id: &str, params: &Json) -> bool {
    // A macro run from the interface gets real message boxes.
    if id == "vba.run"
        && let Some(name) = params.get("macro").or_else(|| params.get("name")).and_then(Json::as_str)
    {
        let name = name.to_string();
        macros::start(app, &name);
        return true;
    }
    // Refresh asks for a database login when the server has none, then refreshes.
    if matches!(id, "data.refresh" | "data.refreshAll") {
        get_data::refresh_with_login(app, id, params.clone());
        return true;
    }
    false
}

/// Interface actions (`open_dialog` names that aren't dialogs) handled here; true when `name` was.
pub fn ui_action(app: &mut SheetApp, name: &str) -> bool {
    match name {
        "enableMacros" => macros::enable(app),
        _ => return false,
    }
    true
}

/// Form-control buttons over the sheet's cells.
pub fn paint_controls(app: &SheetApp, painter: &egui::Painter, geo: &crate::grid::Geo, sh: &gridcraft_engine::model::Sheet) {
    crate::controls::paint(app, painter, geo, sh);
}

/// A click on the sheet: true when a button took it.
pub fn click_control(app: &mut SheetApp, geo: &crate::grid::Geo, sh: &gridcraft_engine::model::Sheet, p: egui::Pos2) -> bool {
    crate::controls::click(app, geo, sh, p)
}

/// After a workbook opened from a file.
pub fn after_open(app: &mut SheetApp) {
    get_data::refresh_on_open(app);
}

/// While a macro runs its worker owns the session: nothing may touch it.
pub fn busy(app: &SheetApp) -> bool {
    app.ext.macro_run.is_some()
}

/// Start of a frame. True when a running macro drew the frame (nothing else may).
pub fn frame(app: &mut SheetApp, ui: &mut egui::Ui) -> bool {
    if macros::frame(app, ui) {
        return true;
    }
    macro_notice(app);
    open_message(app);
    get_data::tick(app, &ui.ctx().clone());
    false
}

/// The import warning the engine words for a build that can't run macros.
const KEPT_NOT_RUN: &str = "macros (VBA project) were kept but not executed";

/// The open-time message box lists every import warning; the macro one is the security bar's
/// (with Enable Content), so it leaves the box, and an empty box isn't shown.
fn open_message(app: &mut SheetApp) {
    let Some((_, text)) = app.message.as_ref() else { return };
    if !text.contains(KEPT_NOT_RUN) {
        return;
    }
    let rest = text.split("\n\n").filter(|p| !p.contains(KEPT_NOT_RUN)).collect::<Vec<_>>().join("\n\n");
    if rest.is_empty() {
        app.message = None;
    } else if let Some(m) = app.message.as_mut() {
        m.1 = rest;
    }
}

/// The status bar's macro state for workbook `uid`: its label and whether macros are on.
pub fn macro_status(app: &SheetApp, uid: u64) -> (&'static str, bool) {
    let on = app.ext.macros_enabled.contains(&uid);
    (app.ui.language.tr(if on { "Macros enabled" } else { "Macros disabled" }), on)
}

/// The open-time "contains VBA macros" notice offers to run them (Enable Content).
fn macro_notice(app: &mut SheetApp) {
    if let Some(n) = app.notice.as_mut()
        && n.action.as_ref().is_some_and(|(_, a)| a == "macros")
    {
        n.text = "SECURITY WARNING Macros have been disabled. Enable them only if you trust where this workbook came from.".into();
        n.action = Some(("Enable Content".into(), "enableMacros".into()));
    }
}

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
    get_data::tick(app, &ui.ctx().clone());
    false
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

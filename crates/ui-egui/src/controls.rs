//! Form buttons whose macro is a VBA macro (upstream draws and presses form controls; this runs
//! the VBA ones on the macro worker). With the workbook's macros disabled it says how to enable
//! them instead.

use gridcraft_engine::model::{ControlAction, ControlKind};
use serde_json::Value as Json;

use crate::SheetApp;

/// Presses button `params.id` of the active sheet when its macro is a VBA macro. True when it was.
pub fn press(app: &mut SheetApp, params: &Json) -> bool {
    let Some(id) = params.get("id").and_then(Json::as_u64) else { return false };
    let Some(doc) = app.session.active() else { return false };
    let uid = doc.uid;
    let Some(c) = doc.wb.active().and_then(|sh| sh.controls.iter().find(|c| u64::from(c.id) == id && c.kind == ControlKind::Button)) else {
        return false;
    };
    let Some(ControlAction::Macro { name: mac, .. }) = &c.action else { return false };
    let (mac, button) = (mac.clone(), c.display_name());
    if !gridcraft_engine::vba_host::is_vba_button_macro(&app.session, &mac) {
        return false;
    }
    if !gridcraft_engine::vba_host::macros_allowed(&app.session, uid) {
        let lang = app.ui.language;
        app.message = Some((
            "GridCraft".into(),
            lang.tr("Macros are disabled in this workbook. Click Enable Content in the security bar to run the button's macro.").to_string(),
        ));
        return true;
    }
    crate::macros::start_from(app, &mac, Some(button));
    true
}

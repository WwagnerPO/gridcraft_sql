//! Form-control buttons drawn over the sheet. Clicking one runs its macro with
//! `Application.Caller` set to the button, as in Excel; with the workbook's macros disabled it
//! says how to enable them instead.

use egui::{Color32, CornerRadius, Painter, Pos2, Rect, Stroke, StrokeKind, pos2};
use gridcraft_engine::model::Sheet;
use gridcraft_engine::vba_host::FormButton;

use crate::SheetApp;
use crate::grid::Geo;

/// The sheet's buttons and where they are on screen.
fn placed(geo: &Geo, sh: &Sheet) -> Vec<(FormButton, Rect)> {
    gridcraft_engine::vba_host::sheet_buttons(sh)
        .into_iter()
        .map(|b| {
            let [c0, dx0, r0, dy0, c1, dx1, r1, dy1] = b.anchor;
            // Offsets are pixels at 96 dpi: 0.75 points each.
            let px = 0.75 * geo.z;
            let a = geo.cell_rect(sh, gridcraft_engine::core::CellRef::new(r0, c0)).min;
            let z = geo.cell_rect(sh, gridcraft_engine::core::CellRef::new(r1, c1)).min;
            let rect = Rect::from_min_max(pos2(a.x + dx0 as f32 * px, a.y + dy0 as f32 * px), pos2(z.x + dx1 as f32 * px, z.y + dy1 as f32 * px));
            (b, rect)
        })
        .filter(|(_, r)| r.width() > 1.0 && r.height() > 1.0)
        .collect()
}

pub fn paint(app: &SheetApp, painter: &Painter, geo: &Geo, sh: &Sheet) {
    let list = placed(geo, sh);
    if list.is_empty() {
        return;
    }
    let hover = painter.ctx().pointer_hover_pos();
    for (b, rect) in list {
        if !painter.clip_rect().intersects(rect) {
            continue;
        }
        let hot = hover.is_some_and(|p| rect.contains(p)) && app.ext.macro_run.is_none();
        let fill = if hot { Color32::from_rgb(0xE5, 0xF1, 0xFB) } else { Color32::from_rgb(0xF0, 0xF0, 0xF0) };
        painter.rect(rect, CornerRadius::same(2), fill, Stroke::new(1.0, Color32::from_rgb(0xAD, 0xAD, 0xAD)), StrokeKind::Inside);
        let font = egui::FontId::proportional((11.0 * geo.z).clamp(6.0, 28.0));
        let galley = painter.layout(b.caption.clone(), font, Color32::BLACK, (rect.width() - 6.0).max(4.0));
        let at = rect.center() - galley.size() / 2.0;
        painter.with_clip_rect(rect.intersect(painter.clip_rect())).galley(at, galley, Color32::BLACK);
        if hot {
            painter.ctx().set_cursor_icon(egui::CursorIcon::PointingHand);
        }
    }
}

/// A click at `p`: runs the button there. True when a button took the click.
pub fn click(app: &mut SheetApp, geo: &Geo, sh: &Sheet, p: Pos2) -> bool {
    let Some((b, _)) = placed(geo, sh).into_iter().find(|(_, r)| r.contains(p)) else { return false };
    if b.macro_name.is_empty() {
        return true;
    }
    let uid = app.session.active().map(|d| d.uid);
    if uid.is_some_and(|u| !app.ext.macros_enabled.contains(&u)) {
        let lang = app.ui.language;
        app.message = Some((
            "GridCraft".into(),
            lang.tr("Macros are disabled in this workbook. Click Enable Content in the security bar to run the button's macro.").to_string(),
        ));
        return true;
    }
    crate::macros::start_from(app, &b.macro_name, Some(b.name));
    true
}

//! VBA events: `Worksheet_Change` in a sheet's code module and `Workbook_SheetChange` in
//! `ThisWorkbook` run when cells of an enabled macro workbook change.
//!
//! Every edit commits through `cmd::commit`, which reports the cells it changed here
//! ([`note_changed`]). When the outermost command ends, [`after_command`] turns those reports into
//! events for the handlers the project has. A host without a window (CLI, MCP) runs them at once,
//! answering message boxes as `vba.run` does; the desktop app defers them
//! ([`set_deferred`]) and runs each on its macro worker, where message boxes are real
//! ([`take_pending`], [`super::run_event_from`]). Changes a macro makes don't raise events (as
//! with Excel's `Application.EnableEvents = False`), so a handler that writes cells can't loop.

use std::cell::RefCell;

use gridcraft_core::{CellRef, RangeRef};

use super::Headless;
use crate::{DocState, Session};

/// One event to run: handler `proc` (`Hoja1.Worksheet_Change`) of document `uid`, with the cells
/// that changed on `sheet`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PendingEvent {
    pub uid: u64,
    pub proc: String,
    /// `Workbook_SheetChange` takes the sheet as its first argument.
    pub with_sheet: bool,
    pub sheet: String,
    pub target: RangeRef,
}

/// Most events kept waiting: past it the oldest are dropped.
const MAX_PENDING: usize = 1000;

thread_local! {
    /// Cells changed by the commands of the current outermost command: (document, sheet, range).
    static CHANGED: RefCell<Vec<(u64, usize, RangeRef)>> = const { RefCell::new(Vec::new()) };
}

/// What [`crate::cmd::commit`] changed in `d`, as one range per sheet. Cheap for a workbook
/// without macros: nothing is kept.
pub(crate) fn note_changed(d: &DocState, changed: &[(usize, CellRef)]) {
    if d.wb.vba_project.is_none() || changed.is_empty() {
        return;
    }
    let mut boxes: Vec<(usize, RangeRef)> = vec![];
    for &(si, c) in changed {
        match boxes.iter_mut().find(|(s, _)| *s == si) {
            Some((_, r)) => {
                r.start = CellRef::new(r.start.row.min(c.row), r.start.col.min(c.col));
                r.end = CellRef::new(r.end.row.max(c.row), r.end.col.max(c.col));
            }
            None => boxes.push((si, RangeRef::new(c, c))),
        }
    }
    CHANGED.with(|v| {
        let mut v = v.borrow_mut();
        if v.len() < MAX_PENDING {
            v.extend(boxes.into_iter().map(|(si, r)| (d.uid, si, r)));
        }
    });
}

/// The desktop app runs events itself, on its macro worker: [`after_command`] only queues them.
pub fn set_deferred(s: &mut Session, deferred: bool) {
    s.data.events_deferred = deferred;
}

/// The events waiting for the app to run them, oldest first.
pub fn take_pending(s: &mut Session) -> Vec<PendingEvent> {
    std::mem::take(&mut s.data.pending_events)
}

/// Called when an outermost command ends: the changes it made become events.
pub(crate) fn after_command(s: &mut Session) {
    let changed = CHANGED.with(|v| std::mem::take(&mut *v.borrow_mut()));
    if changed.is_empty() || s.data.in_macro {
        return;
    }
    let mut events = vec![];
    for (uid, si, target) in changed {
        events.extend(events_for(s, uid, si, target));
    }
    if events.is_empty() {
        return;
    }
    if s.data.events_deferred {
        s.data.pending_events.extend(events);
        let extra = s.data.pending_events.len().saturating_sub(MAX_PENDING);
        s.data.pending_events.drain(..extra);
        return;
    }
    for ev in events {
        let mut ui = Headless::default();
        // An event's error is the handler's: the command that raised it has already succeeded.
        let _ = super::run_event_from(s, &ev, &mut ui, None);
    }
}

/// The handlers a change of `target` on sheet `si` of document `uid` reaches.
fn events_for(s: &Session, uid: u64, si: usize, target: RangeRef) -> Vec<PendingEvent> {
    let Some(d) = s.documents().iter().find(|d| d.uid == uid) else { return vec![] };
    if !super::macros_allowed(s, uid) {
        return vec![];
    }
    let Some(sheet) = d.wb.sheet(si) else { return vec![] };
    let Some(bin) = d.wb.vba_project.as_ref() else { return vec![] };
    let Ok(procs) = super::all_procedures(bin) else { return vec![] };
    let has = |name: &str| procs.iter().find(|p| p.eq_ignore_ascii_case(name)).cloned();
    let mut out = vec![];
    if let Some(code) = sheet.code_name.as_deref()
        && let Some(proc) = has(&format!("{code}.Worksheet_Change"))
    {
        out.push(PendingEvent { uid, proc, with_sheet: false, sheet: sheet.name.clone(), target });
    }
    // The workbook's module is `ThisWorkbook` in English Excel, `EsteLibro` in Spanish…
    if let Some(proc) = procs.iter().find(|p| p.to_ascii_lowercase().ends_with(".workbook_sheetchange")).cloned() {
        out.push(PendingEvent { uid, proc, with_sheet: true, sheet: sheet.name.clone(), target });
    }
    out
}

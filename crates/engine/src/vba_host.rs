//! Excel's object model for VBA macros: `Application`, `Workbooks`, `Worksheets`, `Range`,
//! `Cells`, `Font`, `Interior`, `WorksheetFunction` and the `xl*`/`vb*` constants, implemented
//! over the [`Session`] for the interpreter in `gridcraft-vba`. Behaviour follows Microsoft's
//! public VBA reference for Excel; nothing from Office is used.
//!
//! Objects are handles ([`HostRef`]): a class and an id naming the document (by uid), the sheet
//! (by name) and the cells. Writes go through engine commands or [`crate::cmd`] edits, so they
//! honour sheet protection, recalculate, and undo as one step when the macro ends.

use std::collections::{HashMap, VecDeque};

use gridcraft_core::{CellError, CellRef, RangeRef};
use gridcraft_model::{Color, HAlign, PatternType, Underline, VAlign, Workbook};
use gridcraft_vba::error::err;
use gridcraft_vba::{Host, HostRef, Obj, Value as V, VbaError};
use serde_json::{Value as Json, json};

use crate::Session;

pub use gridcraft_xlsx::form_controls::FormButton;

/// The form-control buttons of a sheet (Developer › Insert › Button).
pub fn sheet_buttons(sh: &gridcraft_model::Sheet) -> Vec<FormButton> {
    sh.preserved.as_deref().map(gridcraft_xlsx::form_controls::buttons).unwrap_or_default()
}

/// What a macro shows the person running it. Macros run on a thread of their own, so it is
/// `Send`.
pub trait MacroUi: Send {
    /// A message box; returns the button pressed (`vbOK` = 1, `vbYes` = 6…).
    fn msgbox(&mut self, prompt: &str, buttons: i64, title: &str) -> i64;
    /// An input box; `None` when cancelled.
    fn inputbox(&mut self, prompt: &str, title: &str, default: &str) -> Option<String>;
    fn debug_print(&mut self, text: &str);
    /// `Application.StatusBar`; `None` gives the status bar back.
    fn status(&mut self, _text: Option<String>) {}
}

/// A UI that asks nobody: message boxes answer their first button (or the queued answers),
/// input boxes the queued inputs. It records what was shown (CLI, MCP, tests).
#[derive(Default, Debug)]
pub struct Headless {
    pub messages: Vec<String>,
    pub output: Vec<String>,
    pub answers: VecDeque<i64>,
    pub inputs: VecDeque<Option<String>>,
}

impl MacroUi for Headless {
    fn msgbox(&mut self, prompt: &str, buttons: i64, _title: &str) -> i64 {
        self.messages.push(prompt.to_string());
        self.answers.pop_front().unwrap_or(match buttons & 7 {
            0 | 1 => 1, // OK
            2 => 3,     // Abort
            3 | 4 => 6, // Yes
            _ => 4,     // Retry
        })
    }
    fn inputbox(&mut self, prompt: &str, _title: &str, default: &str) -> Option<String> {
        self.messages.push(prompt.to_string());
        self.inputs.pop_front().unwrap_or_else(|| Some(default.to_string()))
    }
    fn debug_print(&mut self, text: &str) {
        self.output.push(text.to_string());
    }
}

/// Separates the parts of a handle id.
const SEP: char = '\u{1f}';
/// Most cells one `Value` read or write touches (a whole-column write is capped to it).
const MAX_CELLS: u64 = 2_000_000;

/// A range handle: the document, sheet, what it counts by (`c` cells, `r` rows, `k` columns) and
/// its areas.
#[derive(Clone, Debug)]
struct Rng {
    uid: u64,
    sheet: String,
    kind: char,
    areas: Vec<RangeRef>,
}

impl Rng {
    fn id(&self) -> String {
        let areas: Vec<String> = self.areas.iter().map(|a| a.a1()).collect();
        format!("{}{SEP}{}{SEP}{}{SEP}{}", self.uid, self.sheet, self.kind, areas.join(","))
    }
    fn parse(id: &str) -> Option<Rng> {
        let mut p = id.split(SEP);
        let uid = p.next()?.parse().ok()?;
        let sheet = p.next()?.to_string();
        let kind = p.next()?.chars().next()?;
        let areas = p.next()?.split(',').filter(|x| !x.is_empty()).map(RangeRef::parse).collect::<Option<Vec<_>>>()?;
        Some(Rng { uid, sheet, kind, areas })
    }
    fn first(&self) -> RangeRef {
        self.areas.first().copied().unwrap_or_default()
    }
    fn with(&self, kind: char, areas: Vec<RangeRef>) -> Rng {
        Rng { uid: self.uid, sheet: self.sheet.clone(), kind, areas }
    }
    /// `'Sheet'!A1:B2,'Sheet'!D4` for commands.
    fn address(&self) -> String {
        let q = gridcraft_formula::quote_sheet(&self.sheet);
        self.areas.iter().map(|a| format!("{q}!{}", a.a1())).collect::<Vec<_>>().join(",")
    }
}

fn host(class: &str, id: impl Into<String>) -> V {
    V::Object(Obj::Host(HostRef { class: class.into(), id: id.into() }))
}

fn range_value(r: &Rng) -> V {
    host("Range", r.id())
}

fn engine_err(e: crate::EngineError) -> VbaError {
    err(1004, clean(&e.to_string()))
}

/// An engine error without its `invalid parameters for `x`:` prefix.
fn clean(e: &str) -> String {
    match e.split_once("`: ") {
        Some((_, rest)) => rest.to_string(),
        None => e.to_string(),
    }
}

fn arg(args: &[V], i: usize) -> Option<&V> {
    args.get(i).filter(|v| !matches!(v, V::Error(448)))
}

/// `#N/A` and friends as `CVErr` numbers.
fn error_number(e: CellError) -> i32 {
    match e {
        CellError::Null => 2000,
        CellError::Div0 => 2007,
        CellError::Ref => 2023,
        CellError::Name => 2029,
        CellError::Num => 2036,
        CellError::NA => 2042,
        _ => 2015,
    }
}

fn error_text(n: i32) -> &'static str {
    match n {
        2000 => "#NULL!",
        2007 => "#DIV/0!",
        2023 => "#REF!",
        2029 => "#NAME?",
        2036 => "#NUM!",
        2042 => "#N/A",
        _ => "#VALUE!",
    }
}

/// An Excel serial as a VBA date (they agree from 1 March 1900; before it Excel counts the
/// fictitious 29 February 1900).
fn serial_to_vba(n: f64) -> f64 {
    if (1.0..61.0).contains(&n) { n + 1.0 } else { n }
}

fn vba_to_serial(d: f64) -> f64 {
    if (2.0..61.0).contains(&d) { d - 1.0 } else { d }
}

/// The colour as VBA's `RGB()` long (red in the low byte).
fn color_to_long(c: Color) -> V {
    match c {
        Color::Rgb(x) => V::Long((((x >> 16) & 0xFF) | (x & 0xFF00) | ((x & 0xFF) << 16)) as i32),
        Color::Auto => V::Long(0),
        Color::Theme(..) => V::Long(0),
    }
}

fn long_to_color(v: &V) -> Result<Color, VbaError> {
    let n = v.to_i64()? as u32;
    let (r, g, b) = (n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF);
    Ok(Color::Rgb((r << 16) | (g << 8) | b))
}

/// `ColorIndex` from Excel's default 56-colour palette (the first entries; others map to black).
fn palette(i: i64) -> Option<Color> {
    const P: [u32; 16] = [
        0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080,
        0xC0C0C0, 0x808080,
    ];
    if i < 1 {
        return None;
    }
    Some(Color::Rgb(P.get((i - 1) as usize).copied().unwrap_or(0)))
}

/// The `xl*`, `vb*` and `msoTrue` style constants macros use.
fn constant(name: &str) -> Option<V> {
    let n = name.to_ascii_lowercase();
    let l = |x: i32| Some(V::Long(x));
    match n.as_str() {
        // VBA
        "vbcrlf" | "vbnewline" => Some(V::str("\r\n")),
        "vbcr" => Some(V::str("\r")),
        "vblf" => Some(V::str("\n")),
        "vbtab" => Some(V::str("\t")),
        "vbnullstring" => Some(V::str("")),
        "vbnullchar" => Some(V::str("\0")),
        "vbokonly" | "vbbinarycompare" | "vbusesystemdayofweek" | "vbempty" | "vbblack" => l(0),
        "vbokcancel" | "vbtextcompare" | "vbsunday" | "vbnull" => l(1),
        "vbabortretryignore" | "vbmonday" | "vbinteger" => l(2),
        "vbyesnocancel" | "vblong" => l(3),
        "vbyesno" => l(4),
        "vbretrycancel" => l(5),
        "vbok" => l(1),
        "vbcancel" => l(2),
        "vbabort" => l(3),
        "vbretry" => l(4),
        "vbignore" => l(5),
        "vbyes" => l(6),
        "vbno" => l(7),
        "vbcritical" => l(16),
        "vbquestion" => l(32),
        "vbexclamation" => l(48),
        "vbinformation" => l(64),
        "vbdefaultbutton2" => l(256),
        "vbdefaultbutton3" => l(512),
        "vbred" => l(0x0000FF),
        "vbgreen" => l(0x00FF00),
        "vbblue" => l(0xFF0000),
        "vbyellow" => l(0x00FFFF),
        "vbmagenta" => l(0xFF00FF),
        "vbcyan" => l(0xFFFF00),
        "vbwhite" => l(0xFFFFFF),
        "vbstring" => l(8),
        "vbdouble" => l(5),
        "vbdate" => l(7),
        "vbboolean" => l(11),
        "vbvariant" => l(12),
        "vbarray" => l(8192),
        // Excel
        "xlup" => l(-4162),
        "xldown" => l(-4121),
        "xltoleft" => l(-4159),
        "xltoright" => l(-4161),
        "xlshiftup" => l(-4162),
        "xlshiftdown" => l(-4121),
        "xlshifttoleft" => l(-4159),
        "xlshifttoright" => l(-4161),
        "xlcalculationmanual" => l(-4135),
        "xlcalculationautomatic" | "xlautomatic" => l(-4105),
        "xlcalculationsemiautomatic" => l(2),
        "xlpastevalues" | "xlvalues" => l(-4163),
        "xlpasteformats" => l(-4122),
        "xlpasteall" => l(-4104),
        "xlpasteformulas" | "xlformulas" | "xlcelltypeformulas" => l(-4123),
        "xlpastevaluesandnumberformats" => l(12),
        "xlpastecolumnwidths" => l(8),
        "xlcenter" => l(-4108),
        "xlleft" => l(-4131),
        "xlright" => l(-4152),
        "xltop" => l(-4160),
        "xlbottom" => l(-4107),
        "xlgeneral" => l(1),
        "xljustify" => l(-4130),
        "xlnone" | "xlunderlinestylenone" => l(-4142),
        "xlsolid" | "xlascending" | "xlyes" | "xlwhole" | "xlbyrows" | "xlnext" | "xla1" | "xland" | "xlcontinuous" | "xlpasteoperationnone" => l(1),
        "xldescending" | "xlno" | "xlpart" | "xlbycolumns" | "xlprevious" | "xlor" | "xlthin" | "xlcelltypeconstants" | "xlunderlinestylesingle" => {
            l(2)
        }
        "xlguess" => l(0),
        "xlcelltypelastcell" => l(11),
        "xlcelltypeblanks" => l(4),
        "xlcelltypevisible" => l(12),
        "xlmedium" => l(-4138),
        "xlthick" => l(4),
        "xlhairline" => l(1),
        "xldash" => l(-4115),
        "xldot" => l(-4118),
        "xldouble" => l(-4119),
        "xllinestylenone" => l(-4142),
        "xledgeleft" => l(7),
        "xledgetop" => l(8),
        "xledgebottom" => l(9),
        "xledgeright" => l(10),
        "xlinsidevertical" => l(11),
        "xlinsidehorizontal" => l(12),
        "xlsheetvisible" => l(-1),
        "xlsheethidden" => l(0),
        "xlsheetveryhidden" => l(2),
        "xlopenxmlworkbook" => l(51),
        "xlopenxmlworkbookmacroenabled" => l(52),
        "xlcsv" => l(6),
        "xlfiltervalues" => l(7),
        "xlr1c1" => l(-4150),
        "xlworksheet" => l(-4167),
        "xlmaximized" => l(-4137),
        "xlminimized" => l(-4140),
        "xlnormal" => l(-4143),
        "msotrue" => l(-1),
        "msofalse" => l(0),
        _ => None,
    }
}

pub struct ExcelHost<'a> {
    pub s: &'a mut Session,
    pub ui: &'a mut dyn MacroUi,
    /// The workbook holding the macro (`ThisWorkbook`).
    pub this_uid: u64,
    /// Application settings macros set (ScreenUpdating, DisplayAlerts…): kept, mostly no-ops.
    app: HashMap<String, V>,
    /// Sheets renamed during the macro: old name (lower) → new name.
    renamed: Vec<(String, String)>,
    /// `Application.Caller`: the button that started the macro.
    pub caller: Option<String>,
    streams: Vec<(u32, objects::Stream)>,
    next_stream: u32,
}

impl<'a> ExcelHost<'a> {
    pub fn new(s: &'a mut Session, ui: &'a mut dyn MacroUi) -> ExcelHost<'a> {
        let this_uid = s.active().map(|d| d.uid).unwrap_or(0);
        ExcelHost { s, ui, this_uid, app: HashMap::new(), renamed: vec![], caller: None, streams: vec![], next_stream: 0 }
    }

    // ------------------------------------------------------------ documents and sheets

    fn doc_index(&self, uid: u64) -> Result<usize, VbaError> {
        self.s.documents().iter().position(|d| d.uid == uid).ok_or_else(|| err(9, "Subscript out of range (the workbook is closed)"))
    }

    fn wb(&self, uid: u64) -> Result<&Workbook, VbaError> {
        Ok(&self.s.documents().get(self.doc_index(uid)?).ok_or_else(|| err(9, "Subscript out of range"))?.wb)
    }

    /// The current name of a sheet, following renames during the macro.
    fn sheet_name(&self, name: &str) -> String {
        let mut n = name.to_string();
        for _ in 0..16 {
            match self.renamed.iter().find(|(old, _)| *old == n.to_lowercase()) {
                Some((_, new)) => n = new.clone(),
                None => break,
            }
        }
        n
    }

    fn sheet_index(&self, uid: u64, name: &str) -> Result<usize, VbaError> {
        let name = self.sheet_name(name);
        self.wb(uid)?.sheet_index(&name).ok_or_else(|| err(9, format!("Subscript out of range (no sheet named {name})")))
    }

    fn active_uid(&self) -> u64 {
        self.s.active().map(|d| d.uid).unwrap_or(self.this_uid)
    }

    fn active_sheet_name(&self, uid: u64) -> Result<String, VbaError> {
        let wb = self.wb(uid)?;
        Ok(wb.active().map(|s| s.name.clone()).unwrap_or_default())
    }

    fn sheet_handle(&self, uid: u64, name: &str) -> V {
        host("Worksheet", format!("{uid}{SEP}{name}"))
    }

    fn parse_sheet(id: &str) -> Option<(u64, String)> {
        let (u, n) = id.split_once(SEP)?;
        Some((u.parse().ok()?, n.to_string()))
    }

    /// Runs an engine command on document `uid`.
    fn run(&mut self, uid: u64, id: &str, params: Json) -> Result<Json, VbaError> {
        let di = self.doc_index(uid)?;
        self.s.set_active(di);
        let r = self.s.execute(id, params).map_err(engine_err);
        // Dialogs a command queued have no one to answer them during a macro.
        self.s.ui_requests.retain(|q| !matches!(q, crate::UiRequest::Dialog(..)));
        r
    }

    // ------------------------------------------------------------ ranges

    /// A range from text on a sheet: `A1`, `A1:B2`, `A:A`, `1:3`, `A1,C3`, `'Other'!A1`, or a
    /// defined name.
    fn parse_range(&self, uid: u64, sheet: &str, text: &str) -> Result<Rng, VbaError> {
        let wb = self.wb(uid)?;
        let mut sheet = self.sheet_name(sheet);
        let mut areas = vec![];
        for part in text.split(',') {
            let part = part.trim();
            let (sh, body) = crate::cmd::split_sheet(part);
            if let Some(sh) = sh {
                sheet = sh;
            }
            let body = body.replace('$', "");
            if let Some(r) = RangeRef::parse(&body) {
                areas.push(r);
                continue;
            }
            // A defined name.
            let si = wb.sheet_index(&sheet).unwrap_or(wb.active_sheet());
            if let Some(n) = wb.name(&body, si) {
                let (nsh, nbody) = crate::cmd::split_sheet(n.formula.trim_start_matches('='));
                if let Some(nsh) = nsh {
                    sheet = nsh;
                }
                for p in nbody.replace('$', "").split(',') {
                    let (_, b) = crate::cmd::split_sheet(p);
                    areas.push(RangeRef::parse(b).ok_or_else(|| err(1004, format!("Method 'Range' failed: {text}")))?);
                }
                continue;
            }
            return Err(err(1004, format!("Method 'Range' of object '_Global' failed: {text}")));
        }
        Ok(Rng { uid, sheet, kind: 'c', areas })
    }

    /// `Range(a)`, `Range(a, b)` (two cells or ranges: their bounding box).
    fn range_from_args(&mut self, uid: u64, sheet: &str, args: &[V]) -> Result<Rng, VbaError> {
        let one = |me: &Self, v: &V| -> Result<Rng, VbaError> {
            match v {
                V::Object(Obj::Host(h)) if h.class == "Range" => Rng::parse(&h.id).ok_or_else(|| err(1004, "bad range")),
                other => me.parse_range(uid, sheet, &other.to_str()?),
            }
        };
        let a = one(self, arg(args, 0).ok_or_else(|| err(450, "Range needs an address"))?)?;
        match arg(args, 1) {
            None => Ok(a),
            Some(b) => {
                let b = one(self, b)?;
                let (x, y) = (a.first(), b.first());
                let start = CellRef::new(x.start.row.min(y.start.row), x.start.col.min(y.start.col));
                let end = CellRef::new(x.end.row.max(y.end.row), x.end.col.max(y.end.col));
                Ok(a.with('c', vec![RangeRef::new(start, end)]))
            }
        }
    }

    fn whole_sheet(&self, uid: u64, sheet: &str, kind: char) -> Rng {
        let all = RangeRef::new(CellRef::new(0, 0), CellRef::new(gridcraft_core::MAX_ROWS - 1, gridcraft_core::MAX_COLS - 1));
        Rng { uid, sheet: self.sheet_name(sheet), kind, areas: vec![all] }
    }

    /// The cells actually worth reading in `area` (a whole column is cut to the used range).
    fn bounded(&self, r: &Rng, area: RangeRef) -> Result<RangeRef, VbaError> {
        if area.count() <= MAX_CELLS {
            return Ok(area);
        }
        let wb = self.wb(r.uid)?;
        let si = self.sheet_index(r.uid, &r.sheet)?;
        let used = wb.sheet(si).and_then(|s| s.used_range()).unwrap_or(RangeRef::new(area.start, area.start));
        let end = CellRef::new(area.end.row.min(used.end.row.max(area.start.row)), area.end.col.min(used.end.col.max(area.start.col)));
        Ok(RangeRef::new(area.start, end))
    }

    fn cell_value(wb: &Workbook, si: usize, c: CellRef) -> V {
        let Some(sh) = wb.sheet(si) else { return V::Empty };
        match sh.value(c) {
            gridcraft_core::Value::Empty => V::Empty,
            gridcraft_core::Value::Number(n) => {
                let code = wb.styles.get(sh.style_id(c)).num_fmt.0.clone();
                match gridcraft_numfmt::NumberFormat::parse_cached(&code).kind() {
                    gridcraft_numfmt::FormatKind::Date | gridcraft_numfmt::FormatKind::Time => V::Date(serial_to_vba(n)),
                    gridcraft_numfmt::FormatKind::Currency | gridcraft_numfmt::FormatKind::Accounting => V::Currency((n * 10_000.0).round() as i64),
                    _ => V::Double(n),
                }
            }
            gridcraft_core::Value::Text(t) => V::str(t.as_ref()),
            gridcraft_core::Value::Bool(b) => V::Bool(b),
            gridcraft_core::Value::Error(e) => V::Error(error_number(e)),
            gridcraft_core::Value::Array(a) => a
                .get(0, 0)
                .map(|v| match v {
                    gridcraft_core::Value::Number(n) => V::Double(*n),
                    gridcraft_core::Value::Text(t) => V::str(t.as_ref()),
                    gridcraft_core::Value::Bool(b) => V::Bool(*b),
                    _ => V::Empty,
                })
                .unwrap_or(V::Empty),
        }
    }

    /// `Range.Value` (or `Value2`, without dates as dates).
    fn get_value(&self, r: &Rng, value2: bool) -> Result<V, VbaError> {
        let wb = self.wb(r.uid)?;
        let si = self.sheet_index(r.uid, &r.sheet)?;
        let area = self.bounded(r, r.first())?;
        let cell = |c: CellRef| {
            let v = Self::cell_value(wb, si, c);
            match v {
                V::Date(d) if value2 => V::Double(vba_to_serial(d)),
                V::Currency(c) if value2 => V::Double(c as f64 / 10_000.0),
                other => other,
            }
        };
        if area.count() == 1 {
            return Ok(cell(area.start));
        }
        let (h, w) = (area.height() as usize, area.width() as usize);
        let mut items = Vec::with_capacity(h * w);
        for row in area.start.row..=area.end.row {
            for col in area.start.col..=area.end.col {
                items.push(cell(CellRef::new(row, col)));
            }
        }
        Ok(V::Array(std::rc::Rc::new(std::cell::RefCell::new(gridcraft_vba::value::Array {
            bounds: vec![(1, h as i64), (1, w as i64)],
            items,
            fixed: false,
        }))))
    }

    fn to_json(v: &V) -> Result<(Json, bool), VbaError> {
        Ok(match v {
            V::Empty | V::Null => (Json::Null, false),
            V::Bool(b) => (json!(b), false),
            V::String(s) => (json!(s.as_ref()), false),
            V::Date(d) => (json!(vba_to_serial(*d)), true),
            V::Error(n) => (json!(error_text(*n)), false),
            V::Object(_) | V::Nothing => return Err(err(13, "Type mismatch")),
            V::Array(_) | V::Udt(_) => return Err(err(13, "Type mismatch")),
            other => (json!(other.to_f64()?), false),
        })
    }

    /// `Range.Value = v`: a scalar fills every cell, an array fills from the top-left.
    fn set_value(&mut self, r: &Rng, v: V) -> Result<(), VbaError> {
        for area in r.areas.clone() {
            let area = if area.count() > MAX_CELLS { self.bounded(r, area)? } else { area };
            let (h, w) = (area.height() as usize, area.width() as usize);
            let mut rows: Vec<Json> = Vec::with_capacity(h);
            let mut dates = false;
            match &v {
                V::Array(a) => {
                    let a = a.borrow();
                    let get = |i: usize, j: usize| -> V {
                        match a.bounds.len() {
                            1 => {
                                let k = if h == 1 { j } else { i };
                                a.items.get(k).cloned().unwrap_or(V::Error(2042))
                            }
                            _ => {
                                let cols = (a.bounds.get(1).map(|b| b.1 - b.0 + 1).unwrap_or(1)) as usize;
                                if j >= cols || i >= a.items.len() / cols.max(1) {
                                    V::Error(2042)
                                } else {
                                    a.items.get(i * cols + j).cloned().unwrap_or(V::Empty)
                                }
                            }
                        }
                    };
                    for i in 0..h {
                        let mut row = Vec::with_capacity(w);
                        for j in 0..w {
                            let (x, d) = Self::to_json(&get(i, j))?;
                            dates |= d;
                            row.push(x);
                        }
                        rows.push(Json::Array(row));
                    }
                }
                scalar => {
                    let (x, d) = Self::to_json(scalar)?;
                    dates = d;
                    let row = Json::Array(vec![x; w]);
                    rows = vec![row; h];
                }
            }
            let at = Rng { uid: r.uid, sheet: r.sheet.clone(), kind: 'c', areas: vec![area] };
            self.run(
                r.uid,
                "range.setValues",
                json!({"range": format!("{}!{}", gridcraft_formula::quote_sheet(&at.sheet), area.start.a1()), "values": rows}),
            )?;
            if dates {
                self.date_format_if_general(&at)?;
            }
        }
        Ok(())
    }

    fn date_format_if_general(&mut self, r: &Rng) -> Result<(), VbaError> {
        self.restyle(r, |st| {
            if st.num_fmt.0.as_ref() == "General" {
                st.num_fmt = gridcraft_model::NumFmt::new("m/d/yyyy");
            }
        })
    }

    /// Applies `f` to the style of every cell in the range.
    fn restyle(&mut self, r: &Rng, f: impl Fn(&mut gridcraft_model::Style)) -> Result<(), VbaError> {
        let di = self.doc_index(r.uid)?;
        self.s.set_active(di);
        let si = self.sheet_index(r.uid, &r.sheet)?;
        let areas: Vec<RangeRef> = r.areas.iter().map(|a| self.bounded(r, *a)).collect::<Result<_, _>>()?;
        crate::cmd::edit(self.s, |cx| {
            for area in &areas {
                for c in area.iter() {
                    let sh = cx.wb.sheet(si).ok_or(crate::EngineError::NoDocument)?;
                    let mut st = cx.wb.styles.get(sh.style_id(c)).clone();
                    f(&mut st);
                    let id = cx.wb.styles.intern(st);
                    cx.sheet_mut(si)?.set_style(c, id);
                    cx.touch(si, c);
                }
            }
            Ok(())
        })
        .map_err(engine_err)
    }

    /// The style of a range's first cell.
    fn first_style(&self, r: &Rng) -> Result<gridcraft_model::Style, VbaError> {
        let wb = self.wb(r.uid)?;
        let si = self.sheet_index(r.uid, &r.sheet)?;
        let sh = wb.sheet(si).ok_or_else(|| err(9, "sheet"))?;
        Ok(wb.styles.get(sh.style_id(r.first().start)).clone())
    }

    /// `Range.End(direction)` from the range's first cell: Ctrl+arrow.
    fn end(&self, r: &Rng, dir: i64) -> Result<Rng, VbaError> {
        let wb = self.wb(r.uid)?;
        let si = self.sheet_index(r.uid, &r.sheet)?;
        let sh = wb.sheet(si).ok_or_else(|| err(9, "sheet"))?;
        let filled = |c: CellRef| sh.cell(c).is_some_and(|x| !x.value.is_empty() || x.formula.is_some());
        let (dr, dc) = match dir {
            -4162 => (-1i64, 0i64),
            -4121 => (1, 0),
            -4159 => (0, -1),
            _ => (0, 1),
        };
        let mut c = r.first().start;
        let step = |c: CellRef| c.offset(dr, dc);
        let Some(next) = step(c) else { return Ok(r.with('c', vec![RangeRef::new(c, c)])) };
        if filled(c) && filled(next) {
            // Run to the last filled cell.
            c = next;
            while let Some(n) = step(c) {
                if !filled(n) {
                    break;
                }
                c = n;
            }
        } else {
            // Skip to the next filled cell, or the sheet's edge.
            c = next;
            while !filled(c) {
                match step(c) {
                    Some(n) => c = n,
                    None => break,
                }
            }
        }
        Ok(r.with('c', vec![RangeRef::new(c, c)]))
    }

    /// `Item(i)` / `Item(r, c)` / `Cells(r, c)` relative to the range's top-left.
    fn item(&self, r: &Rng, args: &[V]) -> Result<Rng, VbaError> {
        let a = r.first();
        let one = |row: i64, col: i64| -> Result<Rng, VbaError> {
            let c = a.start.offset(row - 1, col - 1).ok_or_else(|| err(1004, "Application-defined or object-defined error"))?;
            Ok(r.with('c', vec![RangeRef::new(c, c)]))
        };
        match (arg(args, 0), arg(args, 1)) {
            (Some(i), Some(j)) => {
                let row = i.to_i64()?;
                let col = match j {
                    V::String(s) => {
                        gridcraft_core::letters_to_col(s).map(|c| c as i64 + 1 - a.start.col as i64).ok_or_else(|| err(13, "Type mismatch"))?
                    }
                    other => other.to_i64()?,
                };
                one(row, col)
            }
            (Some(i), None) => {
                let i = i.to_i64()?;
                match r.kind {
                    'r' => {
                        let row = a.start.row as i64 + i - 1;
                        let row = u32::try_from(row).map_err(|_| err(1004, "row"))?;
                        Ok(r.with('r', vec![RangeRef::new(CellRef::new(row, a.start.col), CellRef::new(row, a.end.col))]))
                    }
                    'k' => {
                        let col = a.start.col as i64 + i - 1;
                        let col = u32::try_from(col).map_err(|_| err(1004, "column"))?;
                        Ok(r.with('k', vec![RangeRef::new(CellRef::new(a.start.row, col), CellRef::new(a.end.row, col))]))
                    }
                    _ => {
                        let w = a.width() as i64;
                        let k = i - 1;
                        one(k.div_euclid(w) + 1, k.rem_euclid(w) + 1)
                    }
                }
            }
            _ => Ok(r.clone()),
        }
    }

    fn enumerate_range(&self, r: &Rng) -> Result<Vec<V>, VbaError> {
        let mut out = vec![];
        for area in &r.areas {
            let area = self.bounded(r, *area)?;
            match r.kind {
                'r' => {
                    for row in area.start.row..=area.end.row {
                        out.push(range_value(&r.with('r', vec![RangeRef::new(CellRef::new(row, area.start.col), CellRef::new(row, area.end.col))])));
                    }
                }
                'k' => {
                    for col in area.start.col..=area.end.col {
                        out.push(range_value(&r.with('k', vec![RangeRef::new(CellRef::new(area.start.row, col), CellRef::new(area.end.row, col))])));
                    }
                }
                _ => {
                    for c in area.iter() {
                        out.push(range_value(&r.with('c', vec![RangeRef::new(c, c)])));
                    }
                }
            }
        }
        Ok(out)
    }

    fn address(&self, r: &Rng, args: &[V]) -> Result<V, VbaError> {
        let row_abs = arg(args, 0).map(|v| v.to_bool()).transpose()?.unwrap_or(true);
        let col_abs = arg(args, 1).map(|v| v.to_bool()).transpose()?.unwrap_or(true);
        let external = arg(args, 3).map(|v| v.to_bool()).transpose()?.unwrap_or(false);
        let cell = |c: CellRef| {
            format!("{}{}{}{}", if col_abs { "$" } else { "" }, gridcraft_core::col_to_letters(c.col), if row_abs { "$" } else { "" }, c.row + 1)
        };
        let parts: Vec<String> =
            r.areas.iter().map(|a| if a.count() == 1 { cell(a.start) } else { format!("{}:{}", cell(a.start), cell(a.end)) }).collect();
        let body = parts.join(",");
        Ok(V::str(if external { format!("{}!{body}", gridcraft_formula::quote_sheet(&r.sheet)) } else { body }))
    }

    fn range_get(&mut self, r: &Rng, member: &str, args: &[V]) -> Result<V, VbaError> {
        let m = member.to_ascii_lowercase();
        let a = r.first();
        match m.as_str() {
            "value" | "_default" if args.is_empty() => self.get_value(r, false),
            "value2" => self.get_value(r, true),
            "item" | "_default" | "cells" if !args.is_empty() => Ok(range_value(&self.item(r, args)?)),
            "cells" => Ok(range_value(&r.with('c', r.areas.clone()))),
            "rows" if args.is_empty() => Ok(range_value(&r.with('r', r.areas.clone()))),
            "rows" => Ok(range_value(&self.item(&r.with('r', r.areas.clone()), args)?)),
            "columns" if args.is_empty() => Ok(range_value(&r.with('k', r.areas.clone()))),
            "columns" => Ok(range_value(&self.item(&r.with('k', r.areas.clone()), args)?)),
            "formula" | "formulalocal" => {
                let wb = self.wb(r.uid)?;
                let si = self.sheet_index(r.uid, &r.sheet)?;
                Ok(V::str(wb.sheet(si).map(|s| s.input_text(a.start)).unwrap_or_default()))
            }
            "formular1c1" => {
                let wb = self.wb(r.uid)?;
                let si = self.sheet_index(r.uid, &r.sheet)?;
                let sh = wb.sheet(si).ok_or_else(|| err(9, "sheet"))?;
                let text = match sh.cell(a.start).and_then(|c| c.formula.clone()).and_then(|f| f.expr()) {
                    Some(e) => format!("={}", gridcraft_formula::print_r1c1(&e, a.start)),
                    None => sh.input_text(a.start),
                };
                Ok(V::str(text))
            }
            "hasformula" => {
                let wb = self.wb(r.uid)?;
                let si = self.sheet_index(r.uid, &r.sheet)?;
                Ok(V::Bool(wb.sheet(si).and_then(|s| s.cell(a.start)).is_some_and(|c| c.formula.is_some())))
            }
            "text" => {
                let wb = self.wb(r.uid)?;
                let si = self.sheet_index(r.uid, &r.sheet)?;
                let sh = wb.sheet(si).ok_or_else(|| err(9, "sheet"))?;
                Ok(V::str(crate::display::cell_text(wb, sh, a.start)))
            }
            "address" => self.address(r, args),
            "row" => Ok(V::Long(a.start.row as i32 + 1)),
            "column" => Ok(V::Long(a.start.col as i32 + 1)),
            "count" | "countlarge" => Ok(V::Double(match r.kind {
                'r' => r.areas.iter().map(|x| x.height() as f64).sum(),
                'k' => r.areas.iter().map(|x| x.width() as f64).sum(),
                _ => r.areas.iter().map(|x| x.count() as f64).sum(),
            })),
            "offset" => {
                let dr = arg(args, 0).map(|v| v.to_i64()).transpose()?.unwrap_or(0);
                let dc = arg(args, 1).map(|v| v.to_i64()).transpose()?.unwrap_or(0);
                let s = a.start.offset(dr, dc).ok_or_else(|| err(1004, "Application-defined or object-defined error"))?;
                let e = a.end.offset(dr, dc).ok_or_else(|| err(1004, "Application-defined or object-defined error"))?;
                Ok(range_value(&r.with(r.kind, vec![RangeRef::new(s, e)])))
            }
            "resize" => {
                let h = arg(args, 0).map(|v| v.to_i64()).transpose()?.unwrap_or(a.height() as i64);
                let w = arg(args, 1).map(|v| v.to_i64()).transpose()?.unwrap_or(a.width() as i64);
                if h < 1 || w < 1 {
                    return Err(err(1004, "Application-defined or object-defined error"));
                }
                let e = a.start.offset(h - 1, w - 1).ok_or_else(|| err(1004, "Application-defined or object-defined error"))?;
                Ok(range_value(&r.with('c', vec![RangeRef::new(a.start, e)])))
            }
            "end" => Ok(range_value(&self.end(r, arg(args, 0).map(|v| v.to_i64()).transpose()?.unwrap_or(-4121))?)),
            "currentregion" => {
                let wb = self.wb(r.uid)?;
                let si = self.sheet_index(r.uid, &r.sheet)?;
                let sh = wb.sheet(si).ok_or_else(|| err(9, "sheet"))?;
                Ok(range_value(&r.with('c', vec![crate::selection::current_region(sh, a.start)])))
            }
            "entirerow" => Ok(range_value(&r.with('r', r.areas.iter().map(|x| RangeRef::rows(x.start.row, x.end.row)).collect()))),
            "entirecolumn" => Ok(range_value(&r.with('k', r.areas.iter().map(|x| RangeRef::cols(x.start.col, x.end.col)).collect()))),
            "worksheet" | "parent" => Ok(self.sheet_handle(r.uid, &r.sheet)),
            "font" => Ok(host("Font", r.id())),
            "interior" => Ok(host("Interior", r.id())),
            "borders" => Ok(host("Borders", r.id())),
            "numberformat" | "numberformatlocal" => Ok(V::str(self.first_style(r)?.num_fmt.0.as_ref())),
            "horizontalalignment" => Ok(V::Long(match self.first_style(r)?.align.h {
                HAlign::Left => -4131,
                HAlign::Center => -4108,
                HAlign::Right => -4152,
                HAlign::Justify => -4130,
                _ => 1,
            })),
            "verticalalignment" => Ok(V::Long(match self.first_style(r)?.align.v {
                VAlign::Top => -4160,
                VAlign::Center => -4108,
                _ => -4107,
            })),
            "wraptext" => Ok(V::Bool(self.first_style(r)?.align.wrap)),
            "columnwidth" => {
                let wb = self.wb(r.uid)?;
                let si = self.sheet_index(r.uid, &r.sheet)?;
                Ok(V::Double(wb.sheet(si).map(|s| gridcraft_model::points_to_chars(s.col_width(a.start.col) as f64)).unwrap_or(8.43)))
            }
            "rowheight" => {
                let wb = self.wb(r.uid)?;
                let si = self.sheet_index(r.uid, &r.sheet)?;
                Ok(V::Double(wb.sheet(si).map(|s| s.row_height(a.start.row) as f64).unwrap_or(15.0)))
            }
            "hidden" => {
                let wb = self.wb(r.uid)?;
                let si = self.sheet_index(r.uid, &r.sheet)?;
                let sh = wb.sheet(si).ok_or_else(|| err(9, "sheet"))?;
                Ok(V::Bool(if r.kind == 'k' { sh.is_col_hidden(a.start.col) } else { sh.is_row_hidden(a.start.row) }))
            }
            "mergecells" => {
                let wb = self.wb(r.uid)?;
                let si = self.sheet_index(r.uid, &r.sheet)?;
                Ok(V::Bool(wb.sheet(si).is_some_and(|s| s.merge_at(a.start).is_some())))
            }
            "areas" => Ok(range_value(r)),
            "select" | "activate" => {
                self.run(r.uid, "selection.set", json!({"range": r.address()}))?;
                Ok(V::Bool(true))
            }
            "clear" => self.run(r.uid, "edit.clearAll", json!({"range": r.address()})).map(|_| V::Empty),
            "clearcontents" => self.run(r.uid, "edit.clearContents", json!({"range": r.address()})).map(|_| V::Empty),
            "clearformats" => self.run(r.uid, "edit.clearFormats", json!({"range": r.address()})).map(|_| V::Empty),
            "clearcomments" | "clearnotes" => self.run(r.uid, "edit.clearComments", json!({"range": r.address()})).map(|_| V::Empty),
            "delete" => {
                let p = match r.kind {
                    'r' => ("home.deleteRows", json!({"rows": format!("{}:{}", a.start.row + 1, a.end.row + 1), "sheet": r.sheet})),
                    'k' => ("home.deleteColumns", json!({"cols": RangeRef::cols(a.start.col, a.end.col).a1(), "sheet": r.sheet})),
                    _ => {
                        let shift = if arg(args, 0).map(|v| v.to_i64()).transpose()? == Some(-4159) { "left" } else { "up" };
                        ("home.deleteCells", json!({"range": r.address(), "shift": shift}))
                    }
                };
                self.run(r.uid, p.0, p.1).map(|_| V::Empty)
            }
            "insert" => {
                let p = match r.kind {
                    'r' => ("home.insertRows", json!({"rows": format!("{}:{}", a.start.row + 1, a.end.row + 1), "sheet": r.sheet})),
                    'k' => ("home.insertColumns", json!({"cols": RangeRef::cols(a.start.col, a.end.col).a1(), "sheet": r.sheet})),
                    _ => {
                        let shift = if arg(args, 0).map(|v| v.to_i64()).transpose()? == Some(-4161) { "right" } else { "down" };
                        ("home.insertCells", json!({"range": r.address(), "shift": shift}))
                    }
                };
                self.run(r.uid, p.0, p.1).map(|_| V::Empty)
            }
            "copy" | "cut" => {
                let cmd = if m == "cut" { "edit.cut" } else { "edit.copy" };
                self.run(r.uid, cmd, json!({"range": r.address()}))?;
                if let Some(V::Object(Obj::Host(dest))) = arg(args, 0)
                    && let Some(d) = Rng::parse(&dest.id)
                {
                    let di = self.sheet_index(d.uid, &d.sheet)?;
                    let _ = di;
                    self.run(d.uid, "selection.set", json!({"range": d.address()}))?;
                    self.run(d.uid, "edit.paste", json!({"at": d.first().start.a1()}))?;
                }
                Ok(V::Empty)
            }
            "pastespecial" => {
                let what = match arg(args, 0).map(|v| v.to_i64()).transpose()?.unwrap_or(-4104) {
                    -4163 => "values",
                    -4122 => "formats",
                    -4123 => "formulas",
                    12 => "valuesAndNumberFormats",
                    8 => "columnWidths",
                    _ => "all",
                };
                let transpose = arg(args, 3).map(|v| v.to_bool()).transpose()?.unwrap_or(false);
                self.run(r.uid, "selection.set", json!({"range": r.address()}))?;
                self.run(r.uid, "edit.pasteSpecial", json!({"what": what, "transpose": transpose})).map(|_| V::Empty)
            }
            "autofit" => {
                match r.kind {
                    'r' => self.run(r.uid, "home.autofitRowHeight", json!({"rows": format!("{}:{}", a.start.row + 1, a.end.row + 1)}))?,
                    _ => self.run(r.uid, "home.autofitColumnWidth", json!({"cols": RangeRef::cols(a.start.col, a.end.col).a1()}))?,
                };
                Ok(V::Empty)
            }
            "merge" => self.run(r.uid, "home.mergeCells", json!({"range": r.address()})).map(|_| V::Empty),
            "unmerge" => self.run(r.uid, "home.unmergeCells", json!({"range": r.address()})).map(|_| V::Empty),
            "find" => self.find(r, args),
            "specialcells" => {
                let kind = arg(args, 0).map(|v| v.to_i64()).transpose()?.unwrap_or(11);
                let wb = self.wb(r.uid)?;
                let si = self.sheet_index(r.uid, &r.sheet)?;
                let used = wb.sheet(si).and_then(|s| s.used_range()).unwrap_or_default();
                match kind {
                    11 => Ok(range_value(&r.with('c', vec![RangeRef::new(used.end, used.end)]))),
                    _ => Ok(range_value(&r.with('c', vec![used]))),
                }
            }
            "calculate" => Ok(V::Empty),
            "show" => Ok(V::Empty),
            "addcommentthreaded" | "addcomment" => {
                let text = arg(args, 0).map(|v| v.to_str()).transpose()?.unwrap_or_default();
                self.add_comment(r, &text, m == "addcommentthreaded")
            }
            "commentthreaded" => self.range_comment(r, true),
            "comment" => self.range_comment(r, false),
            "left" | "top" | "width" | "height" => {
                let wb = self.wb(r.uid)?;
                let si = self.sheet_index(r.uid, &r.sheet)?;
                let sh = wb.sheet(si).ok_or_else(|| err(9, "sheet"))?;
                Ok(V::Double(match m.as_str() {
                    "left" => sh.col_left(a.start.col),
                    "top" => sh.row_top(a.start.row),
                    "width" => sh.col_left(a.end.col.saturating_add(1)) - sh.col_left(a.start.col),
                    _ => sh.row_top(a.end.row.saturating_add(1)) - sh.row_top(a.start.row),
                }))
            }
            _ => Err(err(438, format!("Object doesn't support this property or method: Range.{member}"))),
        }
    }

    fn find(&self, r: &Rng, args: &[V]) -> Result<V, VbaError> {
        let what = arg(args, 0).ok_or_else(|| err(450, "Find needs What"))?.to_str()?;
        let whole = arg(args, 3).map(|v| v.to_i64()).transpose()?.unwrap_or(2) == 1;
        let case = arg(args, 7).map(|v| v.to_bool()).transpose()?.unwrap_or(false);
        let wb = self.wb(r.uid)?;
        let si = self.sheet_index(r.uid, &r.sheet)?;
        let sh = wb.sheet(si).ok_or_else(|| err(9, "sheet"))?;
        let norm = |s: &str| if case { s.to_string() } else { s.to_lowercase() };
        let needle = norm(&what);
        for area in &r.areas {
            let area = self.bounded(r, *area)?;
            for c in area.iter() {
                let text = norm(&crate::display::cell_text(wb, sh, c));
                if (whole && text == needle) || (!whole && !needle.is_empty() && text.contains(&needle)) {
                    return Ok(range_value(&r.with('c', vec![RangeRef::new(c, c)])));
                }
            }
        }
        Ok(V::Nothing)
    }

    fn range_set(&mut self, r: &Rng, member: &str, args: &[V], v: V) -> Result<(), VbaError> {
        let m = member.to_ascii_lowercase();
        match m.as_str() {
            "value" | "value2" | "_default" | "item" | "formula" | "formulalocal" if args.is_empty() || m == "value" || m == "value2" => {
                let target = if (m == "item" || m == "_default") && !args.is_empty() { self.item(r, args)? } else { r.clone() };
                if m.starts_with("formula") {
                    return self.set_value(&target, V::str(v.to_str()?));
                }
                self.set_value(&target, v)
            }
            "item" | "_default" => {
                let target = self.item(r, args)?;
                self.set_value(&target, v)
            }
            "formular1c1" => {
                // Converted cell by cell: relative references shift with each target.
                let text = v.to_str()?;
                let areas = r.areas.clone();
                for area in areas {
                    let area = self.bounded(r, area)?;
                    let mut rows = vec![];
                    for row in area.start.row..=area.end.row {
                        let mut out = vec![];
                        for col in area.start.col..=area.end.col {
                            out.push(json!(r1c1_formula_to_a1(&text, CellRef::new(row, col))));
                        }
                        rows.push(Json::Array(out));
                    }
                    self.run(
                        r.uid,
                        "range.setValues",
                        json!({"range": format!("{}!{}", gridcraft_formula::quote_sheet(&r.sheet), area.start.a1()), "values": rows}),
                    )?;
                }
                Ok(())
            }
            "numberformat" | "numberformatlocal" => {
                let code = v.to_str()?;
                self.restyle(r, move |st| st.num_fmt = gridcraft_model::NumFmt::new(&code))
            }
            "horizontalalignment" => {
                let h = match v.to_i64()? {
                    -4131 => HAlign::Left,
                    -4108 => HAlign::Center,
                    -4152 => HAlign::Right,
                    -4130 => HAlign::Justify,
                    7 => HAlign::CenterAcross,
                    _ => HAlign::General,
                };
                self.restyle(r, move |st| st.align.h = h)
            }
            "verticalalignment" => {
                let va = match v.to_i64()? {
                    -4160 => VAlign::Top,
                    -4108 => VAlign::Center,
                    _ => VAlign::Bottom,
                };
                self.restyle(r, move |st| st.align.v = va)
            }
            "wraptext" => {
                let on = v.to_bool()?;
                self.restyle(r, move |st| st.align.wrap = on)
            }
            "columnwidth" => {
                let chars = v.to_f64()?.clamp(0.0, 255.0);
                // Characters of the default font to points (inverse of `points_to_chars`).
                let points = chars_to_points(chars);
                self.set_lines(r, 'k', Some(points as f32), None)
            }
            "rowheight" => {
                let pts = v.to_f64()?.clamp(0.0, 409.0) as f32;
                self.set_lines(r, 'r', Some(pts), None)
            }
            "hidden" => {
                let on = v.to_bool()?;
                let kind = if r.kind == 'k' { 'k' } else { 'r' };
                self.set_lines(r, kind, None, Some(on))
            }
            "mergecells" => {
                let cmd = if v.to_bool()? { "home.mergeCells" } else { "home.unmergeCells" };
                self.run(r.uid, cmd, json!({"range": r.address()})).map(|_| ())
            }
            _ => Err(err(438, format!("Object doesn't support this property or method: Range.{member}"))),
        }
    }

    /// Row heights / column widths and hidden flags of the range's lines.
    fn set_lines(&mut self, r: &Rng, kind: char, size: Option<f32>, hidden: Option<bool>) -> Result<(), VbaError> {
        let di = self.doc_index(r.uid)?;
        self.s.set_active(di);
        let si = self.sheet_index(r.uid, &r.sheet)?;
        let areas = r.areas.clone();
        crate::cmd::edit(self.s, |cx| {
            let sh = cx.sheet_mut(si)?;
            for a in &areas {
                let (lo, hi) = if kind == 'k' { (a.start.col, a.end.col) } else { (a.start.row, a.end.row) };
                for i in lo..=hi.min(lo.saturating_add(1_048_575)) {
                    let map = if kind == 'k' { &mut sh.cols } else { &mut sh.rows };
                    let line = map.entry(i).or_default();
                    if let Some(s) = size {
                        line.size = Some(s);
                    }
                    if let Some(h) = hidden {
                        line.hidden = h;
                    }
                }
            }
            cx.structural = true;
            Ok(())
        })
        .map_err(engine_err)
    }

    fn font_get(&self, r: &Rng, member: &str) -> Result<V, VbaError> {
        let f = self.first_style(r)?.font;
        Ok(match member.to_ascii_lowercase().as_str() {
            "bold" => V::Bool(f.bold),
            "italic" => V::Bool(f.italic),
            "size" => V::Double(f64::from(f.size)),
            "name" => V::str(f.name.as_str()),
            "color" => color_to_long(f.color),
            "strikethrough" => V::Bool(f.strike),
            "underline" => V::Long(if f.underline == Underline::None { -4142 } else { 2 }),
            "colorindex" => V::Long(-4105),
            _ => return Err(err(438, format!("Object doesn't support this property or method: Font.{member}"))),
        })
    }

    fn font_set(&mut self, r: &Rng, member: &str, v: V) -> Result<(), VbaError> {
        match member.to_ascii_lowercase().as_str() {
            "bold" => {
                let b = v.to_bool()?;
                self.restyle(r, move |st| st.font.bold = b)
            }
            "italic" => {
                let b = v.to_bool()?;
                self.restyle(r, move |st| st.font.italic = b)
            }
            "strikethrough" => {
                let b = v.to_bool()?;
                self.restyle(r, move |st| st.font.strike = b)
            }
            "size" => {
                let n = v.to_f64()?.clamp(1.0, 409.0) as f32;
                self.restyle(r, move |st| st.font.size = n)
            }
            "name" => {
                let n = v.to_str()?;
                self.restyle(r, move |st| st.font.name = n.clone())
            }
            "color" => {
                let c = long_to_color(&v)?;
                self.restyle(r, move |st| st.font.color = c)
            }
            "colorindex" => {
                let c = palette(v.to_i64()?).unwrap_or(Color::Auto);
                self.restyle(r, move |st| st.font.color = c)
            }
            "underline" => {
                let u = match &v {
                    V::Bool(true) => Underline::Single,
                    V::Bool(false) => Underline::None,
                    other => match other.to_i64()? {
                        -4142 | 0 => Underline::None,
                        -4119 => Underline::Double,
                        _ => Underline::Single,
                    },
                };
                self.restyle(r, move |st| st.font.underline = u)
            }
            _ => Err(err(438, format!("Object doesn't support this property or method: Font.{member}"))),
        }
    }

    fn interior_get(&self, r: &Rng, member: &str) -> Result<V, VbaError> {
        let fill = self.first_style(r)?.fill;
        Ok(match member.to_ascii_lowercase().as_str() {
            "color" => match fill.pattern {
                PatternType::None => V::Long(0xFFFFFF),
                _ => color_to_long(fill.fg),
            },
            "colorindex" => V::Long(if fill.pattern == PatternType::None { -4142 } else { 0 }),
            "pattern" => V::Long(if fill.pattern == PatternType::None { -4142 } else { 1 }),
            _ => return Err(err(438, format!("Object doesn't support this property or method: Interior.{member}"))),
        })
    }

    fn interior_set(&mut self, r: &Rng, member: &str, v: V) -> Result<(), VbaError> {
        match member.to_ascii_lowercase().as_str() {
            "color" => {
                let c = long_to_color(&v)?;
                self.restyle(r, move |st| st.fill = gridcraft_model::Fill::solid(c))
            }
            "colorindex" | "pattern" => {
                let n = v.to_i64()?;
                if n == -4142 || n == 0 && member.eq_ignore_ascii_case("colorindex") {
                    return self.restyle(r, |st| st.fill = gridcraft_model::Fill::default());
                }
                if member.eq_ignore_ascii_case("pattern") {
                    return Ok(());
                }
                let c = palette(n).unwrap_or(Color::Auto);
                self.restyle(r, move |st| st.fill = gridcraft_model::Fill::solid(c))
            }
            "tintandshade" | "patterncolorindex" | "themecolor" => Ok(()),
            _ => Err(err(438, format!("Object doesn't support this property or method: Interior.{member}"))),
        }
    }

    fn borders_set(&mut self, r: &Rng, member: &str, v: V) -> Result<(), VbaError> {
        // `Range.Borders.LineStyle = xlContinuous` (all edges) and `Borders.Weight`.
        match member.to_ascii_lowercase().as_str() {
            "linestyle" | "weight" => {
                let none = v.to_i64()? == -4142;
                self.run(r.uid, "home.borders", json!({"range": r.address(), "preset": if none { "none" } else { "all" }})).map(|_| ())
            }
            "color" | "colorindex" => Ok(()),
            _ => Err(err(438, format!("Object doesn't support this property or method: Borders.{member}"))),
        }
    }

    // ------------------------------------------------------------ worksheets and workbooks

    fn worksheet_get(&mut self, uid: u64, name: &str, member: &str, args: &[V]) -> Result<V, VbaError> {
        let name = self.sheet_name(name);
        let m = member.to_ascii_lowercase();
        match m.as_str() {
            "name" | "codename" => Ok(V::str(name.as_str())),
            "range" => Ok(range_value(&self.range_from_args(uid, &name, args)?)),
            "cells" if args.is_empty() => Ok(range_value(&self.whole_sheet(uid, &name, 'c'))),
            "cells" => Ok(range_value(&self.item(&self.whole_sheet(uid, &name, 'c'), args)?)),
            "rows" if args.is_empty() => Ok(range_value(&self.whole_sheet(uid, &name, 'r'))),
            "rows" => {
                if let Some(V::String(s)) = arg(args, 0) {
                    return Ok(range_value(&self.parse_range(uid, &name, s)?.with('r', vec![self.parse_range(uid, &name, s)?.first()])));
                }
                Ok(range_value(&self.item(&self.whole_sheet(uid, &name, 'r'), args)?))
            }
            "columns" if args.is_empty() => Ok(range_value(&self.whole_sheet(uid, &name, 'k'))),
            "columns" => {
                if let Some(V::String(s)) = arg(args, 0) {
                    let r = self.parse_range(uid, &name, s)?;
                    return Ok(range_value(&r.with('k', vec![r.first()])));
                }
                Ok(range_value(&self.item(&self.whole_sheet(uid, &name, 'k'), args)?))
            }
            "usedrange" => {
                let wb = self.wb(uid)?;
                let si = self.sheet_index(uid, &name)?;
                let used = wb.sheet(si).and_then(|s| s.used_range()).unwrap_or_default();
                Ok(range_value(&Rng { uid, sheet: name, kind: 'c', areas: vec![used] }))
            }
            "activate" | "select" => {
                self.run(uid, "sheet.activate", json!({"sheet": name}))?;
                Ok(V::Empty)
            }
            "delete" => {
                self.run(uid, "home.deleteSheet", json!({"sheet": name}))?;
                Ok(V::Empty)
            }
            "index" => Ok(V::Long(self.sheet_index(uid, &name)? as i32 + 1)),
            "visible" => {
                let wb = self.wb(uid)?;
                let si = self.sheet_index(uid, &name)?;
                Ok(V::Long(match wb.sheet(si).map(|s| s.visibility) {
                    Some(gridcraft_model::Visibility::Hidden) => 0,
                    Some(gridcraft_model::Visibility::VeryHidden) => 2,
                    _ => -1,
                }))
            }
            "parent" => Ok(host("Workbook", uid.to_string())),
            "shapes" | "buttons" if args.is_empty() => Ok(host("Shapes", format!("{uid}{SEP}{name}"))),
            "shapes" | "buttons" => self.shapes_get(&format!("{uid}{SEP}{name}"), "Item", args),
            "oleobjects" if args.is_empty() => Ok(host("OLEObjects", format!("{uid}{SEP}{name}"))),
            "calculate" => Ok(V::Empty),
            "protect" => self
                .run(
                    uid,
                    "review.protectSheet",
                    json!({"sheet": name, "password": arg(args, 0).map(|v| v.to_str()).transpose()?.unwrap_or_default()}),
                )
                .map(|_| V::Empty),
            "unprotect" => self
                .run(
                    uid,
                    "review.unprotectSheet",
                    json!({"sheet": name, "password": arg(args, 0).map(|v| v.to_str()).transpose()?.unwrap_or_default()}),
                )
                .map(|_| V::Empty),
            "next" | "previous" => {
                let si = self.sheet_index(uid, &name)? as i64 + if m == "next" { 1 } else { -1 };
                let wb = self.wb(uid)?;
                match usize::try_from(si).ok().and_then(|i| wb.sheet(i)) {
                    Some(s) => Ok(self.sheet_handle(uid, &s.name)),
                    None => Ok(V::Nothing),
                }
            }
            "copy" | "move" => {
                // `ws.Copy After:=Sheets(n)` and friends: copy/move next to the anchor sheet.
                let (anchor, after) = match (arg(args, 0), arg(args, 1)) {
                    (Some(V::Object(Obj::Host(h))), _) => (Some(h.clone()), false),
                    (_, Some(V::Object(Obj::Host(h)))) => (Some(h.clone()), true),
                    _ => (None, true),
                };
                let to = match anchor.and_then(|h| Self::parse_sheet(&h.id)) {
                    Some((u, n)) => self.sheet_index(u, &n)? + after as usize,
                    None => self.wb(uid)?.sheets.len(),
                };
                self.run(uid, "sheet.move", json!({"sheet": name, "to": to, "copy": m == "copy"}))?;
                Ok(V::Empty)
            }
            _ => Err(err(438, format!("Object doesn't support this property or method: Worksheet.{member}"))),
        }
    }

    fn worksheet_set(&mut self, uid: u64, name: &str, member: &str, v: V) -> Result<(), VbaError> {
        let name = self.sheet_name(name);
        match member.to_ascii_lowercase().as_str() {
            "name" => {
                let new = v.to_str()?;
                self.run(uid, "sheet.rename", json!({"sheet": name, "name": new}))?;
                self.renamed.push((name.to_lowercase(), new));
                Ok(())
            }
            "visible" => {
                let hidden = !matches!(&v, V::Bool(true)) && v.to_i64()? != -1;
                let cmd = if hidden { "sheet.hide" } else { "sheet.unhide" };
                self.run(uid, cmd, json!({"sheet": name})).map(|_| ())
            }
            _ => Err(err(438, format!("Object doesn't support this property or method: Worksheet.{member}"))),
        }
    }

    fn sheets_get(&mut self, uid: u64, member: &str, args: &[V]) -> Result<V, VbaError> {
        match member.to_ascii_lowercase().as_str() {
            "item" | "_default" => {
                let wb = self.wb(uid)?;
                let name = match arg(args, 0) {
                    Some(V::String(s)) => {
                        let s = self.sheet_name(s);
                        wb.sheet_index(&s).and_then(|i| wb.sheet(i)).map(|s| s.name.clone()).ok_or_else(|| err(9, "Subscript out of range"))?
                    }
                    Some(i) => {
                        let i = i.to_i64()?;
                        usize::try_from(i - 1)
                            .ok()
                            .and_then(|i| wb.sheet(i))
                            .map(|s| s.name.clone())
                            .ok_or_else(|| err(9, "Subscript out of range"))?
                    }
                    None => return Err(err(450, "Wrong number of arguments")),
                };
                Ok(self.sheet_handle(uid, &name))
            }
            "count" => Ok(V::Long(self.wb(uid)?.sheets.len() as i32)),
            "add" => {
                let before = match (arg(args, 0), arg(args, 1)) {
                    (Some(V::Object(Obj::Host(h))), _) => Self::parse_sheet(&h.id).map(|(u, n)| self.sheet_index(u, &n)).transpose()?,
                    (_, Some(V::Object(Obj::Host(h)))) => {
                        Self::parse_sheet(&h.id).map(|(u, n)| self.sheet_index(u, &n).map(|i| i + 1)).transpose()?
                    }
                    _ => None,
                };
                let mut p = json!({});
                if let Some(b) = before {
                    p["before"] = json!(b);
                }
                let r = self.run(uid, "home.insertSheet", p)?;
                let name = r["name"].as_str().unwrap_or_default().to_string();
                Ok(self.sheet_handle(uid, &name))
            }
            _ => Err(err(438, format!("Object doesn't support this property or method: Sheets.{member}"))),
        }
    }

    fn workbook_get(&mut self, uid: u64, member: &str, args: &[V]) -> Result<V, VbaError> {
        match member.to_ascii_lowercase().as_str() {
            "name" => {
                let d = self.s.documents().get(self.doc_index(uid)?).ok_or_else(|| err(9, "workbook"))?;
                Ok(V::str(d.display_title()))
            }
            "fullname" | "path" => {
                let d = self.s.documents().get(self.doc_index(uid)?).ok_or_else(|| err(9, "workbook"))?;
                let full = d.path.clone().unwrap_or_else(|| d.display_title());
                if member.eq_ignore_ascii_case("path") {
                    return Ok(V::str(std::path::Path::new(&full).parent().map(|p| p.to_string_lossy().to_string()).unwrap_or_default()));
                }
                Ok(V::str(full))
            }
            "worksheets" | "sheets" if args.is_empty() => Ok(host("Sheets", uid.to_string())),
            "worksheets" | "sheets" => self.sheets_get(uid, "Item", args),
            "activesheet" => {
                let name = self.active_sheet_name(uid)?;
                Ok(self.sheet_handle(uid, &name))
            }
            "save" => self.run(uid, "file.save", json!({})).map(|_| V::Empty),
            "saveas" | "savecopyas" => {
                let path = arg(args, 0).ok_or_else(|| err(450, "SaveAs needs a file name"))?.to_str()?;
                self.run(uid, "file.save", json!({"path": path})).map(|_| V::Empty)
            }
            "saved" => {
                let d = self.s.documents().get(self.doc_index(uid)?).ok_or_else(|| err(9, "workbook"))?;
                Ok(V::Bool(!d.is_dirty()))
            }
            "close" | "activate" => Ok(V::Empty),
            "refreshall" => self.run(uid, "data.refreshAll", json!({})).map(|_| V::Empty),
            _ => Err(err(438, format!("Object doesn't support this property or method: Workbook.{member}"))),
        }
    }

    fn worksheet_function(&self, name: &str, args: &[V]) -> Result<V, VbaError> {
        let uid = self.active_uid();
        let wb = self.wb(uid)?;
        let mut parts = vec![];
        for a in args {
            parts.push(match a {
                V::Object(Obj::Host(h)) if h.class == "Range" => Rng::parse(&h.id).map(|r| r.address()).ok_or_else(|| err(1004, "range"))?,
                V::String(s) => format!("\"{}\"", s.replace('"', "\"\"")),
                V::Bool(b) => (if *b { "TRUE" } else { "FALSE" }).to_string(),
                V::Empty => String::new(),
                V::Array(arr) => {
                    let a = arr.borrow();
                    let items: Vec<String> = a
                        .items
                        .iter()
                        .map(|x| match x {
                            V::String(s) => format!("\"{}\"", s.replace('"', "\"\"")),
                            other => other.to_f64().map(|n| n.to_string()).unwrap_or_default(),
                        })
                        .collect();
                    format!("{{{}}}", items.join(","))
                }
                other => other.to_f64()?.to_string(),
            });
        }
        let formula = format!("={name}({})", parts.join(","));
        let v = gridcraft_calc::evaluate(wb, wb.active_sheet(), CellRef::new(0, 0), &formula);
        core_to_vba(v)
    }

    fn evaluate(&self, text: &str) -> Result<V, VbaError> {
        let uid = self.active_uid();
        let wb = self.wb(uid)?;
        let f = if text.starts_with('=') { text.to_string() } else { format!("={text}") };
        let v = gridcraft_calc::evaluate(wb, wb.active_sheet(), CellRef::new(0, 0), &f);
        core_to_vba(v)
    }

    fn application_get(&mut self, member: &str, args: &[V]) -> Result<V, VbaError> {
        let m = member.to_ascii_lowercase();
        match m.as_str() {
            "screenupdating" | "displayalerts" | "enableevents" | "interactive" | "displaystatusbar" => {
                Ok(self.app.get(&m).cloned().unwrap_or(V::Bool(true)))
            }
            "calculation" => Ok(self.app.get(&m).cloned().unwrap_or(V::Long(-4105))),
            "cutcopymode" => Ok(V::Bool(false)),
            "statusbar" => Ok(self.app.get(&m).cloned().unwrap_or(V::Bool(false))),
            "version" => Ok(V::str("16.0")),
            "name" => Ok(V::str("Microsoft Excel")),
            "username" => Ok(V::str(self.s.prefs.user_name.as_str())),
            "operatingsystem" => Ok(V::str(std::env::consts::OS)),
            "worksheetfunction" => Ok(host("WorksheetFunction", "")),
            "evaluate" => self.evaluate(&arg(args, 0).ok_or_else(|| err(450, "Evaluate"))?.to_str()?),
            "calculate" | "calculatefull" | "volatile" | "wait" | "quit" | "ontime" | "goto" => Ok(V::Empty),
            "inputbox" => {
                let prompt = arg(args, 0).map(|v| v.to_str()).transpose()?.unwrap_or_default();
                let title = arg(args, 1).map(|v| v.to_str()).transpose()?.unwrap_or_else(|| "Input".into());
                let default = arg(args, 2).map(|v| v.to_str()).transpose()?.unwrap_or_default();
                Ok(match self.ui.inputbox(&prompt, &title, &default) {
                    Some(t) => {
                        // Type:=1 asks for a number.
                        if arg(args, 7).map(|v| v.to_i64()).transpose()? == Some(1) {
                            gridcraft_vba::value::parse_number(&t).map(V::Double).unwrap_or(V::Bool(false))
                        } else {
                            V::str(t)
                        }
                    }
                    None => V::Bool(false),
                })
            }
            "intersect" => {
                let mut acc: Option<Rng> = None;
                for a in args {
                    let V::Object(Obj::Host(h)) = a else { continue };
                    let Some(r) = Rng::parse(&h.id) else { continue };
                    acc = Some(match acc {
                        None => r,
                        Some(x) => {
                            let (p, q) = (x.first(), r.first());
                            let s = CellRef::new(p.start.row.max(q.start.row), p.start.col.max(q.start.col));
                            let e = CellRef::new(p.end.row.min(q.end.row), p.end.col.min(q.end.col));
                            if s.row > e.row || s.col > e.col || !x.sheet.eq_ignore_ascii_case(&r.sheet) {
                                return Ok(V::Nothing);
                            }
                            x.with('c', vec![RangeRef::new(s, e)])
                        }
                    });
                }
                Ok(acc.map(|r| range_value(&r)).unwrap_or(V::Nothing))
            }
            "union" => {
                let mut acc: Option<Rng> = None;
                for a in args {
                    let V::Object(Obj::Host(h)) = a else { continue };
                    let Some(r) = Rng::parse(&h.id) else { continue };
                    acc = Some(match acc {
                        None => r,
                        Some(mut x) => {
                            x.areas.extend(r.areas);
                            x
                        }
                    });
                }
                Ok(acc.map(|r| range_value(&r)).unwrap_or(V::Nothing))
            }
            "getopenfilename" | "getsaveasfilename" => Ok(V::Bool(false)),
            "caller" => Ok(self.caller.clone().map(V::str).unwrap_or(V::Error(2023))),
            "pathseparator" => Ok(V::str(std::path::MAIN_SEPARATOR.to_string())),
            "dialogs" => Err(err(1004, "Built-in Excel dialogs (Application.Dialogs) aren't available in GridCraft")),
            "filedialog" => Err(err(1004, "Application.FileDialog isn't available in GridCraft yet")),
            _ => self
                .global(member)
                .and_then(|g| g.ok_or_else(|| err(438, format!("Object doesn't support this property or method: Application.{member}"))))
                .and_then(|v| {
                    if args.is_empty() {
                        return Ok(v);
                    }
                    match &v {
                        V::Object(Obj::Host(h)) => {
                            let h = h.clone();
                            self.get_member(&h, "Item", args)
                        }
                        _ => Ok(v),
                    }
                }),
        }
    }
}

fn core_to_vba(v: gridcraft_core::Value) -> Result<V, VbaError> {
    Ok(match v {
        gridcraft_core::Value::Empty => V::Empty,
        gridcraft_core::Value::Number(n) => V::Double(n),
        gridcraft_core::Value::Text(t) => V::str(t.as_ref()),
        gridcraft_core::Value::Bool(b) => V::Bool(b),
        // WorksheetFunction raises where a cell would show an error.
        gridcraft_core::Value::Error(e) => {
            return Err(err(1004, format!("Unable to get the property of the WorksheetFunction class ({})", error_text(error_number(e)))));
        }
        gridcraft_core::Value::Array(a) => {
            let (h, w) = (a.rows, a.cols);
            let mut items = vec![];
            for r in 0..h {
                for c in 0..w {
                    items.push(match a.get(r, c) {
                        Some(gridcraft_core::Value::Number(n)) => V::Double(*n),
                        Some(gridcraft_core::Value::Text(t)) => V::str(t.as_ref()),
                        Some(gridcraft_core::Value::Bool(b)) => V::Bool(*b),
                        _ => V::Empty,
                    });
                }
            }
            if h == 1 && w == 1 {
                return Ok(items.into_iter().next().unwrap_or(V::Empty));
            }
            V::Array(std::rc::Rc::new(std::cell::RefCell::new(gridcraft_vba::value::Array {
                bounds: vec![(1, h as i64), (1, w as i64)],
                items,
                fixed: false,
            })))
        }
    })
}

/// Column width in characters of the default font to points (inverse of `points_to_chars`).
fn chars_to_points(chars: f64) -> f64 {
    let mut lo = 0.0f64;
    let mut hi = 2000.0f64;
    for _ in 0..50 {
        let mid = (lo + hi) / 2.0;
        if gridcraft_model::points_to_chars(mid) < chars {
            lo = mid;
        } else {
            hi = mid;
        }
    }
    hi
}

/// An R1C1 formula (`=R[-1]C+RC[2]`) as A1 for the cell `at`. Text in quotes is kept.
pub fn r1c1_formula_to_a1(f: &str, at: CellRef) -> String {
    let chars: Vec<char> = f.chars().collect();
    let mut out = String::with_capacity(f.len());
    let mut i = 0usize;
    let ident = |c: char| c.is_alphanumeric() || c == '_' || c == '.';
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            out.push(c);
            i += 1;
            while i < chars.len() {
                out.push(chars[i]);
                if chars[i] == '"' {
                    if chars.get(i + 1) == Some(&'"') {
                        out.push('"');
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        let boundary = i == 0 || !ident(chars[i - 1]);
        if boundary
            && (c == 'R' || c == 'r' || c == 'C' || c == 'c')
            && let Some((text, len)) = r1c1_ref(&chars[i..], at)
            && !chars.get(i + len).is_some_and(|&n| ident(n) || n == '(')
        {
            out.push_str(&text);
            i += len;
            continue;
        }
        out.push(c);
        i += 1;
    }
    out
}

/// One R1C1 reference at the start of `s`: (A1 text, characters used).
fn r1c1_ref(s: &[char], at: CellRef) -> Option<(String, usize)> {
    // Part: R or C, then [n] (relative), n (absolute) or nothing (same row/col).
    fn part(s: &[char], i: &mut usize, letter: char) -> Option<Option<(i64, bool)>> {
        if !s.get(*i).is_some_and(|c| c.eq_ignore_ascii_case(&letter)) {
            return None;
        }
        *i += 1;
        if s.get(*i) == Some(&'[') {
            let end = s[*i..].iter().position(|&c| c == ']')? + *i;
            let n: i64 = s[*i + 1..end].iter().collect::<String>().parse().ok()?;
            *i = end + 1;
            return Some(Some((n, false)));
        }
        let start = *i;
        while s.get(*i).is_some_and(|c| c.is_ascii_digit()) {
            *i += 1;
        }
        if *i > start {
            let n: i64 = s[start..*i].iter().collect::<String>().parse().ok()?;
            return Some(Some((n, true)));
        }
        Some(None)
    }
    let mut i = 0usize;
    let row = part(s, &mut i, 'R');
    let col = part(s, &mut i, 'C');
    if row.is_none() && col.is_none() {
        return None;
    }
    let resolve = |p: Option<(i64, bool)>, base: u32| -> Option<(u32, bool)> {
        match p {
            None => Some((base, false)),
            Some((n, true)) => u32::try_from(n - 1).ok().map(|v| (v, true)),
            Some((n, false)) => u32::try_from(base as i64 + n).ok().map(|v| (v, false)),
        }
    };
    let col_text = |(c, abs): (u32, bool)| format!("{}{}", if abs { "$" } else { "" }, gridcraft_core::col_to_letters(c));
    let row_text = |(r, abs): (u32, bool)| format!("{}{}", if abs { "$" } else { "" }, r + 1);
    match (row, col) {
        (Some(r), Some(c)) => Some((format!("{}{}", col_text(resolve(c, at.col)?), row_text(resolve(r, at.row)?)), i)),
        // Whole rows / columns: `R2` → `$2:$2`, `C[1]` → `B:B`.
        (Some(r), None) => {
            let rt = row_text(resolve(r, at.row)?);
            Some((format!("{rt}:{rt}"), i))
        }
        (None, Some(c)) => {
            let ct = col_text(resolve(c, at.col)?);
            Some((format!("{ct}:{ct}"), i))
        }
        (None, None) => None,
    }
}

impl Host for ExcelHost<'_> {
    fn global(&mut self, name: &str) -> Result<Option<V>, VbaError> {
        let uid = self.active_uid();
        Ok(Some(match name.to_ascii_lowercase().as_str() {
            "application" => host("Application", ""),
            "activeworkbook" => host("Workbook", uid.to_string()),
            "thisworkbook" => host("Workbook", self.this_uid.to_string()),
            "workbooks" => host("Workbooks", ""),
            "worksheets" | "sheets" => host("Sheets", uid.to_string()),
            "activesheet" => {
                let n = self.active_sheet_name(uid)?;
                self.sheet_handle(uid, &n)
            }
            "activecell" => {
                let d = self.s.active().ok_or_else(|| err(91, "no workbook"))?;
                let c = d.selection.active;
                let n = self.active_sheet_name(uid)?;
                range_value(&Rng { uid, sheet: n, kind: 'c', areas: vec![RangeRef::new(c, c)] })
            }
            "selection" => {
                let d = self.s.active().ok_or_else(|| err(91, "no workbook"))?;
                let areas = d.selection.ranges.clone();
                let n = self.active_sheet_name(uid)?;
                range_value(&Rng { uid, sheet: n, kind: 'c', areas })
            }
            "range" => host("RangeFn", format!("{uid}{SEP}{}", self.active_sheet_name(uid)?)),
            "cells" => range_value(&self.whole_sheet(uid, &self.active_sheet_name(uid)?, 'c')),
            "rows" => range_value(&self.whole_sheet(uid, &self.active_sheet_name(uid)?, 'r')),
            "columns" => range_value(&self.whole_sheet(uid, &self.active_sheet_name(uid)?, 'k')),
            "worksheetfunction" => host("WorksheetFunction", ""),
            _ => match constant(name) {
                Some(v) => v,
                None => {
                    // A sheet's code name (`Sheet1.Range(…)`) or a sheet by name.
                    let wb = self.wb(uid)?;
                    match wb.sheet_index(name) {
                        Some(i) => {
                            let n = wb.sheet(i).map(|s| s.name.clone()).unwrap_or_default();
                            self.sheet_handle(uid, &n)
                        }
                        None => return Ok(None),
                    }
                }
            },
        }))
    }

    fn create_object(&mut self, class: &str) -> Result<Option<V>, VbaError> {
        Ok(match class.to_ascii_lowercase().as_str() {
            "adodb.stream" => Some(self.new_stream()),
            _ => None,
        })
    }

    fn call_global(&mut self, name: &str, args: &[V]) -> Result<Option<V>, VbaError> {
        let text = |i: usize| arg(args, i).map(|v| v.to_str()).transpose().map(Option::unwrap_or_default);
        match name.to_ascii_lowercase().as_str() {
            "dir" | "dir$" => return objects::dir(&text(0)?).map(Some),
            "kill" => return objects::kill(&text(0)?).map(|_| Some(V::Empty)),
            "filecopy" | "mkdir" | "rmdir" => return objects::file_op(&name.to_ascii_lowercase(), &text(0)?, &text(1)?).map(|_| Some(V::Empty)),
            "environ" | "environ$" => return Ok(Some(V::str(std::env::var(text(0)?).unwrap_or_default()))),
            "curdir" | "curdir$" => return Ok(Some(V::str(std::env::current_dir().map(|p| p.to_string_lossy().to_string()).unwrap_or_default()))),
            _ => {}
        }
        match name.to_ascii_lowercase().as_str() {
            "evaluate" => Ok(Some(self.evaluate(&arg(args, 0).ok_or_else(|| err(450, "Evaluate"))?.to_str()?)?)),
            "intersect" | "union" => self.application_get(name, args).map(Some),
            _ => Ok(None),
        }
    }

    fn get_member(&mut self, obj: &HostRef, member: &str, args: &[V]) -> Result<V, VbaError> {
        match obj.class.as_str() {
            "Range" => {
                let r = Rng::parse(&obj.id).ok_or_else(|| err(1004, "bad range"))?;
                self.range_get(&r, member, args)
            }
            "RangeFn" => {
                let (uid, sheet) = Self::parse_sheet(&obj.id).ok_or_else(|| err(1004, "range"))?;
                Ok(range_value(&self.range_from_args(uid, &sheet, args)?))
            }
            "Font" => {
                let r = Rng::parse(&obj.id).ok_or_else(|| err(1004, "bad range"))?;
                self.font_get(&r, member)
            }
            "Interior" => {
                let r = Rng::parse(&obj.id).ok_or_else(|| err(1004, "bad range"))?;
                self.interior_get(&r, member)
            }
            "Borders" => match member.to_ascii_lowercase().as_str() {
                "item" | "_default" => Ok(host("Borders", obj.id.clone())),
                _ => Ok(V::Long(1)),
            },
            "Worksheet" => {
                let (uid, name) = Self::parse_sheet(&obj.id).ok_or_else(|| err(1004, "sheet"))?;
                self.worksheet_get(uid, &name, member, args)
            }
            "Sheets" => {
                let uid = obj.id.parse().map_err(|_| err(1004, "sheets"))?;
                self.sheets_get(uid, member, args)
            }
            "Workbook" => {
                let uid = obj.id.parse().map_err(|_| err(1004, "workbook"))?;
                self.workbook_get(uid, member, args)
            }
            "Workbooks" => match member.to_ascii_lowercase().as_str() {
                "count" => Ok(V::Long(self.s.documents().len() as i32)),
                "item" | "_default" => {
                    let d = match arg(args, 0) {
                        Some(V::String(s)) => self.s.documents().iter().find(|d| d.display_title().eq_ignore_ascii_case(s)),
                        Some(i) => usize::try_from(i.to_i64()? - 1).ok().and_then(|i| self.s.documents().get(i)),
                        None => None,
                    };
                    d.map(|d| host("Workbook", d.uid.to_string())).ok_or_else(|| err(9, "Subscript out of range"))
                }
                "add" => {
                    self.s.execute("file.new", json!({})).map_err(engine_err)?;
                    Ok(host("Workbook", self.active_uid().to_string()))
                }
                "open" => {
                    let path = arg(args, 0).ok_or_else(|| err(450, "Open needs a file name"))?.to_str()?;
                    self.s.execute("file.open", json!({"path": path})).map_err(engine_err)?;
                    Ok(host("Workbook", self.active_uid().to_string()))
                }
                _ => Err(err(438, format!("Object doesn't support this property or method: Workbooks.{member}"))),
            },
            "WorksheetFunction" => self.worksheet_function(member, args),
            "Shapes" => self.shapes_get(&obj.id, member, args),
            "Shape" => self.shape_get(&obj.id, member),
            "TextFrame" => self.text_frame_get(&obj.id, member),
            "Characters" => self.characters_get(&obj.id, member),
            "CommentThreaded" | "Comment" => self.comment_get(&obj.id, member, args),
            "Stream" => self.stream_call(&obj.id, member, args),
            "OLEObjects" => match member.to_ascii_lowercase().as_str() {
                "count" => Ok(V::Long(0)),
                _ => Err(err(1004, "Embedding objects (OLEObjects.Add) isn't available in GridCraft")),
            },
            "Application" => self.application_get(member, args),
            _ => Err(err(438, format!("Object doesn't support this property or method: {member}"))),
        }
    }

    fn set_member(&mut self, obj: &HostRef, member: &str, args: &[V], value: V, _by_ref: bool) -> Result<(), VbaError> {
        match obj.class.as_str() {
            "Range" => {
                let r = Rng::parse(&obj.id).ok_or_else(|| err(1004, "bad range"))?;
                self.range_set(&r, member, args, value)
            }
            "RangeFn" => {
                let (uid, sheet) = Self::parse_sheet(&obj.id).ok_or_else(|| err(1004, "range"))?;
                let r = self.range_from_args(uid, &sheet, args)?;
                self.set_value(&r, value)
            }
            "Font" => {
                let r = Rng::parse(&obj.id).ok_or_else(|| err(1004, "bad range"))?;
                self.font_set(&r, member, value)
            }
            "Interior" => {
                let r = Rng::parse(&obj.id).ok_or_else(|| err(1004, "bad range"))?;
                self.interior_set(&r, member, value)
            }
            "Borders" => {
                let r = Rng::parse(&obj.id).ok_or_else(|| err(1004, "bad range"))?;
                self.borders_set(&r, member, value)
            }
            "Worksheet" => {
                let (uid, name) = Self::parse_sheet(&obj.id).ok_or_else(|| err(1004, "sheet"))?;
                self.worksheet_set(uid, &name, member, value)
            }
            "Application" => {
                let m = member.to_ascii_lowercase();
                if m == "statusbar" {
                    self.ui.status(match &value {
                        V::Bool(false) => None,
                        v => Some(v.to_str()?),
                    });
                }
                if m == "calculation" {
                    let manual = value.to_i64()? == -4135;
                    let _ = self.s.execute("formulas.calculationOptions", json!({"mode": if manual { "manual" } else { "automatic" }}));
                }
                self.app.insert(m, value);
                Ok(())
            }
            "Workbook" if member.eq_ignore_ascii_case("Saved") => Ok(()),
            "Shape" => self.shape_set(&obj.id, member, value),
            "Characters" | "TextFrame" => self.shape_set(&obj.id, member, value),
            "Stream" => self.stream_set(&obj.id, member, value),
            _ => Err(err(438, format!("Can't set {member}"))),
        }
    }

    fn enumerate(&mut self, obj: &HostRef) -> Result<Vec<V>, VbaError> {
        match obj.class.as_str() {
            "Range" => {
                let r = Rng::parse(&obj.id).ok_or_else(|| err(1004, "bad range"))?;
                self.enumerate_range(&r)
            }
            "Sheets" => {
                let uid: u64 = obj.id.parse().map_err(|_| err(1004, "sheets"))?;
                let names: Vec<String> = self.wb(uid)?.sheets.iter().map(|s| s.name.clone()).collect();
                Ok(names.iter().map(|n| self.sheet_handle(uid, n)).collect())
            }
            "Workbooks" => Ok(self.s.documents().iter().map(|d| host("Workbook", d.uid.to_string())).collect()),
            "Shapes" => self.shapes_enumerate(&obj.id),
            "OLEObjects" => Ok(vec![]),
            _ => Err(err(438, "Object doesn't support this property or method")),
        }
    }

    fn msgbox(&mut self, prompt: &str, buttons: i64, title: &str) -> Result<i64, VbaError> {
        Ok(self.ui.msgbox(prompt, buttons, title))
    }

    fn inputbox(&mut self, prompt: &str, title: &str, default: &str) -> Result<Option<String>, VbaError> {
        Ok(self.ui.inputbox(prompt, title, default))
    }

    fn debug_print(&mut self, text: &str) {
        self.ui.debug_print(text);
    }
}

fn program(bin: &[u8]) -> Result<gridcraft_vba::Program, String> {
    let proj = gridcraft_xlsx::ovba::read_project(bin)?;
    let sources: Vec<(String, String)> = proj.modules.iter().map(|m| (m.name.clone(), m.source.clone())).collect();
    gridcraft_vba::Program::parse(&sources).map_err(|e| e.to_string())
}

/// The runnable macros of a VBA project (`Module1.Main`…).
pub fn macro_names(bin: &[u8]) -> Result<Vec<String>, String> {
    Ok(program(bin)?.macros())
}

/// Every procedure of a VBA project, private ones and event handlers included.
pub fn all_procedures(bin: &[u8]) -> Result<Vec<String>, String> {
    Ok(program(bin)?.procedures())
}

/// Runs macro `name` (`Proc` or `Module.Proc`) of the active workbook's VBA project. Everything
/// it changes undoes as one step. Not journaled itself: replaying a journal never re-runs it.
pub fn run_macro(s: &mut Session, name: &str, ui: &mut dyn MacroUi) -> Result<V, VbaError> {
    run_macro_with(s, name, ui, None)
}

/// [`run_macro`] that stops when `cancel` becomes true.
pub fn run_macro_with(
    s: &mut Session,
    name: &str,
    ui: &mut dyn MacroUi,
    cancel: Option<std::sync::Arc<std::sync::atomic::AtomicBool>>,
) -> Result<V, VbaError> {
    run_macro_from(s, name, ui, cancel, None)
}

/// [`run_macro_with`] started by a button: `caller` is what `Application.Caller` returns.
pub fn run_macro_from(
    s: &mut Session,
    name: &str,
    ui: &mut dyn MacroUi,
    cancel: Option<std::sync::Arc<std::sync::atomic::AtomicBool>>,
    caller: Option<String>,
) -> Result<V, VbaError> {
    let proj = crate::cmd::vba::project(s).map_err(engine_err)?.ok_or_else(|| err(35, "This workbook has no macros"))?;
    let sources: Vec<(String, String)> = proj.modules.iter().map(|m| (m.name.clone(), m.source.clone())).collect();
    let program = gridcraft_vba::Program::parse(&sources)?;
    let uid = s.active().map(|d| d.uid).ok_or_else(|| err(91, "no workbook"))?;
    let undo_before = s.active().map(|d| d.undo.len()).unwrap_or(0);
    let journal_before = s.journal.len();
    let result = on_macro_thread(|budget| {
        let mut host = ExcelHost::new(s, ui);
        host.caller = caller;
        let mut interp = gridcraft_vba::Interp::new(program, &mut host);
        if let Some(budget) = budget {
            interp.set_stack_budget(budget);
        }
        if let Some(c) = cancel {
            interp.set_cancel(c);
        }
        // A value can't leave the macro's thread (it may hold references): its text does.
        interp.run(name, vec![]).map(|v| v.to_str().unwrap_or_default())
    })
    .map(V::str);
    // One undo step for the whole macro: keep the state from before it, drop the rest.
    if let Some(d) = s.docs.iter_mut().find(|d| d.uid == uid)
        && d.undo.len() > undo_before + 1
    {
        d.undo.truncate(undo_before + 1);
    }
    if let Some(d) = s.docs.iter_mut().find(|d| d.uid == uid)
        && let Some(e) = d.undo.get_mut(undo_before)
    {
        e.label = format!("Macro {name}");
    }
    s.journal.truncate(journal_before);
    // `End` stops the macro without an error.
    match result {
        Err(e) if e.number == -1 => Ok(V::Empty),
        Err(mut e) if e.number == -2 => {
            e.number = 18;
            Err(e)
        }
        other => other,
    }
}

/// Native stack of the thread a macro runs on, and how much of it the interpreter may use (the
/// rest is for the workbook work it asks for: a formula it calculates has its own budget).
const MACRO_STACK: usize = 64 << 20;
const MACRO_STACK_BUDGET: usize = 48 << 20;

/// Runs a macro on a thread of its own with [`MACRO_STACK`] of stack, whatever thread asked
/// (a Windows main thread has 1 MiB), giving `run` the stack budget it may use. A panic in it
/// (a bug) ends the macro with an error; the workbook stays as the macro left it. Without
/// threads (the web) it runs here with the interpreter's own small budget.
fn on_macro_thread(run: impl FnOnce(Option<usize>) -> Result<String, VbaError> + Send) -> Result<String, VbaError> {
    if cfg!(target_arch = "wasm32") {
        return run(None);
    }
    std::thread::scope(|scope| {
        let worker =
            std::thread::Builder::new().name("gridcraft-vba".into()).stack_size(MACRO_STACK).spawn_scoped(scope, || run(Some(MACRO_STACK_BUDGET)));
        match worker {
            Ok(handle) => handle.join().unwrap_or_else(|_| Err(err(51, "The macro stopped because of an internal error"))),
            Err(e) => Err(err(7, format!("The macro couldn't start: {e}"))),
        }
    })
}

mod objects;

#[cfg(test)]
mod tests;

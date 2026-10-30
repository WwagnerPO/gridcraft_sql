//! More of Excel's object model: `Shapes` and form-control buttons (`Application.Caller`,
//! `TopLeftCell`, `TextFrame.Characters.Text`), threaded comments and notes on a range,
//! `ADODB.Stream` for writing text files, and the file functions `Dir`, `Kill`, `FileCopy`,
//! `MkDir`.

use std::sync::Arc;

use gridcraft_core::{CellRef, RangeRef};
use gridcraft_vba::error::err;
use gridcraft_vba::{Value as V, VbaError};
use gridcraft_xlsx::form_controls::{self, FormButton};
use serde_json::json;

use super::{ExcelHost, Rng, SEP, host, range_value};

/// An `ADODB.Stream` being written.
#[derive(Default, Debug)]
pub(super) struct Stream {
    charset: String,
    text: String,
    open: bool,
}

impl ExcelHost<'_> {
    // ------------------------------------------------------------ shapes and buttons

    fn buttons_of(&self, uid: u64, sheet: &str) -> Result<Vec<FormButton>, VbaError> {
        let wb = self.wb(uid)?;
        let si = self.sheet_index(uid, sheet)?;
        Ok(wb.sheet(si).and_then(|s| s.preserved.as_deref()).map(form_controls::buttons).unwrap_or_default())
    }

    fn shape_handle(uid: u64, sheet: &str, b: &FormButton) -> V {
        host("Shape", format!("{uid}{SEP}{sheet}{SEP}{}", b.shape_id))
    }

    fn parse_shape(id: &str) -> Option<(u64, String, u32)> {
        let mut p = id.split(SEP);
        Some((p.next()?.parse().ok()?, p.next()?.to_string(), p.next()?.parse().ok()?))
    }

    fn shape(&self, id: &str) -> Result<(u64, String, FormButton), VbaError> {
        let (uid, sheet, sid) = Self::parse_shape(id).ok_or_else(|| err(1004, "shape"))?;
        let b = self
            .buttons_of(uid, &sheet)?
            .into_iter()
            .find(|b| b.shape_id == sid)
            .ok_or_else(|| err(1004, "The item with the specified name wasn't found."))?;
        Ok((uid, sheet, b))
    }

    /// `ws.Shapes(…)` / `ws.Buttons(…)`: by name or 1-based index.
    pub(super) fn shapes_get(&mut self, id: &str, member: &str, args: &[V]) -> Result<V, VbaError> {
        let (uid, sheet) = Self::parse_sheet(id).ok_or_else(|| err(1004, "shapes"))?;
        let list = self.buttons_of(uid, &sheet)?;
        match member.to_ascii_lowercase().as_str() {
            "count" => Ok(V::Long(list.len() as i32)),
            "item" | "_default" | "range" => {
                let which = super::arg(args, 0).ok_or_else(|| err(450, "Wrong number of arguments"))?;
                let b = match which {
                    V::String(n) => list.iter().find(|b| b.name.eq_ignore_ascii_case(n)),
                    other => usize::try_from(other.to_i64()? - 1).ok().and_then(|i| list.get(i)),
                };
                b.map(|b| Self::shape_handle(uid, &sheet, b)).ok_or_else(|| err(1004, "The item with the specified name wasn't found."))
            }
            _ => Err(err(438, format!("Object doesn't support this property or method: Shapes.{member}"))),
        }
    }

    pub(super) fn shapes_enumerate(&self, id: &str) -> Result<Vec<V>, VbaError> {
        let (uid, sheet) = Self::parse_sheet(id).ok_or_else(|| err(1004, "shapes"))?;
        Ok(self.buttons_of(uid, &sheet)?.iter().map(|b| Self::shape_handle(uid, &sheet, b)).collect())
    }

    /// The cell a point of the anchor (column, row) is in, as a range.
    fn anchor_cell(uid: u64, sheet: &str, col: u32, row: u32) -> V {
        let c = CellRef::new(row, col);
        range_value(&Rng { uid, sheet: sheet.to_string(), kind: 'c', areas: vec![RangeRef::new(c, c)] })
    }

    /// Left/top/right/bottom of a button in points.
    fn shape_box(&self, uid: u64, sheet: &str, b: &FormButton) -> Result<(f64, f64, f64, f64), VbaError> {
        let wb = self.wb(uid)?;
        let si = self.sheet_index(uid, sheet)?;
        let sh = wb.sheet(si).ok_or_else(|| err(9, "sheet"))?;
        let px = 0.75;
        let [c0, dx0, r0, dy0, c1, dx1, r1, dy1] = b.anchor;
        Ok((
            sh.col_left(c0) + f64::from(dx0) * px,
            sh.row_top(r0) + f64::from(dy0) * px,
            sh.col_left(c1) + f64::from(dx1) * px,
            sh.row_top(r1) + f64::from(dy1) * px,
        ))
    }

    pub(super) fn shape_get(&mut self, id: &str, member: &str) -> Result<V, VbaError> {
        let (uid, sheet, b) = self.shape(id)?;
        let m = member.to_ascii_lowercase();
        match m.as_str() {
            "name" => Ok(V::str(b.name.as_str())),
            "topleftcell" => Ok(Self::anchor_cell(uid, &sheet, b.anchor[0], b.anchor[2])),
            "bottomrightcell" => Ok(Self::anchor_cell(uid, &sheet, b.anchor[4], b.anchor[6])),
            "textframe" | "textframe2" | "drawingobject" => Ok(host("TextFrame", id)),
            "characters" => Ok(host("Characters", id)),
            "text" | "caption" => Ok(V::str(b.caption.as_str())),
            "onaction" => Ok(V::str(b.macro_name.as_str())),
            "left" | "top" | "width" | "height" => {
                let (l, t, r, bt) = self.shape_box(uid, &sheet, &b)?;
                Ok(V::Double(match m.as_str() {
                    "left" => l,
                    "top" => t,
                    "width" => r - l,
                    _ => bt - t,
                }))
            }
            "type" => Ok(V::Long(8)), // msoFormControl
            "visible" => Ok(V::Long(-1)),
            "select" => Ok(V::Empty),
            _ => Err(err(438, format!("Object doesn't support this property or method: Shape.{member}"))),
        }
    }

    pub(super) fn shape_set(&mut self, id: &str, member: &str, v: V) -> Result<(), VbaError> {
        match member.to_ascii_lowercase().as_str() {
            "text" | "caption" => self.set_caption(id, &v.to_str()?),
            "onaction" | "visible" | "locked" => Ok(()),
            _ => Err(err(438, format!("Object doesn't support this property or method: Shape.{member}"))),
        }
    }

    /// `TextFrame.Characters` (any arguments) and `TextFrame2.TextRange`.
    pub(super) fn text_frame_get(&mut self, id: &str, member: &str) -> Result<V, VbaError> {
        match member.to_ascii_lowercase().as_str() {
            "characters" | "textrange" => Ok(host("Characters", id)),
            "text" | "caption" => self.shape_get(id, "Text"),
            _ => Err(err(438, format!("Object doesn't support this property or method: TextFrame.{member}"))),
        }
    }

    pub(super) fn characters_get(&mut self, id: &str, member: &str) -> Result<V, VbaError> {
        match member.to_ascii_lowercase().as_str() {
            "text" | "caption" | "_default" => self.shape_get(id, "Text"),
            "count" => Ok(V::Long(self.shape(id)?.2.caption.chars().count() as i32)),
            _ => Err(err(438, format!("Object doesn't support this property or method: Characters.{member}"))),
        }
    }

    /// A button's new caption, written into the kept drawing so the file saves it.
    pub(super) fn set_caption(&mut self, id: &str, text: &str) -> Result<(), VbaError> {
        let (uid, sheet, sid) = Self::parse_shape(id).ok_or_else(|| err(1004, "shape"))?;
        let di = self.doc_index(uid)?;
        self.s.set_active(di);
        let si = self.sheet_index(uid, &sheet)?;
        let text = text.to_string();
        crate::cmd::edit(self.s, |cx| {
            let sh = cx.sheet_mut(si)?;
            let p = sh.preserved.as_mut().ok_or_else(|| crate::EngineError::Other("The button is gone.".into()))?;
            if !form_controls::set_caption(Arc::make_mut(p), sid, &text) {
                return Err(crate::EngineError::Other("The button is gone.".into()));
            }
            Ok(())
        })
        .map_err(super::engine_err)
    }

    // ------------------------------------------------------------ comments

    /// `Range.CommentThreaded` / `Range.Comment`: the first cell's, or Nothing.
    pub(super) fn range_comment(&self, r: &Rng, threaded: bool) -> Result<V, VbaError> {
        let wb = self.wb(r.uid)?;
        let si = self.sheet_index(r.uid, &r.sheet)?;
        let c = r.first().start;
        let has = wb.sheet(si).and_then(|s| s.comments.get(&c)).is_some_and(|x| x.threaded == threaded);
        let one = r.with('c', vec![RangeRef::new(c, c)]);
        Ok(if has { host(if threaded { "CommentThreaded" } else { "Comment" }, one.id()) } else { V::Nothing })
    }

    pub(super) fn add_comment(&mut self, r: &Rng, text: &str, threaded: bool) -> Result<V, VbaError> {
        let c = r.first().start;
        let author = self.s.prefs.user_name.clone();
        let cmd = if threaded { "review.newComment" } else { "review.newNote" };
        self.run(r.uid, cmd, json!({"sheet": r.sheet, "cell": c.a1(), "text": text, "author": author}))?;
        self.range_comment(r, threaded)
    }

    pub(super) fn comment_get(&mut self, id: &str, member: &str, args: &[V]) -> Result<V, VbaError> {
        let r = Rng::parse(id).ok_or_else(|| err(1004, "comment"))?;
        let c = r.first().start;
        let comment = {
            let wb = self.wb(r.uid)?;
            let si = self.sheet_index(r.uid, &r.sheet)?;
            wb.sheet(si).and_then(|s| s.comments.get(&c)).cloned()
        };
        let comment = comment.ok_or_else(|| err(91, "Object variable or With block variable not set"))?;
        match member.to_ascii_lowercase().as_str() {
            "text" if args.is_empty() => Ok(V::str(comment.text.as_str())),
            "text" => {
                // `Comment.Text Text:="…"` replaces the text.
                let t = super::arg(args, 0).map(|v| v.to_str()).transpose()?.unwrap_or_default();
                let threaded = comment.threaded;
                self.run(r.uid, "review.deleteComment", json!({"sheet": r.sheet, "cell": c.a1()}))?;
                self.add_comment(&r, &t, threaded)?;
                Ok(V::str(t))
            }
            "author" => Ok(V::str(comment.author.as_str())),
            "delete" => {
                self.run(r.uid, "review.deleteComment", json!({"sheet": r.sheet, "cell": c.a1()}))?;
                Ok(V::Empty)
            }
            "parent" => Ok(range_value(&r)),
            "visible" => Ok(V::Bool(comment.visible)),
            _ => Err(err(438, format!("Object doesn't support this property or method: Comment.{member}"))),
        }
    }

    // ------------------------------------------------------------ ADODB.Stream

    pub(super) fn new_stream(&mut self) -> V {
        self.next_stream += 1;
        let id = self.next_stream;
        self.streams.push((id, Stream { charset: "unicode".into(), ..Default::default() }));
        host("Stream", id.to_string())
    }

    fn stream(&mut self, id: &str) -> Result<&mut Stream, VbaError> {
        let id: u32 = id.parse().map_err(|_| err(3001, "stream"))?;
        self.streams
            .iter_mut()
            .find(|(i, _)| *i == id)
            .map(|(_, s)| s)
            .ok_or_else(|| err(3704, "Operation is not allowed when the object is closed."))
    }

    pub(super) fn stream_call(&mut self, id: &str, member: &str, args: &[V]) -> Result<V, VbaError> {
        let a0 = super::arg(args, 0).cloned();
        let m = member.to_ascii_lowercase();
        match m.as_str() {
            "open" => {
                let s = self.stream(id)?;
                s.open = true;
                s.text.clear();
                Ok(V::Empty)
            }
            "writetext" => {
                let t = a0.map(|v| v.to_str()).transpose()?.unwrap_or_default();
                let line = super::arg(args, 1).map(|v| v.to_i64()).transpose()?.unwrap_or(0) == 1;
                let s = self.stream(id)?;
                if !s.open {
                    return Err(err(3704, "Operation is not allowed when the object is closed."));
                }
                if s.text.len() + t.len() > 512 << 20 {
                    return Err(err(7, "Out of memory"));
                }
                s.text.push_str(&t);
                if line {
                    s.text.push_str("\r\n");
                }
                Ok(V::Empty)
            }
            "readtext" => Ok(V::str(self.stream(id)?.text.clone())),
            "savetofile" => {
                let path = a0.ok_or_else(|| err(450, "SaveToFile needs a file name"))?.to_str()?;
                let overwrite = super::arg(args, 1).map(|v| v.to_i64()).transpose()?.unwrap_or(1) == 2;
                let s = self.stream(id)?;
                let bytes = encode(&s.text, &s.charset);
                write_file(&path, &bytes, overwrite)?;
                Ok(V::Empty)
            }
            "close" => {
                let id: u32 = id.parse().unwrap_or(0);
                self.streams.retain(|(i, _)| *i != id);
                Ok(V::Empty)
            }
            "charset" => Ok(V::str(self.stream(id)?.charset.clone())),
            "size" => {
                let s = self.stream(id)?;
                Ok(V::Long(encode(&s.text, &s.charset).len() as i32))
            }
            "type" | "position" | "state" | "linesseparator" => Ok(V::Long(2)),
            _ => Err(err(438, format!("Object doesn't support this property or method: Stream.{member}"))),
        }
    }

    pub(super) fn stream_set(&mut self, id: &str, member: &str, v: V) -> Result<(), VbaError> {
        match member.to_ascii_lowercase().as_str() {
            "charset" => {
                let c = v.to_str()?;
                self.stream(id)?.charset = c;
                Ok(())
            }
            "type" | "position" | "linesseparator" | "mode" => Ok(()),
            _ => Err(err(438, format!("Object doesn't support this property or method: Stream.{member}"))),
        }
    }
}

/// Text in an ADODB charset; UTF-8 and Unicode get the byte-order mark ADODB writes.
fn encode(text: &str, charset: &str) -> Vec<u8> {
    match charset.to_ascii_lowercase().as_str() {
        "utf-8" | "utf8" => {
            let mut b = vec![0xEF, 0xBB, 0xBF];
            b.extend_from_slice(text.as_bytes());
            b
        }
        "ascii" | "us-ascii" => text.chars().map(|c| if c.is_ascii() { c as u8 } else { b'?' }).collect(),
        "iso-8859-1" | "windows-1252" | "latin1" => gridcraft_xlsx::ovba::encode(text, 1252),
        _ => {
            // "unicode" (the default): UTF-16LE with a BOM.
            let mut b = vec![0xFF, 0xFE];
            for u in text.encode_utf16() {
                b.extend_from_slice(&u.to_le_bytes());
            }
            b
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn io_err(e: std::io::Error, path: &str) -> VbaError {
    match e.kind() {
        std::io::ErrorKind::NotFound => err(53, format!("File not found: {path}")),
        std::io::ErrorKind::PermissionDenied => err(70, format!("Permission denied: {path}")),
        std::io::ErrorKind::AlreadyExists => err(58, format!("File already exists: {path}")),
        _ => err(75, format!("Path/File access error: {path} ({e})")),
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn write_file(path: &str, bytes: &[u8], overwrite: bool) -> Result<(), VbaError> {
    use std::io::Write as _;
    let mut o = std::fs::OpenOptions::new();
    o.write(true);
    if overwrite {
        o.create(true).truncate(true);
    } else {
        o.create_new(true);
    }
    let mut f = o.open(path).map_err(|e| if overwrite { io_err(e, path) } else { err(3004, format!("Write to file failed: {path}")) })?;
    f.write_all(bytes).map_err(|e| io_err(e, path))
}

#[cfg(target_arch = "wasm32")]
fn write_file(_: &str, _: &[u8], _: bool) -> Result<(), VbaError> {
    Err(err(75, "Files can't be written from the web version"))
}

/// Larger text files aren't opened for input.
#[cfg(not(target_arch = "wasm32"))]
const MAX_TEXT_FILE: u64 = 64 << 20;

/// What `Open … For Input` reads: UTF-8 (with or without its byte-order mark) or UTF-16 when the
/// file says so, else ANSI (Windows-1252), as VBA reads files.
#[cfg(not(target_arch = "wasm32"))]
pub(super) fn decode_text(bytes: &[u8]) -> String {
    if let Some(rest) = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]) {
        return String::from_utf8_lossy(rest).into_owned();
    }
    if let Some(rest) = bytes.strip_prefix(&[0xFF, 0xFE]) {
        let units: Vec<u16> = rest.chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]])).collect();
        return String::from_utf16_lossy(&units);
    }
    match std::str::from_utf8(bytes) {
        Ok(s) => s.to_string(),
        Err(_) => gridcraft_xlsx::ovba::decode(bytes, 1252),
    }
}

#[cfg(not(target_arch = "wasm32"))]
pub(super) fn read_text(path: &str) -> Result<String, VbaError> {
    use std::io::Read as _;
    let f = std::fs::File::open(path).map_err(|e| io_err(e, path))?;
    let meta = f.metadata().map_err(|e| io_err(e, path))?;
    if !meta.is_file() {
        return Err(err(75, format!("Path/File access error: {path}")));
    }
    let mut bytes = vec![];
    (&f).take(MAX_TEXT_FILE + 1).read_to_end(&mut bytes).map_err(|e| io_err(e, path))?;
    if bytes.len() as u64 > MAX_TEXT_FILE {
        return Err(err(7, format!("Out of memory: {path} is too large")));
    }
    Ok(decode_text(&bytes))
}

/// Writes what `Print #`/`Write #` wrote to a file, in ANSI (Windows-1252) as VBA does.
#[cfg(not(target_arch = "wasm32"))]
pub(super) fn write_text(path: &str, text: &str, append: bool) -> Result<(), VbaError> {
    use std::io::Write as _;
    let mut o = std::fs::OpenOptions::new();
    if append {
        o.append(true).create(true);
    } else {
        o.write(true).create(true).truncate(true);
    }
    let mut f = o.open(path).map_err(|e| io_err(e, path))?;
    f.write_all(&gridcraft_xlsx::ovba::encode(text, 1252)).map_err(|e| io_err(e, path))
}

#[cfg(target_arch = "wasm32")]
pub(super) fn read_text(_: &str) -> Result<String, VbaError> {
    Err(err(75, "Files can't be read from the web version"))
}

#[cfg(target_arch = "wasm32")]
pub(super) fn write_text(_: &str, _: &str, _: bool) -> Result<(), VbaError> {
    Err(err(75, "Files can't be written from the web version"))
}

/// `Dir(pattern)`: the first file name matching it (`*` and `?` in the name part), or "".
#[cfg(not(target_arch = "wasm32"))]
pub(super) fn dir(pattern: &str) -> Result<V, VbaError> {
    if pattern.is_empty() {
        return Ok(V::str(""));
    }
    let p = std::path::Path::new(pattern);
    let name = p.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default();
    if !name.contains(['*', '?']) {
        return Ok(V::str(if p.is_file() { name } else { String::new() }));
    }
    let folder = p.parent().filter(|f| !f.as_os_str().is_empty()).unwrap_or(std::path::Path::new("."));
    let Ok(rd) = std::fs::read_dir(folder) else { return Ok(V::str("")) };
    let mut names: Vec<String> = rd
        .filter_map(|e| e.ok())
        .filter(|e| e.path().is_file())
        .map(|e| e.file_name().to_string_lossy().to_string())
        .filter(|n| gridcraft_vba::value::like(n, &name, true))
        .collect();
    names.sort();
    Ok(V::str(names.into_iter().next().unwrap_or_default()))
}

#[cfg(not(target_arch = "wasm32"))]
pub(super) fn kill(pattern: &str) -> Result<(), VbaError> {
    let p = std::path::Path::new(pattern);
    let name = p.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default();
    if !name.contains(['*', '?']) {
        return std::fs::remove_file(p).map_err(|e| io_err(e, pattern));
    }
    let folder = p.parent().filter(|f| !f.as_os_str().is_empty()).unwrap_or(std::path::Path::new("."));
    let rd = std::fs::read_dir(folder).map_err(|e| io_err(e, pattern))?;
    let mut any = false;
    for e in rd.filter_map(|e| e.ok()) {
        if e.path().is_file() && gridcraft_vba::value::like(&e.file_name().to_string_lossy(), &name, true) {
            std::fs::remove_file(e.path()).map_err(|x| io_err(x, pattern))?;
            any = true;
        }
    }
    if any { Ok(()) } else { Err(err(53, format!("File not found: {pattern}"))) }
}

#[cfg(target_arch = "wasm32")]
pub(super) fn dir(_: &str) -> Result<V, VbaError> {
    Ok(V::str(""))
}

#[cfg(target_arch = "wasm32")]
pub(super) fn kill(p: &str) -> Result<(), VbaError> {
    Err(err(53, format!("File not found: {p}")))
}

/// `FileCopy`, `MkDir`, `RmDir` and `Name … As` are all paths.
#[cfg(not(target_arch = "wasm32"))]
pub(super) fn file_op(op: &str, a: &str, b: &str) -> Result<(), VbaError> {
    match op {
        "filecopy" => std::fs::copy(a, b).map(|_| ()).map_err(|e| io_err(e, a)),
        "mkdir" => std::fs::create_dir(a).map_err(|e| io_err(e, a)),
        "rmdir" => std::fs::remove_dir(a).map_err(|e| io_err(e, a)),
        _ => Ok(()),
    }
}

#[cfg(target_arch = "wasm32")]
pub(super) fn file_op(_: &str, a: &str, _: &str) -> Result<(), VbaError> {
    Err(err(75, format!("Path/File access error: {a}")))
}

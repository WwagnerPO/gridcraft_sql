//! Sequential text files: `Open … For Input | Output | Append As #n`, `Print #`, `Write #`,
//! `Input #`, `Line Input #`, `Close`, `Reset`, and `EOF`, `LOF`, `FreeFile` ([MS-VBAL] 5.4.5).
//!
//! The interpreter keeps the open files; the host reads and writes them ([`crate::Host`]), so this
//! crate never touches the disk itself. A file opened for output is written when it is closed
//! (or when the macro ends).

use std::collections::BTreeMap;

use crate::error::{VbaError, err};
use crate::value::{self, Value};

/// File numbers run from 1 to 511.
const MAX_FILE_NUMBER: i64 = 511;
/// `Print #` moves to the next of these columns after a comma.
const PRINT_ZONE: usize = 14;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FileMode {
    Input,
    Output,
    Append,
}

/// What follows an item of `Print #`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PrintSep {
    /// `;`: the next item follows at once.
    Semicolon,
    /// `,`: the next item starts at the next print zone.
    Comma,
    /// Nothing: the line ends.
    End,
}

struct OpenFile {
    path: String,
    mode: FileMode,
    /// Input: the file's text and where reading is.
    text: Vec<char>,
    pos: usize,
    /// Output and append: what was written, and the column the line is at.
    out: String,
    col: usize,
}

#[derive(Default)]
pub struct Files {
    open: BTreeMap<i64, OpenFile>,
}

fn bad_number() -> VbaError {
    err(52, "Bad file name or number")
}

fn past_end() -> VbaError {
    err(62, "Input past end of file")
}

/// A file number from a macro (`#1`, `f`).
pub fn number(v: &Value) -> Result<i64, VbaError> {
    let n = v.to_i64()?;
    if (1..=MAX_FILE_NUMBER).contains(&n) { Ok(n) } else { Err(bad_number()) }
}

/// Text `Print #` writes for a value: numbers with a space for the sign before them and one after.
fn print_text(v: &Value) -> Result<String, VbaError> {
    Ok(match v {
        Value::Byte(_) | Value::Integer(_) | Value::Long(_) | Value::LongLong(_) | Value::Single(_) | Value::Double(_) | Value::Currency(_) => {
            let s = v.to_str()?;
            if s.starts_with('-') { format!("{s} ") } else { format!(" {s} ") }
        }
        Value::Null => "Null".into(),
        other => other.to_str()?,
    })
}

/// A value as `Write #` writes it: strings quoted, `#TRUE#`, `#NULL#`, `#2026-10-30#`.
fn write_text(v: &Value) -> Result<String, VbaError> {
    Ok(match v {
        Value::Empty => String::new(),
        Value::Null => "#NULL#".into(),
        Value::Bool(b) => (if *b { "#TRUE#" } else { "#FALSE#" }).into(),
        Value::Error(n) => format!("#ERROR {n}#"),
        Value::Date(d) => format!("#{}#", iso_date(*d)),
        Value::String(s) => format!("\"{s}\""),
        other => other.to_str()?,
    })
}

/// `yyyy-mm-dd`, with ` hh:mm:ss` when there is a time.
fn iso_date(serial: f64) -> String {
    let days = serial.floor();
    // 1899-12-30 is day 0; civil-from-days after shifting to the Unix epoch (day 25569).
    let z = days as i64 - 25569 + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = yoe + era * 400 + i64::from(m <= 2);
    let secs = ((serial - days) * 86_400.0).round() as i64;
    if secs == 0 {
        format!("{y:04}-{m:02}-{d:02}")
    } else {
        format!("{y:04}-{m:02}-{d:02} {:02}:{:02}:{:02}", secs / 3600, secs / 60 % 60, secs % 60)
    }
}

/// A field `Input #` read: quoted text is a string, `#TRUE#`/`#FALSE#`/`#NULL#` what they say,
/// a number a Double, anything else text.
fn field_value(text: &str, quoted: bool) -> Value {
    if quoted {
        return Value::str(text);
    }
    match text.to_ascii_uppercase().as_str() {
        "#TRUE#" => Value::Bool(true),
        "#FALSE#" => Value::Bool(false),
        "#NULL#" => Value::Null,
        _ => match text.trim().parse::<f64>() {
            Ok(n) if n.is_finite() => Value::Double(n),
            _ => Value::str(text),
        },
    }
}

impl Files {
    pub fn is_open(&self, n: i64) -> bool {
        self.open.contains_key(&n)
    }

    /// `FreeFile`: the lowest number not in use.
    pub fn free(&self) -> Result<i64, VbaError> {
        (1..=MAX_FILE_NUMBER).find(|n| !self.open.contains_key(n)).ok_or_else(|| err(67, "Too many files"))
    }

    /// Opens `path` as file `n`; `text` is its content for input.
    pub fn open(&mut self, n: i64, path: &str, mode: FileMode, text: Option<String>) -> Result<(), VbaError> {
        if self.open.contains_key(&n) {
            return Err(err(55, "File already open"));
        }
        let text: Vec<char> = text.unwrap_or_default().chars().collect();
        self.open.insert(n, OpenFile { path: path.to_string(), mode, text, pos: 0, out: String::new(), col: 0 });
        Ok(())
    }

    fn writable(&mut self, n: i64) -> Result<&mut OpenFile, VbaError> {
        let f = self.open.get_mut(&n).ok_or_else(bad_number)?;
        if f.mode == FileMode::Input {
            return Err(err(54, "Bad file mode"));
        }
        Ok(f)
    }

    fn readable(&mut self, n: i64) -> Result<&mut OpenFile, VbaError> {
        let f = self.open.get_mut(&n).ok_or_else(bad_number)?;
        if f.mode != FileMode::Input {
            return Err(err(54, "Bad file mode"));
        }
        Ok(f)
    }

    fn push(f: &mut OpenFile, s: &str) -> Result<(), VbaError> {
        value::check_length(f.out.len().saturating_add(s.len()))?;
        f.out.push_str(s);
        match s.rfind('\n') {
            Some(i) => f.col = s.get(i + 1..).map_or(0, |t| t.chars().count()),
            None => f.col += s.chars().count(),
        }
        Ok(())
    }

    /// `Print #n, items`: each value with what follows it.
    pub fn print(&mut self, n: i64, items: &[(Value, PrintSep)]) -> Result<(), VbaError> {
        let f = self.writable(n)?;
        let mut ends_line = true;
        for (v, sep) in items {
            Self::push(f, &print_text(v)?)?;
            ends_line = *sep == PrintSep::End;
            if *sep == PrintSep::Comma {
                let pad = PRINT_ZONE - f.col % PRINT_ZONE;
                Self::push(f, &" ".repeat(pad))?;
            }
        }
        if ends_line {
            Self::push(f, "\r\n")?;
        }
        Ok(())
    }

    /// `Write #n, items`: comma-separated, strings quoted, then the line ends.
    pub fn write(&mut self, n: i64, items: &[Value]) -> Result<(), VbaError> {
        let f = self.writable(n)?;
        let fields: Vec<String> = items.iter().map(write_text).collect::<Result<_, _>>()?;
        Self::push(f, &fields.join(","))?;
        Self::push(f, "\r\n")
    }

    /// `Line Input #n`: the next line, without its end.
    pub fn line_input(&mut self, n: i64) -> Result<String, VbaError> {
        let f = self.readable(n)?;
        if f.pos >= f.text.len() {
            return Err(past_end());
        }
        let mut line = String::new();
        while let Some(&c) = f.text.get(f.pos) {
            f.pos += 1;
            match c {
                '\n' => break,
                '\r' => {
                    if f.text.get(f.pos) == Some(&'\n') {
                        f.pos += 1;
                    }
                    break;
                }
                c => line.push(c),
            }
        }
        Ok(line)
    }

    /// `Input #n`: the next field (comma, or line end, separated).
    pub fn input_field(&mut self, n: i64) -> Result<Value, VbaError> {
        let f = self.readable(n)?;
        let at = |f: &OpenFile| f.text.get(f.pos).copied();
        // Blank space and empty lines before a field.
        while matches!(at(f), Some(' ' | '\t' | '\r' | '\n')) {
            f.pos += 1;
        }
        if f.pos >= f.text.len() {
            return Err(past_end());
        }
        let mut text = String::new();
        let quoted = at(f) == Some('"');
        if quoted {
            f.pos += 1;
            while let Some(c) = at(f) {
                f.pos += 1;
                if c == '"' {
                    break;
                }
                text.push(c);
            }
        }
        // Up to the separator (a quoted field ignores what follows its closing quote).
        while let Some(c) = at(f) {
            if matches!(c, ',' | '\r' | '\n') {
                break;
            }
            if !quoted {
                text.push(c);
            }
            f.pos += 1;
        }
        match at(f) {
            Some(',') => f.pos += 1,
            Some('\r') => {
                f.pos += 1;
                if at(f) == Some('\n') {
                    f.pos += 1;
                }
            }
            Some('\n') => f.pos += 1,
            _ => {}
        }
        let text = if quoted { text } else { text.trim().to_string() };
        Ok(field_value(&text, quoted))
    }

    /// `EOF(n)`: no more to read.
    pub fn eof(&self, n: i64) -> Result<bool, VbaError> {
        let f = self.open.get(&n).ok_or_else(bad_number)?;
        Ok(f.mode != FileMode::Input || f.pos >= f.text.len())
    }

    /// `LOF(n)`: the file's length (what was written so far for output).
    pub fn lof(&self, n: i64) -> Result<i64, VbaError> {
        let f = self.open.get(&n).ok_or_else(bad_number)?;
        let len = if f.mode == FileMode::Input { f.text.iter().map(|c| c.len_utf8()).sum() } else { f.out.len() };
        Ok(i64::try_from(len).unwrap_or(i64::MAX))
    }

    /// Closes file `n`: what to write, `(path, text, append)`, for output.
    pub fn close(&mut self, n: i64) -> Result<Option<(String, String, bool)>, VbaError> {
        let f = self.open.remove(&n).ok_or_else(bad_number)?;
        Ok(match f.mode {
            FileMode::Input => None,
            FileMode::Output => Some((f.path, f.out, false)),
            FileMode::Append => Some((f.path, f.out, true)),
        })
    }

    /// The numbers of the open files.
    pub fn numbers(&self) -> Vec<i64> {
        self.open.keys().copied().collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn print_and_write_as_vba_does() {
        let mut f = Files::default();
        f.open(1, "x", FileMode::Output, None).unwrap();
        f.print(1, &[(Value::str("a"), PrintSep::Semicolon), (Value::Integer(5), PrintSep::Semicolon), (Value::Integer(-2), PrintSep::End)]).unwrap();
        f.print(1, &[(Value::str("ab"), PrintSep::Comma), (Value::str("c"), PrintSep::End)]).unwrap();
        f.print(1, &[(Value::str("no end"), PrintSep::Semicolon)]).unwrap();
        f.write(1, &[Value::str("x,y"), Value::Double(1.5), Value::Bool(true), Value::Null, Value::Date(45_960.5)]).unwrap();
        let (path, text, append) = f.close(1).unwrap().unwrap();
        assert_eq!((path.as_str(), append), ("x", false));
        assert_eq!(text, "a 5 -2 \r\nab            c\r\nno end\"x,y\",1.5,#TRUE#,#NULL#,#2025-10-30 12:00:00#\r\n");
    }

    #[test]
    fn input_reads_fields_and_lines() {
        let mut f = Files::default();
        f.open(2, "y", FileMode::Input, Some("\"Ana, M\",  42 ,#TRUE#\r\nlast line\nz".into())).unwrap();
        let show = |v: Value| format!("{} {}", v.type_name(), v.to_str().unwrap());
        assert_eq!(show(f.input_field(2).unwrap()), "String Ana, M");
        assert_eq!(show(f.input_field(2).unwrap()), "Double 42");
        assert_eq!(show(f.input_field(2).unwrap()), "Boolean True");
        assert!(!f.eof(2).unwrap());
        assert_eq!(f.line_input(2).unwrap(), "last line");
        assert_eq!(f.line_input(2).unwrap(), "z");
        assert!(f.eof(2).unwrap());
        assert_eq!(f.line_input(2).unwrap_err().number, 62);
        assert_eq!(f.write(2, &[]).unwrap_err().number, 54);
    }

    #[test]
    fn numbers_are_checked() {
        let mut f = Files::default();
        assert_eq!(number(&Value::Integer(0)).unwrap_err().number, 52);
        assert_eq!(number(&Value::Long(512)).unwrap_err().number, 52);
        f.open(1, "a", FileMode::Output, None).unwrap();
        assert_eq!(f.open(1, "b", FileMode::Output, None).unwrap_err().number, 55);
        assert_eq!(f.free().unwrap(), 2);
        assert_eq!(f.close(3).unwrap_err().number, 52);
    }
}

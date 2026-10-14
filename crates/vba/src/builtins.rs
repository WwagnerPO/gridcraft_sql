//! VBA's built-in functions and statements ([MS-VBAL] 6.1.1), the ones independent of the host:
//! text, maths, dates, type conversions and tests, plus the Collection and (Scripting) Dictionary
//! objects. Host functions (MsgBox, cells, …) are supplied by the embedder.

use std::cell::RefCell;
use std::rc::Rc;

use crate::error::{VbaError, err};
use crate::value::{self, Value};

/// The `Collection` object: 1-based, optional string keys.
#[derive(Debug, Default)]
pub struct Collection {
    pub items: Vec<Value>,
    pub keys: Vec<Option<String>>,
}

impl Collection {
    pub fn add(&mut self, item: Value, key: Option<String>, before: Option<i64>, after: Option<i64>) -> Result<(), VbaError> {
        if let Some(k) = &key
            && self.keys.iter().flatten().any(|x| x.eq_ignore_ascii_case(k))
        {
            return Err(err(457, "This key is already associated with an element of this collection"));
        }
        let at = match (before, after) {
            (Some(b), _) => (b - 1).max(0) as usize,
            (_, Some(a)) => a as usize,
            _ => self.items.len(),
        }
        .min(self.items.len());
        self.items.insert(at, item);
        self.keys.insert(at, key.map(|k| k.to_lowercase()));
        Ok(())
    }
    fn position(&self, index: &Value) -> Result<usize, VbaError> {
        if let Value::String(s) = index {
            let s = s.to_lowercase();
            return self.keys.iter().position(|k| k.as_deref() == Some(s.as_str())).ok_or_else(|| err(5, "Invalid procedure call or argument"));
        }
        let i = index.to_i64()?;
        if i < 1 || i as usize > self.items.len() {
            return Err(err(9, "Subscript out of range"));
        }
        Ok((i - 1) as usize)
    }
    pub fn item(&self, index: &Value) -> Result<Value, VbaError> {
        Ok(self.items[self.position(index)?].clone())
    }
    pub fn remove(&mut self, index: &Value) -> Result<(), VbaError> {
        let at = self.position(index)?;
        self.items.remove(at);
        self.keys.remove(at);
        Ok(())
    }
}

/// The `Scripting.Dictionary` object.
#[derive(Debug, Default)]
pub struct Dictionary {
    pub pairs: Vec<(Value, Value)>,
    /// 0 binary, 1 text.
    pub compare: i32,
}

impl Dictionary {
    fn find(&self, key: &Value) -> Option<usize> {
        self.pairs.iter().position(|(k, _)| key_eq(k, key, self.compare == 1))
    }
    pub fn get(&self, key: &Value) -> Value {
        self.find(key).map(|i| self.pairs[i].1.clone()).unwrap_or(Value::Empty)
    }
    pub fn set(&mut self, key: Value, val: Value) {
        match self.find(&key) {
            Some(i) => self.pairs[i].1 = val,
            None => self.pairs.push((key, val)),
        }
    }
    pub fn exists(&self, key: &Value) -> bool {
        self.find(key).is_some()
    }
    pub fn remove(&mut self, key: &Value) -> Result<(), VbaError> {
        let i = self.find(key).ok_or_else(|| err(32811, "Element not found"))?;
        self.pairs.remove(i);
        Ok(())
    }
}

fn key_eq(a: &Value, b: &Value, text: bool) -> bool {
    value::compare(a, b, text).map(|o| o == std::cmp::Ordering::Equal).unwrap_or(false)
}

/// Now, as a date serial (local time unavailable off the host: UTC).
pub fn now() -> f64 {
    gridcraft_core::date::unix_now_ms().map(|ms| ms / 86_400_000.0 + 25_569.0).unwrap_or(0.0)
}

fn arg(args: &[Value], i: usize) -> Value {
    args.get(i).cloned().unwrap_or(Value::Empty)
}

/// Calls a built-in function by name (upper-cased). `Ok(None)` when the name isn't a built-in, so
/// the caller can look elsewhere (user procedures, host functions).
pub fn call(name: &str, args: &[Value], text_compare: bool) -> Result<Option<Value>, VbaError> {
    let upper = name.to_ascii_uppercase();
    let s = |i: usize| arg(args, i).to_str();
    let n = |i: usize| arg(args, i).to_f64();
    let int = |i: usize| arg(args, i).to_i64();
    let v = Value::str;
    let r = match upper.as_str() {
        // ---- text
        "LEN" => match arg(args, 0) {
            Value::Null => Value::Null,
            other => Value::from_i64(other.to_str()?.chars().count() as i64),
        },
        "LEFT" | "LEFT$" => {
            let t = s(0)?;
            let k = int(1)?.max(0) as usize;
            v(t.chars().take(k).collect::<String>())
        }
        "RIGHT" | "RIGHT$" => {
            let t: Vec<char> = s(0)?.chars().collect();
            let k = (int(1)?.max(0) as usize).min(t.len());
            v(t[t.len() - k..].iter().collect::<String>())
        }
        "MID" | "MID$" => {
            let t: Vec<char> = s(0)?.chars().collect();
            let start = int(1)?;
            if start < 1 {
                return Err(err(5, "Invalid procedure call or argument"));
            }
            let start = (start - 1) as usize;
            let len = if args.len() > 2 && !matches!(arg(args, 2), Value::Empty) { int(2)?.max(0) as usize } else { t.len() };
            v(t.iter().skip(start).take(len).collect::<String>())
        }
        "UCASE" | "UCASE$" => v(s(0)?.to_uppercase()),
        "LCASE" | "LCASE$" => v(s(0)?.to_lowercase()),
        "TRIM" | "TRIM$" => v(s(0)?.trim_matches(' ').to_string()),
        "LTRIM" | "LTRIM$" => v(s(0)?.trim_start_matches(' ').to_string()),
        "RTRIM" | "RTRIM$" => v(s(0)?.trim_end_matches(' ').to_string()),
        "SPACE" | "SPACE$" => v(" ".repeat(int(0)?.max(0) as usize)),
        "STRING" | "STRING$" => {
            let c = match arg(args, 1) {
                Value::String(s) => s.chars().next().unwrap_or(' '),
                other => char::from_u32(other.to_i64()? as u32).unwrap_or(' '),
            };
            v(c.to_string().repeat(int(0)?.max(0) as usize))
        }
        "INSTR" => instr(args, false, text_compare)?,
        "INSTRREV" => instr(args, true, text_compare)?,
        "REPLACE" => {
            let t = s(0)?;
            let find = s(1)?;
            let repl = s(2)?;
            if find.is_empty() {
                v(t)
            } else {
                let compare = if args.len() > 5 { int(5)? == 1 } else { text_compare };
                v(replace(&t, &find, &repl, compare))
            }
        }
        "SPLIT" => {
            let t = s(0)?;
            let sep = if args.len() > 1 && !matches!(arg(args, 1), Value::Empty) { s(1)? } else { " ".into() };
            let parts: Vec<Value> = if sep.is_empty() { vec![v(t)] } else { t.split(&sep).map(|x| v(x.to_string())).collect() };
            array_1d(parts)
        }
        "JOIN" => {
            let a = as_array(&arg(args, 0))?;
            let sep = if args.len() > 1 { s(1)? } else { " ".into() };
            let parts: Result<Vec<String>, _> = a.borrow().items.iter().map(|x| x.to_str()).collect();
            v(parts?.join(&sep))
        }
        "STRREVERSE" => v(s(0)?.chars().rev().collect::<String>()),
        "STRCOMP" => {
            let o = value::compare(&Value::str(s(0)?), &Value::str(s(1)?), if args.len() > 2 { int(2)? == 1 } else { text_compare })?;
            Value::Integer(match o {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1,
            })
        }
        "ASC" => {
            let t = s(0)?;
            Value::from_i64(t.chars().next().map(|c| c as i64).unwrap_or(0))
        }
        "CHR" | "CHR$" | "CHRW" => v(char::from_u32(int(0)? as u32).map(|c| c.to_string()).unwrap_or_default()),
        "FORMAT" | "FORMAT$" => v(format_value(&arg(args, 0), &s(1).unwrap_or_default())?),
        "LSET" | "RSET" => v(s(0)?),
        // ---- maths
        "ABS" => match arg(args, 0) {
            Value::Null => Value::Null,
            x => num_same(&x, n(0)?.abs()),
        },
        "SGN" => Value::Integer(n(0)?.signum() as i16 * (n(0)? != 0.0) as i16),
        "INT" => num_same(&arg(args, 0), n(0)?.floor()),
        "FIX" => num_same(&arg(args, 0), n(0)?.trunc()),
        "SQR" => {
            let x = n(0)?;
            if x < 0.0 {
                return Err(err(5, "Invalid procedure call or argument"));
            }
            Value::Double(x.sqrt())
        }
        "SIN" => Value::Double(n(0)?.sin()),
        "COS" => Value::Double(n(0)?.cos()),
        "TAN" => Value::Double(n(0)?.tan()),
        "ATN" => Value::Double(n(0)?.atan()),
        "EXP" => Value::Double(n(0)?.exp()),
        "LOG" => {
            let x = n(0)?;
            if x <= 0.0 {
                return Err(err(5, "Invalid procedure call or argument"));
            }
            Value::Double(x.ln())
        }
        "ROUND" => {
            let x = n(0)?;
            let d = if args.len() > 1 { int(1)?.max(0) as i32 } else { 0 };
            let f = 10f64.powi(d);
            Value::Double(value::round_even(x * f) / f)
        }
        "RND" => Value::Single(pseudo_random()),
        // ---- type conversions
        "CINT" => Value::Integer(i16::try_from(arg(args, 0).to_i64()?).map_err(|_| err(6, "Overflow"))?),
        "CLNG" => Value::Long(arg(args, 0).to_i32()?),
        "CLNGLNG" => Value::LongLong(arg(args, 0).to_i64()?),
        "CBYTE" => Value::Byte(u8::try_from(arg(args, 0).to_i64()?).map_err(|_| err(6, "Overflow"))?),
        "CDBL" => Value::Double(arg(args, 0).to_f64()?),
        "CSNG" => Value::Single(arg(args, 0).to_f64()? as f32),
        "CCUR" => Value::Currency((arg(args, 0).to_f64()? * 10_000.0).round() as i64),
        "CBOOL" => Value::Bool(arg(args, 0).to_bool()?),
        "CSTR" => v(arg(args, 0).to_str()?),
        "CDATE" | "CVDATE" => Value::Date(arg(args, 0).to_date()?),
        "CVAR" => arg(args, 0),
        "VAL" => Value::Double(val_number(&s(0)?)),
        "STR" | "STR$" => {
            let t = arg(args, 0).to_str()?;
            v(if t.starts_with('-') { t } else { format!(" {t}") })
        }
        "HEX" | "HEX$" => v(format!("{:X}", int(0)? & 0xFFFF_FFFF)),
        "OCT" | "OCT$" => v(format!("{:o}", int(0)? & 0xFFFF_FFFF)),
        // ---- type tests
        "ISNUMERIC" => Value::Bool(match arg(args, 0) {
            Value::String(s) => value::parse_number(&s).is_some(),
            x => x.is_numeric_type() || matches!(x, Value::Empty),
        }),
        "ISDATE" => Value::Bool(match arg(args, 0) {
            Value::Date(_) => true,
            Value::String(s) => value::parse_date(&s).is_some(),
            _ => false,
        }),
        "ISEMPTY" => Value::Bool(matches!(arg(args, 0), Value::Empty)),
        "ISNULL" => Value::Bool(matches!(arg(args, 0), Value::Null)),
        "ISARRAY" => Value::Bool(matches!(arg(args, 0), Value::Array(_))),
        "ISOBJECT" => Value::Bool(arg(args, 0).is_object()),
        "ISERROR" => Value::Bool(matches!(arg(args, 0), Value::Error(_))),
        "VARTYPE" => Value::Integer(arg(args, 0).var_type() as i16),
        "TYPENAME" => v(arg(args, 0).type_name()),
        "CVERR" => Value::Error(int(0)? as i32),
        // ---- dates
        "NOW" => Value::Date(now()),
        "DATE" => Value::Date(now().floor()),
        "TIME" => Value::Date(now().fract()),
        "TIMER" => Value::Single((now().fract() * 86_400.0) as f32),
        "YEAR" => Value::Integer(value::date_parts(arg(args, 0).to_date()?).0 as i16),
        "MONTH" => Value::Integer(value::date_parts(arg(args, 0).to_date()?).1 as i16),
        "DAY" => Value::Integer(value::date_parts(arg(args, 0).to_date()?).2 as i16),
        "HOUR" => Value::Integer(value::date_parts(arg(args, 0).to_date()?).3 as i16),
        "MINUTE" => Value::Integer(value::date_parts(arg(args, 0).to_date()?).4 as i16),
        "SECOND" => Value::Integer(value::date_parts(arg(args, 0).to_date()?).5 as i16),
        "WEEKDAY" => {
            let first = if args.len() > 1 && int(1)? != 0 { int(1)? } else { 1 };
            let dow = (arg(args, 0).to_date()?.floor() as i64 + 6).rem_euclid(7); // 0 = Sunday
            Value::Integer(((dow - (first - 1)).rem_euclid(7) + 1) as i16)
        }
        "DATESERIAL" => Value::Date(value::date_serial(int(0)?, int(1)?, int(2)?).ok_or_else(|| err(5, "Invalid procedure call or argument"))?),
        "TIMESERIAL" => Value::Date(gridcraft_core::date::time_fraction(n(0)?, n(1)?, n(2)?)),
        "DATEVALUE" => Value::Date(value::parse_date(&s(0)?).ok_or_else(|| err(13, "Type mismatch"))?.floor()),
        "TIMEVALUE" => Value::Date(value::parse_date(&s(0)?).ok_or_else(|| err(13, "Type mismatch"))?.fract()),
        "DATEADD" => return date_add(&s(0)?, int(1)?, arg(args, 2).to_date()?).map(Some),
        "DATEDIFF" => return date_diff(&s(0)?, arg(args, 1).to_date()?, arg(args, 2).to_date()?).map(Some),
        "DATEPART" => return date_part(&s(0)?, arg(args, 1).to_date()?).map(Some),
        "MONTHNAME" => {
            let m = int(0)?;
            let names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
            v(names.get((m - 1) as usize).copied().unwrap_or("").to_string())
        }
        "WEEKDAYNAME" => {
            let d = int(0)?;
            let names = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
            v(names.get((d - 1).rem_euclid(7) as usize).copied().unwrap_or("").to_string())
        }
        // ---- choice
        "IIF" => {
            if arg(args, 0).to_bool()? {
                arg(args, 1)
            } else {
                arg(args, 2)
            }
        }
        "CHOOSE" => {
            let i = int(0)?;
            if i < 1 || i as usize >= args.len() { Value::Null } else { arg(args, i as usize) }
        }
        "SWITCH" => {
            let mut i = 0;
            let mut out = Value::Null;
            while i + 1 < args.len() {
                if arg(args, i).to_bool()? {
                    out = arg(args, i + 1);
                    break;
                }
                i += 2;
            }
            out
        }
        // ---- arrays
        "ARRAY" => array_1d(args.to_vec()),
        "UBOUND" => {
            let a = as_array(&arg(args, 0))?;
            let d = if args.len() > 1 { int(1)? as usize } else { 1 };
            Value::from_i64(a.borrow().bounds.get(d - 1).map(|b| b.1).ok_or_else(|| err(9, "Subscript out of range"))?)
        }
        "LBOUND" => {
            let a = as_array(&arg(args, 0))?;
            let d = if args.len() > 1 { int(1)? as usize } else { 1 };
            Value::from_i64(a.borrow().bounds.get(d - 1).map(|b| b.0).ok_or_else(|| err(9, "Subscript out of range"))?)
        }
        "FILTER" => {
            let a = as_array(&arg(args, 0))?;
            let needle = s(1)?;
            let include = if args.len() > 2 { arg(args, 2).to_bool()? } else { true };
            let text = if args.len() > 3 { int(3)? == 1 } else { text_compare };
            let hay = |x: &str| if text { x.to_lowercase().contains(&needle.to_lowercase()) } else { x.contains(&needle) };
            let out: Vec<Value> = a.borrow().items.iter().filter_map(|x| x.to_str().ok()).filter(|x| hay(x) == include).map(v).collect();
            array_1d(out)
        }
        // ---- misc, host-independent
        "RGB" => Value::Long((int(0)?.clamp(0, 255) | (int(1)?.clamp(0, 255) << 8) | (int(2)?.clamp(0, 255) << 16)) as i32),
        "QBCOLOR" => Value::Long(
            [
                0, 0x80_0000, 0x00_8000, 0x80_8000, 0x00_0080, 0x80_0080, 0x00_8080, 0xC0_C0C0, 0x80_8080, 0xFF_0000, 0x00_FF00, 0xFF_FF00,
                0x00_00FF, 0xFF_00FF, 0x00_FFFF, 0xFF_FFFF,
            ]
            .get(int(0)? as usize & 15)
            .copied()
            .unwrap_or(0),
        ),
        _ => return Ok(None),
    };
    Ok(Some(r))
}

/// The built-in names that take `ByRef` effects or are statements, handled by the interpreter.
pub fn is_builtin(name: &str) -> bool {
    const EXTRA: &[&str] = &["MSGBOX", "INPUTBOX", "CREATEOBJECT", "GETOBJECT", "TYPEOF", "DOEVENTS"];
    let u = name.to_ascii_uppercase();
    EXTRA.contains(&u.as_str())
        || call(name, &[], false).map(|r| r.is_some()).unwrap_or(false)
        || matches!(u.as_str(), "LEFT" | "RIGHT" | "MID" | "INSTR")
}

fn num_same(orig: &Value, r: f64) -> Value {
    match orig {
        Value::Integer(_) | Value::Long(_) | Value::LongLong(_) | Value::Byte(_) => Value::from_i64(r as i64),
        Value::Single(_) => Value::Single(r as f32),
        Value::Currency(_) => Value::Currency((r * 10_000.0).round() as i64),
        _ => Value::Double(r),
    }
}

fn instr(args: &[Value], rev: bool, text_compare: bool) -> Result<Value, VbaError> {
    // InStr([start], string1, string2, [compare]); InStrRev(string, find, [start], [compare]).
    let (start, hay, needle, compare) = if rev {
        let hay = arg(args, 0).to_str()?;
        let find = arg(args, 1).to_str()?;
        let start = if args.len() > 2 && !matches!(arg(args, 2), Value::Empty) { arg(args, 2).to_i64()? } else { -1 };
        let compare = if args.len() > 3 { arg(args, 3).to_i64()? == 1 } else { text_compare };
        (start, hay, find, compare)
    } else if matches!(arg(args, 0), Value::String(_) | Value::Null) || args.len() < 3 {
        (1, arg(args, 0).to_str().unwrap_or_default(), arg(args, 1).to_str()?, text_compare)
    } else {
        let compare = if args.len() > 3 { arg(args, 3).to_i64()? == 1 } else { text_compare };
        (arg(args, 0).to_i64()?, arg(args, 1).to_str()?, arg(args, 2).to_str()?, compare)
    };
    if matches!(args.first(), Some(Value::Null)) || matches!(args.get(1), Some(Value::Null)) {
        return Ok(Value::Null);
    }
    let h: Vec<char> = hay.chars().collect();
    let nd: Vec<char> = needle.chars().collect();
    let found = find_sub(&h, &nd, start, rev, compare);
    Ok(Value::from_i64(found.map(|i| i as i64 + 1).unwrap_or(0)))
}

fn find_sub(hay: &[char], needle: &[char], start: i64, rev: bool, text: bool) -> Option<usize> {
    let eq = |a: char, b: char| if text { a.eq_ignore_ascii_case(&b) || a.to_lowercase().eq(b.to_lowercase()) } else { a == b };
    let matches_at = |i: usize| needle.iter().enumerate().all(|(k, &c)| hay.get(i + k).is_some_and(|&h| eq(h, c)));
    if needle.is_empty() {
        return Some(if rev { hay.len().saturating_sub(1) } else { (start - 1).max(0) as usize });
    }
    if rev {
        let from = if start < 0 { hay.len() } else { (start as usize).min(hay.len()) };
        (0..from.saturating_sub(needle.len()) + 1).rev().find(|&i| matches_at(i))
    } else {
        let from = (start - 1).max(0) as usize;
        (from..=hay.len().saturating_sub(needle.len())).find(|&i| matches_at(i))
    }
}

fn replace(text: &str, find: &str, repl: &str, text_compare: bool) -> String {
    if !text_compare {
        return text.replace(find, repl);
    }
    let (lower_text, lower_find) = (text.to_lowercase(), find.to_lowercase());
    let mut out = String::new();
    let chars: Vec<char> = text.chars().collect();
    let lc: Vec<char> = lower_text.chars().collect();
    let fc: Vec<char> = lower_find.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if i + fc.len() <= lc.len() && lc[i..i + fc.len()] == fc[..] {
            out.push_str(repl);
            i += fc.len();
        } else {
            out.push(chars[i]);
            i += 1;
        }
    }
    out
}

pub fn array_1d(items: Vec<Value>) -> Value {
    let hi = items.len() as i64 - 1;
    Value::Array(Rc::new(RefCell::new(crate::value::Array { bounds: vec![(0, hi)], items, fixed: false })))
}

fn as_array(v: &Value) -> Result<Rc<RefCell<crate::value::Array>>, VbaError> {
    match v {
        Value::Array(a) => Ok(a.clone()),
        _ => Err(err(13, "Type mismatch")),
    }
}

/// `Val`: leading numeric part of a string, spaces ignored.
fn val_number(s: &str) -> f64 {
    let cleaned: String = s.chars().filter(|c| !c.is_whitespace()).collect();
    if let Some(h) = cleaned.strip_prefix("&H").or_else(|| cleaned.strip_prefix("&h")) {
        return i64::from_str_radix(h.trim_end_matches(|c: char| !c.is_ascii_hexdigit()), 16).unwrap_or(0) as f64;
    }
    let mut end = 0;
    let bytes = cleaned.as_bytes();
    let mut seen_dot = false;
    let mut seen_e = false;
    for (i, &b) in bytes.iter().enumerate() {
        match b {
            b'0'..=b'9' => end = i + 1,
            b'+' | b'-' if i == 0 || matches!(bytes.get(i - 1), Some(b'e' | b'E')) => end = i + 1,
            b'.' if !seen_dot && !seen_e => {
                seen_dot = true;
                end = i + 1;
            }
            b'e' | b'E' if !seen_e && i > 0 => {
                seen_e = true;
                end = i + 1;
            }
            _ => break,
        }
    }
    cleaned.get(..end).and_then(|x| x.parse().ok()).unwrap_or(0.0)
}

fn pseudo_random() -> f32 {
    use std::cell::Cell;
    thread_local!(static STATE: Cell<u64> = const { Cell::new(0x2545F4914F6CDD1D) });
    STATE.with(|s| {
        let mut x = s.get();
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        s.set(x);
        ((x >> 40) as f32) / (1u64 << 24) as f32
    })
}

// ---- Format()

fn format_value(v: &Value, pattern: &str) -> Result<String, VbaError> {
    if pattern.is_empty() {
        return v.to_str();
    }
    let named = match pattern.to_ascii_lowercase().as_str() {
        "general number" => return v.to_str(),
        "currency" => "$#,##0.00",
        "fixed" => "0.00",
        "standard" => "#,##0.00",
        "percent" => "0.00%",
        "scientific" => "0.00E+00",
        "yes/no" => return Ok(if v.to_bool()? { "Yes" } else { "No" }.into()),
        "true/false" => return Ok(if v.to_bool()? { "True" } else { "False" }.into()),
        "on/off" => return Ok(if v.to_bool()? { "On" } else { "Off" }.into()),
        "long date" => "dddd, mmmm d, yyyy",
        "medium date" | "short date" => "m/d/yyyy",
        "long time" => "h:mm:ss AM/PM",
        "medium time" => "h:mm AM/PM",
        "short time" => "hh:mm",
        _ => pattern,
    };
    let code = gridcraft_numfmt::NumberFormat::parse(named);
    let core = match v {
        Value::String(_) if !code.is_text() => return v.to_str(),
        Value::Date(d) => gridcraft_core::Value::number(*d),
        Value::Bool(b) => gridcraft_core::Value::number(if *b { -1.0 } else { 0.0 }),
        Value::Empty => gridcraft_core::Value::number(0.0),
        _ => match v.to_f64() {
            Ok(n) => gridcraft_core::Value::number(n),
            Err(_) => gridcraft_core::Value::text(v.to_str()?),
        },
    };
    Ok(gridcraft_numfmt::format_value(&core, &code, gridcraft_core::date::DateSystem::D1900).text)
}

fn interval_unit(interval: &str) -> Option<&'static str> {
    Some(match interval.to_ascii_lowercase().as_str() {
        "yyyy" => "y",
        "q" => "q",
        "m" => "m",
        "y" | "d" => "d",
        "w" | "ww" => "w",
        "h" => "h",
        "n" => "n",
        "s" => "s",
        _ => return None,
    })
}

fn date_add(interval: &str, number: i64, date: f64) -> Result<Value, VbaError> {
    let unit = interval_unit(interval).ok_or_else(|| err(5, "Invalid procedure call or argument"))?;
    let (y, m, d, _, _, _) = value::date_parts(date);
    let time = date - date.floor();
    Ok(Value::Date(match unit {
        "y" => date + number as f64,
        "w" => date + number as f64 * 7.0,
        "h" => date + number as f64 / 24.0,
        "n" => date + number as f64 / 1440.0,
        "s" => date + number as f64 / 86400.0,
        "m" => value::date_serial(y as i64, m as i64 + number, d as i64).ok_or_else(|| err(5, "bad date"))? + time,
        "q" => value::date_serial(y as i64, m as i64 + number * 3, d as i64).ok_or_else(|| err(5, "bad date"))? + time,
        _ => value::date_serial(y as i64 + number, m as i64, d as i64).ok_or_else(|| err(5, "bad date"))? + time,
    }))
}

fn date_diff(interval: &str, a: f64, b: f64) -> Result<Value, VbaError> {
    let unit = interval_unit(interval).ok_or_else(|| err(5, "Invalid procedure call or argument"))?;
    let (ay, am, ..) = value::date_parts(a);
    let (by, bm, ..) = value::date_parts(b);
    let diff = match unit {
        "d" | "w" => (b.floor() - a.floor()) / if unit == "w" { 7.0 } else { 1.0 },
        "h" => (b - a) * 24.0,
        "n" => (b - a) * 1440.0,
        "s" => (b - a) * 86400.0,
        "m" => ((by - ay) * 12 + bm as i32 - am as i32) as f64,
        "q" => (((by - ay) * 12 + bm as i32 - am as i32) / 3) as f64,
        _ => (by - ay) as f64,
    };
    Ok(Value::from_i64(diff.trunc() as i64))
}

fn date_part(interval: &str, date: f64) -> Result<Value, VbaError> {
    let (y, m, d, h, mi, s) = value::date_parts(date);
    Ok(Value::Integer(match interval.to_ascii_lowercase().as_str() {
        "yyyy" => y as i16,
        "q" => ((m - 1) / 3 + 1) as i16,
        "m" => m as i16,
        "d" => d as i16,
        "h" => h as i16,
        "n" => mi as i16,
        "s" => s as i16,
        "y" => value::date_parts(date).2 as i16,
        "w" => ((date.floor() as i64 + 6).rem_euclid(7) + 1) as i16,
        _ => return Err(err(5, "Invalid procedure call or argument")),
    }))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn c(name: &str, args: &[Value]) -> Value {
        call(name, args, false).unwrap().unwrap()
    }
    fn s(x: &str) -> Value {
        Value::str(x)
    }

    #[test]
    fn text() {
        assert_eq!(c("Left", &[s("Hello"), Value::Integer(2)]).to_str().unwrap(), "He");
        assert_eq!(c("Mid", &[s("Hello"), Value::Integer(2), Value::Integer(3)]).to_str().unwrap(), "ell");
        assert_eq!(c("Right", &[s("Hello"), Value::Integer(2)]).to_str().unwrap(), "lo");
        assert_eq!(c("InStr", &[s("abcabc"), s("bc")]).to_i64().unwrap(), 2);
        assert_eq!(c("InStr", &[Value::Integer(3), s("abcabc"), s("bc")]).to_i64().unwrap(), 5);
        assert_eq!(c("InStrRev", &[s("abcabc"), s("bc")]).to_i64().unwrap(), 5);
        assert_eq!(c("Replace", &[s("a-b-c"), s("-"), s("+")]).to_str().unwrap(), "a+b+c");
        assert_eq!(c("Trim", &[s("  x  ")]).to_str().unwrap(), "x");
        assert_eq!(c("UCase", &[s("año")]).to_str().unwrap(), "AÑO");
        assert_eq!(c("Len", &[s("café")]).to_i64().unwrap(), 4);
        let split = c("Split", &[s("a,b,c"), s(",")]);
        assert_eq!(c("UBound", std::slice::from_ref(&split)).to_i64().unwrap(), 2);
        assert_eq!(c("Join", &[split, s("-")]).to_str().unwrap(), "a-b-c");
    }

    #[test]
    fn maths_and_conversions() {
        assert_eq!(c("Abs", &[Value::Integer(-5)]).to_i64().unwrap(), 5);
        assert_eq!(c("Int", &[Value::Double(-2.5)]).to_f64().unwrap(), -3.0);
        assert_eq!(c("Fix", &[Value::Double(-2.5)]).to_f64().unwrap(), -2.0);
        assert_eq!(c("Round", &[Value::Double(2.5), Value::Integer(0)]).to_f64().unwrap(), 2.0);
        assert_eq!(c("Round", &[Value::Double(1.23456), Value::Integer(2)]).to_f64().unwrap(), 1.23);
        assert_eq!(c("CInt", &[Value::Double(2.5)]).to_i64().unwrap(), 2);
        assert_eq!(c("Val", &[s("  12.5abc")]).to_f64().unwrap(), 12.5);
        assert!(c("IsNumeric", &[s("1.5")]).to_bool().unwrap());
        assert!(!c("IsNumeric", &[s("x")]).to_bool().unwrap());
        assert_eq!(c("Hex", &[Value::Integer(255)]).to_str().unwrap(), "FF");
        assert_eq!(c("RGB", &[Value::Integer(255), Value::Integer(0), Value::Integer(0)]).to_i64().unwrap(), 255);
    }

    #[test]
    fn dates_and_choice() {
        let d = c("DateSerial", &[Value::Integer(2023), Value::Integer(3), Value::Integer(15)]);
        assert_eq!(d.to_f64().unwrap(), 45000.0);
        assert_eq!(c("Year", std::slice::from_ref(&d)).to_i64().unwrap(), 2023);
        assert_eq!(c("DateAdd", &[s("m"), Value::Integer(1), d.clone()]).to_f64().unwrap(), 45031.0);
        assert_eq!(c("DateDiff", &[s("d"), d.clone(), Value::Double(45010.0)]).to_i64().unwrap(), 10);
        assert_eq!(c("IIf", &[Value::Bool(true), s("yes"), s("no")]).to_str().unwrap(), "yes");
        assert_eq!(c("Choose", &[Value::Integer(2), s("a"), s("b"), s("c")]).to_str().unwrap(), "b");
        assert_eq!(c("Format", &[Value::Double(1234.5), s("#,##0.00")]).to_str().unwrap(), "1,234.50");
    }

    #[test]
    fn collection_and_dictionary() {
        let mut col = Collection::default();
        col.add(s("x"), Some("a".into()), None, None).unwrap();
        col.add(s("y"), None, None, None).unwrap();
        assert_eq!(col.item(&s("a")).unwrap().to_str().unwrap(), "x");
        assert_eq!(col.item(&Value::Integer(2)).unwrap().to_str().unwrap(), "y");
        assert_eq!(col.add(s("z"), Some("A".into()), None, None).unwrap_err().number, 457);
        let mut d = Dictionary::default();
        d.set(s("k"), Value::Integer(1));
        d.set(s("k"), Value::Integer(2));
        assert_eq!(d.get(&s("k")).to_i64().unwrap(), 2);
        assert!(d.exists(&s("k")));
    }
}

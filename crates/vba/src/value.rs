//! VBA values (the Variant and its subtypes, [MS-VBAL] 2.1 and 5.5), conversions and the
//! arithmetic, comparison and logical operators (5.6.9).

use std::cell::RefCell;
use std::fmt;
use std::rc::Rc;

use crate::error::{VbaError, err};

/// A host object handle (the Excel object model): its class and an opaque reference the host
/// understands. Hosts compare handles by `class` and `id`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HostRef {
    pub class: String,
    pub id: String,
}

/// An array: lower bounds and lengths per dimension, elements in row-major order.
#[derive(Clone, Debug)]
pub struct Array {
    pub bounds: Vec<(i64, i64)>,
    pub items: Vec<Value>,
    /// Declared with a fixed size (`Dim a(3)`): ReDim refuses it.
    pub fixed: bool,
}

impl Array {
    pub fn new(bounds: Vec<(i64, i64)>, fill: Value) -> Result<Array, VbaError> {
        let mut n: usize = 1;
        for (lo, hi) in &bounds {
            if hi < lo && !(hi + 1 == *lo) {
                return Err(err(9, "Subscript out of range"));
            }
            let len = usize::try_from((hi - lo + 1).max(0)).map_err(|_| err(7, "Out of memory"))?;
            n = n.checked_mul(len).ok_or_else(|| err(7, "Out of memory"))?;
        }
        if n > 50_000_000 {
            return Err(err(7, "Out of memory"));
        }
        Ok(Array { bounds, items: vec![fill; n], fixed: false })
    }

    /// The flat index of `idx`, checked against the bounds.
    pub fn index(&self, idx: &[i64]) -> Result<usize, VbaError> {
        if idx.len() != self.bounds.len() {
            return Err(err(9, "Subscript out of range"));
        }
        let mut at: usize = 0;
        for (i, (lo, hi)) in idx.iter().zip(&self.bounds) {
            if i < lo || i > hi {
                return Err(err(9, "Subscript out of range"));
            }
            let len = (hi - lo + 1) as usize;
            at = at * len + (i - lo) as usize;
        }
        Ok(at)
    }
}

/// A user-defined type's value (`Type … End Type`): field names and values, copied on assignment.
#[derive(Clone, Debug)]
pub struct Udt {
    pub type_name: String,
    pub fields: Vec<(String, Value)>,
}

/// An object reference.
#[derive(Clone, Debug)]
pub enum Obj {
    Host(HostRef),
    /// An instance of a class module.
    Instance(Rc<RefCell<crate::interp::Instance>>),
    Collection(Rc<RefCell<crate::builtins::Collection>>),
    Dictionary(Rc<RefCell<crate::builtins::Dictionary>>),
    /// The `Err` object.
    Err,
}

impl Obj {
    pub fn same(&self, other: &Obj) -> bool {
        match (self, other) {
            (Obj::Host(a), Obj::Host(b)) => a == b,
            (Obj::Instance(a), Obj::Instance(b)) => Rc::ptr_eq(a, b),
            (Obj::Collection(a), Obj::Collection(b)) => Rc::ptr_eq(a, b),
            (Obj::Dictionary(a), Obj::Dictionary(b)) => Rc::ptr_eq(a, b),
            (Obj::Err, Obj::Err) => true,
            _ => false,
        }
    }
    pub fn type_name(&self) -> String {
        match self {
            Obj::Host(h) => h.class.clone(),
            Obj::Instance(i) => i.borrow().class.clone(),
            Obj::Collection(_) => "Collection".into(),
            Obj::Dictionary(_) => "Dictionary".into(),
            Obj::Err => "ErrObject".into(),
        }
    }
}

#[derive(Clone, Debug, Default)]
pub enum Value {
    #[default]
    Empty,
    Null,
    Nothing,
    Bool(bool),
    Byte(u8),
    Integer(i16),
    Long(i32),
    LongLong(i64),
    Single(f32),
    Double(f64),
    /// Scaled by 10,000.
    Currency(i64),
    /// Days since 1899-12-30, time as the fraction.
    Date(f64),
    String(Rc<str>),
    /// `CVErr(n)`: an error value (cells holding #N/A come back as `Error 2042`).
    Error(i32),
    Object(Obj),
    Array(Rc<RefCell<Array>>),
    Udt(Box<Udt>),
}

impl Value {
    pub fn str(s: impl Into<Rc<str>>) -> Value {
        Value::String(s.into())
    }

    pub fn is_numeric_type(&self) -> bool {
        matches!(
            self,
            Value::Byte(_)
                | Value::Integer(_)
                | Value::Long(_)
                | Value::LongLong(_)
                | Value::Single(_)
                | Value::Double(_)
                | Value::Currency(_)
                | Value::Bool(_)
        )
    }

    pub fn is_object(&self) -> bool {
        matches!(self, Value::Object(_) | Value::Nothing)
    }

    /// `TypeName`.
    pub fn type_name(&self) -> String {
        match self {
            Value::Empty => "Empty".into(),
            Value::Null => "Null".into(),
            Value::Nothing => "Nothing".into(),
            Value::Bool(_) => "Boolean".into(),
            Value::Byte(_) => "Byte".into(),
            Value::Integer(_) => "Integer".into(),
            Value::Long(_) => "Long".into(),
            Value::LongLong(_) => "LongLong".into(),
            Value::Single(_) => "Single".into(),
            Value::Double(_) => "Double".into(),
            Value::Currency(_) => "Currency".into(),
            Value::Date(_) => "Date".into(),
            Value::String(_) => "String".into(),
            Value::Error(_) => "Error".into(),
            Value::Object(o) => o.type_name(),
            Value::Array(a) => {
                let base = a.borrow().items.first().map(|v| v.type_name()).unwrap_or_else(|| "Variant".into());
                format!("{base}()")
            }
            Value::Udt(u) => u.type_name.clone(),
        }
    }

    /// `VarType`.
    pub fn var_type(&self) -> i32 {
        match self {
            Value::Empty => 0,
            Value::Null => 1,
            Value::Integer(_) => 2,
            Value::Long(_) => 3,
            Value::Single(_) => 4,
            Value::Double(_) => 5,
            Value::Currency(_) => 6,
            Value::Date(_) => 7,
            Value::String(_) => 8,
            Value::Object(_) | Value::Nothing => 9,
            Value::Error(_) => 10,
            Value::Bool(_) => 11,
            Value::Udt(_) => 36,
            Value::Byte(_) => 17,
            Value::LongLong(_) => 20,
            Value::Array(_) => 8192 + 12,
        }
    }

    // ------------------------------------------------------------ conversions

    pub fn to_f64(&self) -> Result<f64, VbaError> {
        Ok(match self {
            Value::Empty => 0.0,
            Value::Bool(b) => {
                if *b {
                    -1.0
                } else {
                    0.0
                }
            }
            Value::Byte(n) => f64::from(*n),
            Value::Integer(n) => f64::from(*n),
            Value::Long(n) => f64::from(*n),
            Value::LongLong(n) => *n as f64,
            Value::Single(n) => f64::from(*n),
            Value::Double(n) | Value::Date(n) => *n,
            Value::Currency(n) => *n as f64 / 10_000.0,
            Value::String(s) => parse_number(s).ok_or_else(|| err(13, "Type mismatch"))?,
            Value::Null => return Err(err(94, "Invalid use of Null")),
            Value::Error(_) | Value::Object(_) | Value::Nothing | Value::Array(_) | Value::Udt(_) => return Err(err(13, "Type mismatch")),
        })
    }

    /// Rounded to an integer the way VBA does (banker's rounding), checked to fit `i64`.
    pub fn to_i64(&self) -> Result<i64, VbaError> {
        match self {
            Value::Byte(n) => Ok(i64::from(*n)),
            Value::Integer(n) => Ok(i64::from(*n)),
            Value::Long(n) => Ok(i64::from(*n)),
            Value::LongLong(n) => Ok(*n),
            Value::Bool(b) => Ok(if *b { -1 } else { 0 }),
            v => {
                let f = v.to_f64()?;
                let r = round_even(f);
                if !r.is_finite() || r.abs() >= 9.2e18 {
                    return Err(err(6, "Overflow"));
                }
                Ok(r as i64)
            }
        }
    }

    pub fn to_i32(&self) -> Result<i32, VbaError> {
        i32::try_from(self.to_i64()?).map_err(|_| err(6, "Overflow"))
    }

    pub fn to_bool(&self) -> Result<bool, VbaError> {
        match self {
            Value::Bool(b) => Ok(*b),
            Value::String(s) => match s.trim().to_ascii_lowercase().as_str() {
                "true" | "verdadero" => Ok(true),
                "false" | "falso" => Ok(false),
                _ => Ok(parse_number(s).ok_or_else(|| err(13, "Type mismatch"))? != 0.0),
            },
            v => Ok(v.to_f64()? != 0.0),
        }
    }

    /// `CStr`: the value as text.
    pub fn to_str(&self) -> Result<String, VbaError> {
        Ok(match self {
            Value::Empty => String::new(),
            Value::String(s) => s.to_string(),
            Value::Bool(b) => (if *b { "True" } else { "False" }).into(),
            Value::Byte(n) => n.to_string(),
            Value::Integer(n) => n.to_string(),
            Value::Long(n) => n.to_string(),
            Value::LongLong(n) => n.to_string(),
            Value::Single(n) => format_number(f64::from(*n), 7),
            Value::Double(n) => format_number(*n, 15),
            Value::Currency(n) => format_number(*n as f64 / 10_000.0, 19),
            Value::Date(d) => format_date(*d),
            Value::Error(n) => format!("Error {n}"),
            Value::Null => return Err(err(94, "Invalid use of Null")),
            Value::Object(_) | Value::Nothing | Value::Array(_) | Value::Udt(_) => return Err(err(13, "Type mismatch")),
        })
    }

    pub fn to_date(&self) -> Result<f64, VbaError> {
        match self {
            Value::Date(d) => Ok(*d),
            Value::String(s) => parse_date(s).ok_or_else(|| err(13, "Type mismatch")),
            v => v.to_f64(),
        }
    }

    /// A For-loop counter value: an integer subtype when it has no fraction, else Double.
    pub fn from_f64_natural(n: f64) -> Value {
        if n.fract() == 0.0 && n.abs() < 2.1e9 { Value::from_i64(n as i64) } else { Value::Double(n) }
    }

    /// The narrowest integer subtype holding `n`, else Double (Variant arithmetic promotes on
    /// overflow instead of failing).
    pub fn from_i64(n: i64) -> Value {
        if let Ok(x) = i16::try_from(n) {
            Value::Integer(x)
        } else if let Ok(x) = i32::try_from(n) {
            Value::Long(x)
        } else {
            Value::Double(n as f64)
        }
    }
}

impl fmt::Display for Value {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Value::Null => f.write_str("Null"),
            Value::Nothing => f.write_str("Nothing"),
            v => f.write_str(&v.to_str().unwrap_or_else(|_| v.type_name())),
        }
    }
}

/// VBA's rounding for conversions to integers: halves go to the even neighbour.
pub fn round_even(f: f64) -> f64 {
    let r = f.round();
    if (f - f.trunc()).abs() == 0.5 { 2.0 * (f / 2.0).round() } else { r }
}

/// A number written as text: decimal (with `.`), `&H` hex, `&O` octal, currency and percent
/// signs and spaces around it ignored. `None` when it isn't one.
pub fn parse_number(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    let lower = t.to_ascii_lowercase();
    if let Some(h) = lower.strip_prefix("&h") {
        return i64::from_str_radix(h, 16).ok().map(|n| n as f64);
    }
    if let Some(o) = lower.strip_prefix("&o") {
        return i64::from_str_radix(o, 8).ok().map(|n| n as f64);
    }
    let cleaned: String = t.chars().filter(|c| !matches!(c, '$' | '€' | ',' | ' ')).collect();
    let (neg, body) =
        if let Some(b) = cleaned.strip_prefix('(').and_then(|b| b.strip_suffix(')')) { (true, b.to_string()) } else { (false, cleaned) };
    let (body, pct) = match body.strip_suffix('%') {
        Some(b) => (b.to_string(), true),
        None => (body, false),
    };
    if !body.chars().any(|c| c.is_ascii_digit()) || body.chars().any(|c| c.is_alphabetic() && c != 'e' && c != 'E') {
        return None;
    }
    let n: f64 = body.parse().ok()?;
    let n = if pct { n / 100.0 } else { n };
    Some(if neg { -n } else { n })
}

/// A double the way `CStr` writes it: up to `digits` significant digits, no trailing zeros,
/// scientific notation outside 1E-5 … 1E15 (`1E+16`, `1E-06`).
pub fn format_number(n: f64, digits: usize) -> String {
    if n == 0.0 {
        return "0".into();
    }
    if !n.is_finite() {
        return if n.is_nan() {
            "NaN".into()
        } else if n > 0.0 {
            "Infinity".into()
        } else {
            "-Infinity".into()
        };
    }
    let exp = n.abs().log10().floor() as i32;
    if exp >= digits as i32 || exp < -5 {
        let s = format!("{:.*E}", digits.saturating_sub(1), n);
        let (mant, e) = s.split_once('E').unwrap_or((&s, "0"));
        let mant = if mant.contains('.') { mant.trim_end_matches('0').trim_end_matches('.') } else { mant };
        let e: i32 = e.parse().unwrap_or(0);
        return format!("{mant}E{}{:02}", if e < 0 { '-' } else { '+' }, e.abs());
    }
    let decimals = (digits as i32 - 1 - exp).max(0) as usize;
    let s = format!("{:.*}", decimals, n);
    if s.contains('.') { s.trim_end_matches('0').trim_end_matches('.').to_string() } else { s }
}

/// Days since 1899-12-30 of a date.
pub fn date_serial(y: i64, m: i64, d: i64) -> Option<f64> {
    // Months roll over years like DateSerial does.
    let total = y.checked_mul(12)?.checked_add(m - 1)?;
    let (y, m) = (total.div_euclid(12), total.rem_euclid(12) + 1);
    let y = i32::try_from(y).ok()?;
    let first = gridcraft_core::date::days_from_civil(y, m as u32, 1) + 25_569;
    Some((first + d - 1) as f64)
}

/// (year, month, day, hour, minute, second) of a date value.
pub fn date_parts(d: f64) -> (i32, u32, u32, u32, u32, u32) {
    let day = d.trunc() as i64;
    let frac = (d - d.trunc()).abs();
    let secs = (frac * 86_400.0).round() as i64;
    let (day, secs) = if secs >= 86_400 { (day + if d < 0.0 { -1 } else { 1 }, secs - 86_400) } else { (day, secs) };
    let (y, m, dd) = gridcraft_core::date::civil_from_days(day - 25_569);
    (y, m, dd, (secs / 3600) as u32, (secs / 60 % 60) as u32, (secs % 60) as u32)
}

/// A date the way `CStr` writes it (U.S. short date and long time; the time alone for a date
/// with no day part, the date alone with no time).
pub fn format_date(d: f64) -> String {
    let (y, m, dd, h, mi, s) = date_parts(d);
    let date = format!("{m}/{dd}/{y}");
    let (h12, ampm) = match h {
        0 => (12, "AM"),
        1..=11 => (h, "AM"),
        12 => (12, "PM"),
        _ => (h - 12, "PM"),
    };
    let time = format!("{h12}:{mi:02}:{s:02} {ampm}");
    let has_time = h != 0 || mi != 0 || s != 0;
    if d.trunc() == 0.0 && has_time {
        time
    } else if has_time {
        format!("{date} {time}")
    } else {
        date
    }
}

/// A date written as text: `m/d/yyyy`, `yyyy-mm-dd`, with an optional time (`h:mm[:ss] [AM|PM]`),
/// or a time alone.
pub fn parse_date(s: &str) -> Option<f64> {
    let t = s.trim().trim_matches('#');
    if t.is_empty() {
        return None;
    }
    let mut date = 0.0;
    let mut rest = t;
    let first = t.split_whitespace().next().unwrap_or("");
    let has_date = first.contains('/') || (first.matches('-').count() == 2);
    if has_date {
        let sep = if first.contains('/') { '/' } else { '-' };
        let p: Vec<i64> = first.split(sep).map(|x| x.parse::<i64>()).collect::<Result<_, _>>().ok()?;
        let (y, m, d) = match p.as_slice() {
            [a, b, c] if *a > 31 => (*a, *b, *c),
            [m, d, y] => (if *y < 100 { if *y < 30 { 2000 + y } else { 1900 + y } } else { *y }, *m, *d),
            [m, d] => (date_parts(crate::builtins::now()).0 as i64, *m, *d),
            _ => return None,
        };
        if !(1..=12).contains(&m) || !(1..=31).contains(&d) {
            return None;
        }
        date = date_serial(y, m, d)?;
        rest = t.get(first.len()..).unwrap_or("").trim();
    }
    if rest.is_empty() {
        return has_date.then_some(date);
    }
    let lower = rest.to_ascii_lowercase();
    let (clock, pm, am) = match (lower.strip_suffix("pm"), lower.strip_suffix("am")) {
        (Some(c), _) => (c.trim().to_string(), true, false),
        (_, Some(c)) => (c.trim().to_string(), false, true),
        _ => (lower.clone(), false, false),
    };
    let p: Vec<f64> = clock.split(':').map(|x| x.trim().parse::<f64>()).collect::<Result<_, _>>().ok()?;
    if p.len() < 2 || p.len() > 3 {
        return None;
    }
    let mut h = p[0];
    if pm && h < 12.0 {
        h += 12.0;
    }
    if am && h == 12.0 {
        h = 0.0;
    }
    let secs = h * 3600.0 + p[1] * 60.0 + p.get(2).copied().unwrap_or(0.0);
    Some(date + secs / 86_400.0)
}

// ---------------------------------------------------------------- operators

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BinOp {
    Add,
    Sub,
    Mul,
    Div,
    IntDiv,
    Mod,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Gt,
    Le,
    Ge,
    Like,
    Is,
    And,
    Or,
    Xor,
    Eqv,
    Imp,
}

/// Which numeric type an arithmetic result takes.
fn rank(v: &Value) -> u8 {
    match v {
        Value::Byte(_) => 1,
        Value::Integer(_) | Value::Bool(_) | Value::Empty => 2,
        Value::Long(_) => 3,
        Value::LongLong(_) => 4,
        Value::Single(_) => 5,
        Value::Currency(_) => 6,
        Value::Double(_) | Value::String(_) => 7,
        Value::Date(_) => 8,
        _ => 7,
    }
}

fn integer_result(n: i64, a: &Value, b: &Value) -> Value {
    match rank(a).max(rank(b)) {
        1 if (0..=255).contains(&n) => Value::Byte(n as u8),
        r if r <= 2 => Value::from_i64(n),
        3 => i32::try_from(n).map(Value::Long).unwrap_or(Value::Double(n as f64)),
        _ => Value::LongLong(n),
    }
}

pub fn binary(op: BinOp, a: &Value, b: &Value, text_compare: bool) -> Result<Value, VbaError> {
    use BinOp::*;
    // Null propagates through arithmetic and comparisons.
    let null = matches!(a, Value::Null) || matches!(b, Value::Null);
    match op {
        Concat => {
            if matches!(a, Value::Null) && matches!(b, Value::Null) {
                return Ok(Value::Null);
            }
            let s = |v: &Value| if matches!(v, Value::Null) { Ok(String::new()) } else { v.to_str() };
            return Ok(Value::str(format!("{}{}", s(a)?, s(b)?)));
        }
        Is => {
            return match (a, b) {
                (Value::Nothing, Value::Nothing) => Ok(Value::Bool(true)),
                (Value::Object(_), Value::Nothing) | (Value::Nothing, Value::Object(_)) => Ok(Value::Bool(false)),
                (Value::Object(x), Value::Object(y)) => Ok(Value::Bool(x.same(y))),
                _ => Err(err(424, "Object required")),
            };
        }
        Eq | Ne | Lt | Gt | Le | Ge => {
            if null {
                return Ok(Value::Null);
            }
            let ord = compare(a, b, text_compare)?;
            use std::cmp::Ordering::*;
            let r = match op {
                Eq => ord == Equal,
                Ne => ord != Equal,
                Lt => ord == Less,
                Gt => ord == Greater,
                Le => ord != Greater,
                _ => ord != Less,
            };
            return Ok(Value::Bool(r));
        }
        Like => {
            if null {
                return Ok(Value::Null);
            }
            return Ok(Value::Bool(like(&a.to_str()?, &b.to_str()?, text_compare)));
        }
        And | Or | Xor | Eqv | Imp => {
            if let (Value::Bool(x), Value::Bool(y)) = (a, b) {
                return Ok(Value::Bool(match op {
                    And => *x && *y,
                    Or => *x || *y,
                    Xor => x != y,
                    Eqv => x == y,
                    _ => !*x || *y,
                }));
            }
            if null {
                // Null And False is False, Null Or True is True; otherwise Null.
                let other = if matches!(a, Value::Null) { b } else { a };
                if !matches!(other, Value::Null) {
                    let o = other.to_i64()?;
                    if op == And && o == 0 {
                        return Ok(Value::Bool(false));
                    }
                    if op == Or && o == -1 {
                        return Ok(Value::Bool(true));
                    }
                }
                return Ok(Value::Null);
            }
            let bools = matches!(a, Value::Bool(_)) || matches!(b, Value::Bool(_));
            let (x, y) = (a.to_i64()?, b.to_i64()?);
            let r = match op {
                And => x & y,
                Or => x | y,
                Xor => x ^ y,
                Eqv => !(x ^ y),
                _ => !x | y,
            };
            if bools && (r == 0 || r == -1) && matches!((a, b), (Value::Bool(_), Value::Bool(_))) {
                return Ok(Value::Bool(r != 0));
            }
            return Ok(integer_result(r, a, b));
        }
        _ => {}
    }
    if null {
        return Ok(Value::Null);
    }
    // `+` on two strings concatenates.
    if op == Add && matches!(a, Value::String(_)) && matches!(b, Value::String(_)) {
        return Ok(Value::str(format!("{}{}", a.to_str()?, b.to_str()?)));
    }
    let date = matches!(a, Value::Date(_)) || matches!(b, Value::Date(_));
    match op {
        Add | Sub | Mul => {
            let ints = !matches!(rank(a).max(rank(b)), 5..);
            if ints {
                let (x, y) = (a.to_i64()?, b.to_i64()?);
                let r = match op {
                    Add => x.checked_add(y),
                    Sub => x.checked_sub(y),
                    _ => x.checked_mul(y),
                };
                return Ok(match r {
                    Some(n) => integer_result(n, a, b),
                    None => Value::Double(match op {
                        Add => x as f64 + y as f64,
                        Sub => x as f64 - y as f64,
                        _ => x as f64 * y as f64,
                    }),
                });
            }
            let (x, y) = (a.to_f64()?, b.to_f64()?);
            let r = match op {
                Add => x + y,
                Sub => x - y,
                _ => x * y,
            };
            if date && op != Mul && !(matches!(a, Value::Date(_)) && matches!(b, Value::Date(_)) && op == Sub) {
                return Ok(Value::Date(r));
            }
            if matches!(a, Value::Currency(_)) || matches!(b, Value::Currency(_)) {
                return Ok(Value::Currency((r * 10_000.0).round() as i64));
            }
            if rank(a).max(rank(b)) == 5 && !matches!(a, Value::Double(_) | Value::String(_)) && !matches!(b, Value::Double(_) | Value::String(_)) {
                return Ok(Value::Single(r as f32));
            }
            Ok(Value::Double(r))
        }
        Div => {
            let (x, y) = (a.to_f64()?, b.to_f64()?);
            if y == 0.0 {
                return Err(if x == 0.0 { err(6, "Overflow") } else { err(11, "Division by zero") });
            }
            Ok(Value::Double(x / y))
        }
        IntDiv | Mod => {
            let (x, y) = (a.to_i64()?, b.to_i64()?);
            if y == 0 {
                return Err(err(11, "Division by zero"));
            }
            let r = if op == IntDiv { x / y } else { x % y };
            Ok(integer_result(r, a, b))
        }
        Pow => {
            let (x, y) = (a.to_f64()?, b.to_f64()?);
            let r = x.powf(y);
            if r.is_nan() {
                return Err(err(5, "Invalid procedure call or argument"));
            }
            Ok(Value::Double(r))
        }
        _ => Err(err(13, "Type mismatch")),
    }
}

pub fn negate(v: &Value) -> Result<Value, VbaError> {
    Ok(match v {
        Value::Null => Value::Null,
        Value::Empty => Value::Integer(0),
        Value::Integer(n) => n.checked_neg().map(Value::Integer).unwrap_or(Value::Long(-i32::from(*n))),
        Value::Long(n) => n.checked_neg().map(Value::Long).unwrap_or(Value::Double(-f64::from(*n))),
        Value::LongLong(n) => n.checked_neg().map(Value::LongLong).ok_or_else(|| err(6, "Overflow"))?,
        Value::Byte(n) => Value::Integer(-i16::from(*n)),
        Value::Bool(b) => Value::Integer(if *b { 1 } else { 0 }),
        Value::Single(n) => Value::Single(-n),
        Value::Currency(n) => Value::Currency(-n),
        Value::Date(d) => Value::Date(-d),
        other => Value::Double(-other.to_f64()?),
    })
}

pub fn not(v: &Value) -> Result<Value, VbaError> {
    Ok(match v {
        Value::Null => Value::Null,
        Value::Bool(b) => Value::Bool(!b),
        other => Value::from_i64(!other.to_i64()?),
    })
}

/// Compares two values: numbers numerically, text as text (binary or case-insensitive), a
/// number against text numerically when the text is a number, else as text.
pub fn compare(a: &Value, b: &Value, text: bool) -> Result<std::cmp::Ordering, VbaError> {
    let is_str = |v: &Value| matches!(v, Value::String(_));
    if is_str(a) && is_str(b) || (is_str(a) && matches!(b, Value::Empty)) || (matches!(a, Value::Empty) && is_str(b)) {
        let (x, y) = (a.to_str()?, b.to_str()?);
        return Ok(if text { x.to_lowercase().cmp(&y.to_lowercase()) } else { x.cmp(&y) });
    }
    if is_str(a) || is_str(b) {
        let (s, n, flip) = if is_str(a) { (a, b, false) } else { (b, a, true) };
        if let Some(x) = parse_number(&s.to_str()?) {
            let y = n.to_f64()?;
            let o = x.partial_cmp(&y).unwrap_or(std::cmp::Ordering::Equal);
            return Ok(if flip { o.reverse() } else { o });
        }
        let (x, y) = (a.to_str()?, b.to_str()?);
        return Ok(if text { x.to_lowercase().cmp(&y.to_lowercase()) } else { x.cmp(&y) });
    }
    if matches!(a, Value::Error(_)) || matches!(b, Value::Error(_)) {
        return Err(err(13, "Type mismatch"));
    }
    let (x, y) = (a.to_f64()?, b.to_f64()?);
    Ok(x.partial_cmp(&y).unwrap_or(std::cmp::Ordering::Equal))
}

/// The `Like` operator: `?` one character, `*` any run, `#` one digit, `[a-z]`, `[!a-z]`.
pub fn like(s: &str, pattern: &str, text: bool) -> bool {
    let (s, p): (Vec<char>, Vec<char>) = if text {
        (s.to_lowercase().chars().collect(), pattern.to_lowercase().chars().collect())
    } else {
        (s.chars().collect(), pattern.chars().collect())
    };
    fn m(s: &[char], p: &[char], depth: usize) -> bool {
        if depth > 10_000 {
            return false;
        }
        let Some((&c, rest)) = p.split_first() else { return s.is_empty() };
        match c {
            '*' => (0..=s.len()).any(|i| m(&s[i..], rest, depth + 1)),
            '?' => !s.is_empty() && m(&s[1..], rest, depth + 1),
            '#' => s.first().is_some_and(|x| x.is_ascii_digit()) && m(&s[1..], rest, depth + 1),
            '[' => {
                let Some(close) = rest.iter().position(|&x| x == ']') else { return s.first() == Some(&'[') && m(&s[1..], rest, depth + 1) };
                let set = &rest[..close];
                let (neg, set) = match set.split_first() {
                    Some(('!', r)) => (true, r),
                    _ => (false, set),
                };
                let Some(&ch) = s.first() else { return false };
                let mut hit = false;
                let mut i = 0;
                while i < set.len() {
                    if i + 2 < set.len() && set[i + 1] == '-' {
                        hit |= set[i] <= ch && ch <= set[i + 2];
                        i += 3;
                    } else {
                        hit |= set[i] == ch;
                        i += 1;
                    }
                }
                hit != neg && m(&s[1..], &rest[close + 1..], depth + 1)
            }
            c => s.first() == Some(&c) && m(&s[1..], rest, depth + 1),
        }
    }
    m(&s, &p, 0)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn v(n: i16) -> Value {
        Value::Integer(n)
    }

    #[test]
    fn arithmetic_types() {
        assert!(matches!(binary(BinOp::Add, &v(1), &v(2), false).unwrap(), Value::Integer(3)));
        assert!(matches!(binary(BinOp::Mul, &v(300), &v(300), false).unwrap(), Value::Long(90_000)));
        assert!(matches!(binary(BinOp::Div, &v(1), &v(2), false).unwrap(), Value::Double(x) if x == 0.5));
        assert!(matches!(binary(BinOp::Add, &Value::str("1"), &v(2), false).unwrap(), Value::Double(x) if x == 3.0));
        assert_eq!(binary(BinOp::Add, &Value::str("a"), &Value::str("b"), false).unwrap().to_str().unwrap(), "ab");
        assert_eq!(binary(BinOp::Concat, &v(1), &Value::Null, false).unwrap().to_str().unwrap(), "1");
        assert!(matches!(binary(BinOp::IntDiv, &Value::Double(7.5), &v(2), false).unwrap(), Value::LongLong(4)));
        assert!(matches!(binary(BinOp::Mod, &v(-7), &v(3), false).unwrap(), Value::Integer(-1)));
        assert_eq!(binary(BinOp::Div, &v(1), &v(0), false).unwrap_err().number, 11);
        assert!(matches!(binary(BinOp::Add, &Value::Date(45000.0), &v(1), false).unwrap(), Value::Date(x) if x == 45001.0));
    }

    #[test]
    fn comparisons_and_logic() {
        assert!(matches!(binary(BinOp::Eq, &Value::str("A"), &Value::str("a"), true).unwrap(), Value::Bool(true)));
        assert!(matches!(binary(BinOp::Eq, &Value::str("A"), &Value::str("a"), false).unwrap(), Value::Bool(false)));
        assert!(matches!(binary(BinOp::Eq, &Value::str("5"), &v(5), false).unwrap(), Value::Bool(true)));
        assert!(matches!(binary(BinOp::Lt, &Value::Empty, &v(1), false).unwrap(), Value::Bool(true)));
        assert!(matches!(binary(BinOp::And, &v(12), &v(10), false).unwrap(), Value::Integer(8)));
        assert!(matches!(not(&Value::Integer(0)).unwrap(), Value::Integer(-1)));
        assert!(matches!(binary(BinOp::Eq, &Value::Null, &v(1), false).unwrap(), Value::Null));
    }

    #[test]
    fn like_patterns() {
        assert!(like("Factura 2023", "Factura ####", false));
        assert!(like("abc", "a*", false));
        assert!(like("b", "[a-c]", false));
        assert!(!like("d", "[a-c]", false));
        assert!(like("d", "[!a-c]", false));
        assert!(like("ABC", "a?c", true));
        assert!(!like("ABC", "a?c", false));
    }

    #[test]
    fn text_forms() {
        assert_eq!(format_number(0.1 + 0.2, 15), "0.3");
        assert_eq!(format_number(1e16, 15), "1E+16");
        assert_eq!(format_number(0.000001, 15), "1E-06");
        assert_eq!(format_number(-2.5, 15), "-2.5");
        assert_eq!(Value::Bool(true).to_str().unwrap(), "True");
        assert_eq!(parse_number(" 1,234.5 "), Some(1234.5));
        assert_eq!(parse_number("&HFF"), Some(255.0));
        assert_eq!(parse_number("abc"), None);
        assert_eq!(round_even(2.5), 2.0);
        assert_eq!(round_even(3.5), 4.0);
        assert_eq!(Value::Double(2.5).to_i64().unwrap(), 2);
    }

    #[test]
    fn dates() {
        let d = date_serial(2023, 3, 15).unwrap();
        assert_eq!(d, 45000.0);
        assert_eq!(format_date(d + 0.5), "3/15/2023 12:00:00 PM");
        assert_eq!(format_date(0.75), "6:00:00 PM");
        assert_eq!(parse_date("2023-03-15"), Some(45000.0));
        assert_eq!(parse_date("3/15/2023 6:00 PM"), Some(45000.75));
        assert_eq!(date_serial(2023, 13, 1), date_serial(2024, 1, 1));
        assert_eq!(date_parts(45000.5), (2023, 3, 15, 12, 0, 0));
    }
}

//! Tokens of VBA source ([MS-VBAL] 3.3): identifiers (with type suffixes), literals, operators,
//! line ends. Comments, `Attribute` lines and line continuations (` _`) are removed here.

use crate::error::{VbaError, syntax};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum NumKind {
    Integer,
    Long,
    LongLong,
    Single,
    Double,
    Currency,
}

#[derive(Clone, Debug, PartialEq)]
pub enum Tok {
    /// A name: `suffix` is its type character (`$`, `%`, `&`, `!`, `#`, `@`, `^`).
    Ident {
        name: String,
        suffix: Option<char>,
        bracketed: bool,
    },
    Int(i64, NumKind),
    Float(f64, NumKind),
    Str(String),
    Date(f64),
    Op(&'static str),
    /// `:` between statements on a line.
    Colon,
    Newline,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Token {
    pub tok: Tok,
    /// 1-based source line.
    pub line: usize,
    /// Whitespace (or a line start) comes right before it: `With` blocks tell `.Value` from
    /// `x.Value` by it.
    pub spaced: bool,
}

const OPS: [&str; 20] = [":=", "<=", ">=", "<>", "=<", "=>", "+", "-", "*", "/", "\\", "^", "&", "=", "<", ">", ".", "!", "(", ")"];

/// Joins continued lines (` _`), remembering each logical line's first physical line.
fn logical_lines(src: &str) -> Vec<(usize, String)> {
    let mut logical: Vec<(usize, String)> = vec![];
    let mut pending: Option<(usize, String)> = None;
    for (i, raw) in src.lines().enumerate() {
        let line = raw.trim_end_matches('\r');
        let continued = line.trim_end().ends_with(" _") || line.trim_end() == "_";
        let body = if continued { line.trim_end().trim_end_matches('_').to_string() } else { line.to_string() };
        match pending.as_mut() {
            Some((_, acc)) => {
                acc.push(' ');
                acc.push_str(&body);
            }
            None => pending = Some((i + 1, body)),
        }
        if !continued && let Some(p) = pending.take() {
            logical.push(p);
        }
    }
    if let Some(p) = pending.take() {
        logical.push(p);
    }
    logical
}

/// Reads a string literal whose opening quote is at `start`: its text (`""` is a quote) and the
/// position after the closing quote.
fn string_literal(chars: &[char], start: usize, module: &str, line_no: usize) -> Result<(String, usize), VbaError> {
    let mut s = String::new();
    let mut i = start + 1;
    loop {
        match chars.get(i) {
            Some('"') if chars.get(i + 1) == Some(&'"') => {
                s.push('"');
                i += 2;
            }
            Some('"') => return Ok((s, i + 1)),
            Some(&x) => {
                s.push(x);
                i += 1;
            }
            None => return Err(syntax(module, line_no, "Expected: \"")),
        }
    }
}

/// Reads a `&H…`/`&O…` literal starting at `i` (an optional `&` suffix makes it a Long).
fn radix_literal(chars: &[char], i: usize, module: &str, line_no: usize) -> Result<(Tok, usize), VbaError> {
    let radix = if matches!(chars[i + 1], 'h' | 'H') { 16 } else { 8 };
    let mut j = i + 2;
    while chars.get(j).is_some_and(|x| x.is_digit(radix)) {
        j += 1;
    }
    let digits: String = chars[i + 2..j].iter().collect();
    let n = u64::from_str_radix(&digits, radix).map_err(|_| syntax(module, line_no, "Overflow"))?;
    let (n, kind) = match chars.get(j) {
        Some('&') => {
            j += 1;
            ((n as u32) as i32 as i64, NumKind::Long)
        }
        _ if n <= 0xFFFF => ((n as u16) as i16 as i64, NumKind::Integer),
        _ => ((n as u32) as i32 as i64, NumKind::Long),
    };
    Ok((Tok::Int(n, kind), j))
}

/// Reads a decimal number starting at `i`, with its exponent and type suffix.
fn number_literal(chars: &[char], i: usize, module: &str, line_no: usize) -> Result<(Tok, usize), VbaError> {
    let mut j = i;
    while chars.get(j).is_some_and(|x| x.is_ascii_digit()) {
        j += 1;
    }
    let mut float = false;
    if chars.get(j) == Some(&'.') && chars.get(j + 1).is_some_and(|x| x.is_ascii_digit()) {
        float = true;
        j += 1;
        while chars.get(j).is_some_and(|x| x.is_ascii_digit()) {
            j += 1;
        }
    } else if chars.get(j) == Some(&'.') && !chars.get(j + 1).is_some_and(|x| x.is_alphabetic()) {
        float = true;
        j += 1;
    }
    if matches!(chars.get(j), Some('e' | 'E' | 'd' | 'D'))
        && (chars.get(j + 1).is_some_and(|x| x.is_ascii_digit())
            || (matches!(chars.get(j + 1), Some('+' | '-')) && chars.get(j + 2).is_some_and(|x| x.is_ascii_digit())))
    {
        float = true;
        j += 2;
        while chars.get(j).is_some_and(|x| x.is_ascii_digit()) {
            j += 1;
        }
    }
    let text: String = chars[i..j].iter().map(|&x| if x == 'd' || x == 'D' { 'e' } else { x }).collect();
    let suffix = match chars.get(j) {
        Some(&s @ ('%' | '&' | '!' | '#' | '@' | '^')) => {
            j += 1;
            Some(s)
        }
        _ => None,
    };
    let tok = match (float, suffix) {
        (_, Some('!')) => Tok::Float(text.parse().unwrap_or(0.0), NumKind::Single),
        (_, Some('#')) => Tok::Float(text.parse().unwrap_or(0.0), NumKind::Double),
        (_, Some('@')) => Tok::Float(text.parse().unwrap_or(0.0), NumKind::Currency),
        (true, _) => Tok::Float(text.parse().map_err(|_| syntax(module, line_no, "Invalid number"))?, NumKind::Double),
        (false, s) => {
            let n: f64 = text.parse().map_err(|_| syntax(module, line_no, "Invalid number"))?;
            match s {
                Some('%') => Tok::Int(n as i64, NumKind::Integer),
                Some('&') => Tok::Int(n as i64, NumKind::Long),
                Some('^') => Tok::Int(n as i64, NumKind::LongLong),
                _ if n <= 32_767.0 => Tok::Int(n as i64, NumKind::Integer),
                _ if n <= 2_147_483_647.0 => Tok::Int(n as i64, NumKind::Long),
                _ => Tok::Float(n, NumKind::Double),
            }
        }
    };
    Ok((tok, j))
}

pub fn tokenize(module: &str, src: &str) -> Result<Vec<Token>, VbaError> {
    let mut out: Vec<Token> = vec![];
    for (line_no, text) in logical_lines(src) {
        let trimmed = text.trim_start();
        if trimmed.len() >= 10 && trimmed.get(..10).is_some_and(|t| t.eq_ignore_ascii_case("attribute ")) {
            continue;
        }
        let chars: Vec<char> = text.chars().collect();
        let mut i = 0usize;
        let mut spaced = true;
        let mut statement_start = true;
        while i < chars.len() {
            let c = chars[i];
            if c == ' ' || c == '\t' {
                i += 1;
                spaced = true;
                continue;
            }
            let push = |out: &mut Vec<Token>, tok: Tok, spaced: bool| out.push(Token { tok, line: line_no, spaced });
            if c == '\'' {
                break;
            }
            // Rem comments at the start of a statement.
            if statement_start && (c == 'R' || c == 'r') {
                let w: String = chars[i..chars.len().min(i + 4)].iter().collect();
                if w.len() >= 3 && w.get(..3).is_some_and(|x| x.eq_ignore_ascii_case("rem")) && w.chars().nth(3).is_none_or(|x| x == ' ' || x == '\t')
                {
                    break;
                }
            }
            statement_start = false;
            if c == '"' {
                let (s, next) = string_literal(&chars, i, module, line_no)?;
                i = next;
                push(&mut out, Tok::Str(s), spaced);
                spaced = false;
                continue;
            }
            if c == '[' {
                let Some(end) = chars[i..].iter().position(|&x| x == ']') else { return Err(syntax(module, line_no, "Expected: ]")) };
                let name: String = chars[i + 1..i + end].iter().collect();
                push(&mut out, Tok::Ident { name, suffix: None, bracketed: true }, spaced);
                i += end + 1;
                spaced = false;
                continue;
            }
            if c == '#' {
                // A date literal: #...# on one line that reads as a date.
                if let Some(end) = chars[i + 1..].iter().position(|&x| x == '#') {
                    let inner: String = chars[i + 1..i + 1 + end].iter().collect();
                    if let Some(d) = crate::value::parse_date(&inner) {
                        push(&mut out, Tok::Date(d), spaced);
                        i += end + 2;
                        spaced = false;
                        continue;
                    }
                }
                push(&mut out, Tok::Op("#"), spaced);
                i += 1;
                spaced = false;
                continue;
            }
            if c == '&' && matches!(chars.get(i + 1), Some('h' | 'H' | 'o' | 'O')) && chars.get(i + 2).is_some_and(|x| x.is_ascii_hexdigit()) {
                let (tok, next) = radix_literal(&chars, i, module, line_no)?;
                push(&mut out, tok, spaced);
                i = next;
                spaced = false;
                continue;
            }
            if c.is_ascii_digit() || (c == '.' && chars.get(i + 1).is_some_and(|x| x.is_ascii_digit()) && spaced) {
                let (tok, next) = number_literal(&chars, i, module, line_no)?;
                push(&mut out, tok, spaced);
                i = next;
                spaced = false;
                continue;
            }
            if c.is_alphabetic() || c == '_' {
                let mut j = i;
                while chars.get(j).is_some_and(|x| x.is_alphanumeric() || *x == '_') {
                    j += 1;
                }
                let name: String = chars[i..j].iter().collect();
                let suffix = match chars.get(j) {
                    // `x!Field` is a bang member access, `x!` a Single.
                    Some('!') if chars.get(j + 1).is_some_and(|x| x.is_alphabetic() || *x == '[') => None,
                    Some(&s @ ('%' | '&' | '!' | '#' | '@' | '$' | '^'))
                        if !(s == '&' && chars.get(j + 1).is_some_and(|x| x.is_alphanumeric()))
                            && !(s == '#' && chars.get(j + 1).is_some_and(|x| x.is_ascii_digit())) =>
                    {
                        j += 1;
                        Some(s)
                    }
                    _ => None,
                };
                push(&mut out, Tok::Ident { name, suffix, bracketed: false }, spaced);
                i = j;
                spaced = false;
                continue;
            }
            if c == ':' && chars.get(i + 1) != Some(&'=') {
                push(&mut out, Tok::Colon, spaced);
                i += 1;
                spaced = true;
                statement_start = true;
                continue;
            }
            if c == ',' || c == ';' {
                push(&mut out, Tok::Op(if c == ',' { "," } else { ";" }), spaced);
                i += 1;
                spaced = false;
                continue;
            }
            let rest: String = chars[i..chars.len().min(i + 2)].iter().collect();
            let Some(matched) = OPS.iter().find(|op| rest.starts_with(**op)) else {
                return Err(syntax(module, line_no, format!("Unexpected character `{c}`")));
            };
            let op: &'static str = match *matched {
                "=<" => "<=",
                "=>" => ">=",
                o => o,
            };
            push(&mut out, Tok::Op(op), spaced);
            i += matched.len();
            spaced = false;
        }
        // A line that produced no tokens (blank, comment-only, Attribute) adds no Newline.
        if out.last().map(|t| &t.tok) != Some(&Tok::Newline) && !out.is_empty() {
            out.push(Token { tok: Tok::Newline, line: line_no, spaced: true });
        }
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn toks(s: &str) -> Vec<Tok> {
        tokenize("M", s).unwrap().into_iter().map(|t| t.tok).collect()
    }

    fn id(s: &str) -> Tok {
        Tok::Ident { name: s.into(), suffix: None, bracketed: false }
    }

    #[test]
    fn basics() {
        assert_eq!(
            toks("x = \"a\"\"b\" & 3.5 ' note"),
            vec![id("x"), Tok::Op("="), Tok::Str("a\"b".into()), Tok::Op("&"), Tok::Float(3.5, NumKind::Double), Tok::Newline]
        );
        assert_eq!(toks("a = &HFF&"), vec![id("a"), Tok::Op("="), Tok::Int(255, NumKind::Long), Tok::Newline]);
        assert_eq!(toks("s = Left$(t, 2)")[2], Tok::Ident { name: "Left".into(), suffix: Some('$'), bracketed: false });
        assert_eq!(toks("Rem whole line\nx")[0], id("x"));
        assert_eq!(toks("Attribute VB_Name = \"M\"\nx").len(), 2);
        assert_eq!(toks("d = #3/15/2023#")[2], Tok::Date(45000.0));
        assert_eq!(toks("a <> b")[1], Tok::Op("<>"));
        assert_eq!(toks("x = 40000")[2], Tok::Int(40000, NumKind::Long));
        assert!(matches!(toks("Print #1, x")[1], Tok::Op("#")));
    }

    #[test]
    fn continuations_and_lines() {
        let t = tokenize("M", "a = 1 + _\r\n    2\r\nb = 3").unwrap();
        assert_eq!(t.iter().filter(|x| x.tok == Tok::Newline).count(), 2);
        assert_eq!(t.last().map(|x| x.line), Some(3));
        assert!(tokenize("M", "x = \"open").is_err());
        assert_eq!(toks("rs!Name")[1], Tok::Op("!"));
        assert_eq!(toks("a: b")[1], Tok::Colon);
        assert_eq!(toks("f x:=1")[2], Tok::Op(":="));
    }
}

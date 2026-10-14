//! Power Query definitions (`customXml` DataMashup) that GridCraft can run itself.
//!
//! Excel's Get Data saves each query's M formula in a base64 package inside a custom XML part,
//! and a `Microsoft.Mashup.OleDb.1` connection that names the query (`Location=Query1`). Queries
//! that only read SQL Server, `Sql.Database("server", "db", [Query="…"])` or a table picked in the
//! Navigator (`Source{[Schema="dbo",Item="T"]}[Data]`), become GridCraft queries; any other M
//! (filters, merges, other sources) is left to the pass-through, untouched.
//!
//! Layout of the package ([MS-QDEFF] §2): a 4-byte version, a 4-byte length, then a ZIP holding
//! `Formulas/Section1.m`.

use gridcraft_model::QueryCommand;

/// A Power Query that reads SQL Server and nothing else.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SqlQuery {
    /// The query's name (`shared Name = …`), which the connection's `Location` names.
    pub name: String,
    pub server: String,
    pub database: String,
    pub command: QueryCommand,
}

/// Largest DataMashup part decoded, and largest M formula read.
const MAX_PART: usize = 32 << 20;
const MAX_FORMULA: usize = 4 << 20;

/// The SQL Server queries of a DataMashup custom XML part (UTF-16 or UTF-8). Anything unreadable
/// gives no queries.
pub fn sql_queries(part: &[u8]) -> Vec<SqlQuery> {
    if part.len() > MAX_PART {
        return vec![];
    }
    let text = decode_text(part);
    let Some(at) = text.find("DataMashup") else { return vec![] };
    let rest = text.get(at..).unwrap_or("");
    let (Some(open), Some(close)) = (rest.find('>'), rest.find("</")) else { return vec![] };
    let Some(b64) = rest.get(open + 1..close) else { return vec![] };
    let Ok(bytes) = crate::crypto::base64_decode(b64) else { return vec![] };
    let len = |i: usize| bytes.get(i..i + 4).and_then(|b| b.try_into().ok()).map(u32::from_le_bytes);
    let Some(n) = len(4).map(|n| n as usize) else { return vec![] };
    let Some(pkg) = bytes.get(8..8usize.saturating_add(n)) else { return vec![] };
    let Ok(mut zip) = crate::package::Package::open(pkg) else { return vec![] };
    let Ok(Some(m)) = zip.read_limited("Formulas/Section1.m", MAX_FORMULA as u64) else { return vec![] };
    let m = String::from_utf8_lossy(&m);
    parse_section(m.trim_start_matches('\u{feff}'))
}

fn decode_text(b: &[u8]) -> String {
    let utf16 = |le: bool, b: &[u8]| {
        let units: Vec<u16> = b.as_chunks::<2>().0.iter().map(|c| if le { u16::from_le_bytes(*c) } else { u16::from_be_bytes(*c) }).collect();
        String::from_utf16_lossy(&units)
    };
    match b {
        [0xFF, 0xFE, rest @ ..] => utf16(true, rest),
        [0xFE, 0xFF, rest @ ..] => utf16(false, rest),
        [b'<', 0, ..] => utf16(true, b),
        _ => String::from_utf8_lossy(b).into_owned(),
    }
}

// ---------------------------------------------------------------- M

#[derive(Clone, Debug, PartialEq)]
enum Tok {
    Str(String),
    Ident(String),
    Punct(char),
    Other,
}

/// `#(lf)`, `#(cr)`, `#(tab)`, `#(#)` and `#(000D)`-style escapes in a string literal.
fn unescape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut rest = s;
    while let Some(i) = rest.find("#(") {
        out.push_str(rest.get(..i).unwrap_or(""));
        let after = rest.get(i + 2..).unwrap_or("");
        let Some(end) = after.find(')') else {
            out.push_str(rest.get(i..).unwrap_or(""));
            return out;
        };
        let body = after.get(..end).unwrap_or("");
        let mut ok = true;
        for code in body.split(',') {
            match code.trim() {
                "lf" => out.push('\n'),
                "cr" => out.push('\r'),
                "tab" => out.push('\t'),
                "#" => out.push('#'),
                h if (h.len() == 4 || h.len() == 8) && h.chars().all(|c| c.is_ascii_hexdigit()) => {
                    match u32::from_str_radix(h, 16).ok().and_then(char::from_u32) {
                        Some(c) => out.push(c),
                        None => ok = false,
                    }
                }
                _ => ok = false,
            }
        }
        if !ok {
            out.push_str("#(");
            out.push_str(body);
            out.push(')');
        }
        rest = after.get(end + 1..).unwrap_or("");
    }
    out.push_str(rest);
    out
}

fn tokenize(m: &str) -> Vec<Tok> {
    let mut out = vec![];
    let mut chars = m.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            c if c.is_whitespace() => {}
            '/' if chars.peek() == Some(&'/') => {
                for x in chars.by_ref() {
                    if x == '\n' {
                        break;
                    }
                }
            }
            '/' if chars.peek() == Some(&'*') => {
                chars.next();
                let mut prev = ' ';
                for x in chars.by_ref() {
                    if prev == '*' && x == '/' {
                        break;
                    }
                    prev = x;
                }
            }
            '"' => {
                let mut s = String::new();
                loop {
                    match chars.next() {
                        Some('"') if chars.peek() == Some(&'"') => {
                            chars.next();
                            s.push('"');
                        }
                        Some('"') | None => break,
                        Some(x) => s.push(x),
                    }
                }
                out.push(Tok::Str(unescape(&s)));
            }
            '#' if chars.peek() == Some(&'"') => {
                chars.next();
                let mut s = String::new();
                loop {
                    match chars.next() {
                        Some('"') if chars.peek() == Some(&'"') => {
                            chars.next();
                            s.push('"');
                        }
                        Some('"') | None => break,
                        Some(x) => s.push(x),
                    }
                }
                out.push(Tok::Ident(s));
            }
            c if c.is_alphabetic() || c == '_' => {
                let mut s = String::from(c);
                while let Some(&x) = chars.peek() {
                    if x.is_alphanumeric() || x == '_' || x == '.' {
                        s.push(x);
                        chars.next();
                    } else {
                        break;
                    }
                }
                out.push(Tok::Ident(s));
            }
            '=' | ',' | ';' | '(' | ')' | '[' | ']' | '{' | '}' => out.push(Tok::Punct(c)),
            _ => out.push(Tok::Other),
        }
    }
    out
}

/// The SQL Server queries among a section document's `shared` members.
pub fn parse_section(m: &str) -> Vec<SqlQuery> {
    let toks = tokenize(m);
    let mut out = vec![];
    let mut i = 0;
    while i < toks.len() {
        if toks.get(i) == Some(&Tok::Ident("shared".into()))
            && let (Some(Tok::Ident(name)), Some(Tok::Punct('='))) = (toks.get(i + 1), toks.get(i + 2))
        {
            // The member runs to the `;` at bracket depth 0.
            let start = i + 3;
            let mut depth = 0i32;
            let mut j = start;
            while let Some(t) = toks.get(j) {
                match t {
                    Tok::Punct('(' | '[' | '{') => depth += 1,
                    Tok::Punct(')' | ']' | '}') => depth -= 1,
                    Tok::Punct(';') if depth <= 0 => break,
                    _ => {}
                }
                j += 1;
            }
            if let Some(q) = toks.get(start..j).and_then(|body| member(name, body)) {
                out.push(q);
            }
            i = j + 1;
            continue;
        }
        i += 1;
    }
    out
}

/// What a step evaluates to.
#[derive(Clone)]
enum Step {
    /// `Sql.Database(server, db, [Query = …])`.
    Database { server: String, database: String, query: Option<String> },
    /// A table navigated to in a database step.
    Table { server: String, database: String, schema: String, name: String },
}

/// `let a = …, b = … in b`, or a bare expression.
fn member(name: &str, body: &[Tok]) -> Option<SqlQuery> {
    let mut steps: Vec<(String, Step)> = vec![];
    let result = if body.first() == Some(&Tok::Ident("let".into())) {
        let in_at = body.iter().rposition(|t| *t == Tok::Ident("in".into()))?;
        let defs = body.get(1..in_at)?;
        for def in split_top(defs, ',') {
            let (Some(Tok::Ident(step)), Some(Tok::Punct('='))) = (def.first(), def.get(1)) else { return None };
            let value = expr(def.get(2..)?, &steps)?;
            steps.push((step.clone(), value));
        }
        match body.get(in_at + 1..)? {
            [Tok::Ident(last)] => steps.iter().find(|(n, _)| n == last)?.1.clone(),
            _ => return None,
        }
    } else {
        expr(body, &steps)?
    };
    let (server, database, command) = match result {
        Step::Database { server, database, query: Some(q) } => (server, database, QueryCommand::Sql { text: q }),
        Step::Database { query: None, .. } => return None,
        Step::Table { server, database, schema, name } => (server, database, QueryCommand::Table { schema, name }),
    };
    Some(SqlQuery { name: name.to_string(), server, database, command })
}

/// Pieces of `toks` between top-level `sep`s.
fn split_top(toks: &[Tok], sep: char) -> Vec<&[Tok]> {
    let mut out = vec![];
    let (mut depth, mut from) = (0i32, 0);
    for (i, t) in toks.iter().enumerate() {
        match t {
            Tok::Punct('(' | '[' | '{') => depth += 1,
            Tok::Punct(')' | ']' | '}') => depth -= 1,
            Tok::Punct(c) if *c == sep && depth == 0 => {
                out.push(toks.get(from..i).unwrap_or(&[]));
                from = i + 1;
            }
            _ => {}
        }
    }
    out.push(toks.get(from..).unwrap_or(&[]));
    out
}

/// A record of literal fields: `[A = "x", B = true]`.
fn record(toks: &[Tok]) -> Option<Vec<(String, Tok)>> {
    let inner = match toks {
        [Tok::Punct('['), inner @ .., Tok::Punct(']')] => inner,
        _ => return None,
    };
    if inner.is_empty() {
        return Some(vec![]);
    }
    split_top(inner, ',')
        .into_iter()
        .map(|f| match f {
            [Tok::Ident(k), Tok::Punct('='), v @ (Tok::Str(_) | Tok::Ident(_))] => Some((k.clone(), v.clone())),
            _ => None,
        })
        .collect()
}

fn expr(toks: &[Tok], steps: &[(String, Step)]) -> Option<Step> {
    match toks {
        // Sql.Database("server", "db" [, options])
        [Tok::Ident(f), Tok::Punct('('), args @ .., Tok::Punct(')')] if f == "Sql.Database" => {
            let args = split_top(args, ',');
            let s = |a: Option<&&[Tok]>| match a {
                Some([Tok::Str(v)]) => Some(v.clone()),
                _ => None,
            };
            let (server, database) = (s(args.first())?, s(args.get(1))?);
            let mut query = None;
            match args.get(2) {
                None => {}
                Some(opts) => {
                    for (k, v) in record(opts)? {
                        match (k.as_str(), v) {
                            ("Query", Tok::Str(q)) => query = Some(q),
                            ("Query", _) => return None,
                            // Options that don't change what's read.
                            (
                                "CreateNavigationProperties" | "HierarchicalNavigation" | "MultiSubnetFailover" | "EnableCrossDatabaseFolding",
                                Tok::Ident(b),
                            ) if b == "true" || b == "false" => {}
                            _ => return None,
                        }
                    }
                }
            }
            if args.len() > 3 {
                return None;
            }
            Some(Step::Database { server, database, query })
        }
        // Source{[Schema = "dbo", Item = "T"]}[Data]
        [Tok::Ident(prev), Tok::Punct('{'), sel @ .., Tok::Punct('}'), Tok::Punct('['), Tok::Ident(data), Tok::Punct(']')] if data == "Data" => {
            let Step::Database { server, database, query: None } = steps.iter().find(|(n, _)| n == prev)?.1.clone() else { return None };
            let fields = record(sel)?;
            let get = |k: &str| {
                fields.iter().find(|(f, _)| f == k).and_then(|(_, v)| match v {
                    Tok::Str(s) => Some(s.clone()),
                    _ => None,
                })
            };
            if fields.iter().any(|(k, _)| k != "Schema" && k != "Item") {
                return None;
            }
            Some(Step::Table { server, database, schema: get("Schema").unwrap_or_default(), name: get("Item")? })
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_sql_statement() {
        let m = "section Section1;\r\n\r\nshared Consulta1 = let\r\n    Origen = Sql.Database(\"10.0.0.1\", \"Shop\", [Query=\"SELECT * FROM ms_devices#(lf)WHERE a = \"\"x\"\"\", CreateNavigationProperties=false])\r\nin\r\n    Origen;";
        let q = parse_section(m);
        assert_eq!(q.len(), 1);
        assert_eq!(q[0].name, "Consulta1");
        assert_eq!((q[0].server.as_str(), q[0].database.as_str()), ("10.0.0.1", "Shop"));
        assert_eq!(q[0].command, QueryCommand::Sql { text: "SELECT * FROM ms_devices\nWHERE a = \"x\"".into() });
    }

    #[test]
    fn a_navigated_table_with_quoted_names() {
        let m = r#"section Section1;
shared #"dbo Clientes" = let
    Source = Sql.Database("srv\SQL", "Ventas"),
    // the Navigator's step
    #"dbo_Clientes" = Source{[Schema="dbo",Item="Clientes"]}[Data]
in
    #"dbo_Clientes";
shared Other = 1;"#;
        let q = parse_section(m);
        assert_eq!(q.len(), 1);
        assert_eq!(q[0].name, "dbo Clientes");
        assert_eq!(q[0].command, QueryCommand::Table { schema: "dbo".into(), name: "Clientes".into() });
    }

    #[test]
    fn transformed_queries_are_left_alone() {
        let filtered = r#"section Section1;
shared Q = let
    Source = Sql.Database("s", "d"),
    T = Source{[Schema="dbo",Item="T"]}[Data],
    Filtered = Table.SelectRows(T, each [x] > 1)
in
    Filtered;"#;
        assert!(parse_section(filtered).is_empty());
        assert!(parse_section(r#"shared Q = Sql.Database("s", "d", [CommandTimeout=#duration(0,0,1,0)]);"#).is_empty());
        assert!(parse_section(r#"shared Q = Sql.Database("s", "d");"#).is_empty(), "a whole database isn't a table");
        assert!(parse_section(r#"shared Q = Csv.Document(File.Contents("x"));"#).is_empty());
        assert!(parse_section("shared Q = let in ;").is_empty());
        assert!(parse_section("\"unterminated").is_empty());
    }

    #[test]
    fn escapes() {
        assert_eq!(unescape("a#(cr,lf)b#(tab)c#(#)(d)#(0041)"), "a\r\nb\tc#(d)A");
        assert_eq!(unescape("#(nope)"), "#(nope)");
        assert_eq!(unescape("#(unclosed"), "#(unclosed");
    }

    #[test]
    fn a_whole_part() {
        // A DataMashup part as Excel writes it: UTF-16, base64 of version + length + ZIP.
        let m = "section Section1;\nshared Q1 = let\n Origen = Sql.Database(\"h\", \"d\", [Query=\"SELECT 1\"])\nin\n Origen;";
        let zip = crate::testing::zip_parts(&[("Formulas/Section1.m", m.as_bytes())]);
        let mut bin = 0u32.to_le_bytes().to_vec();
        bin.extend((zip.len() as u32).to_le_bytes());
        bin.extend(&zip);
        bin.extend([0u8; 16]);
        let xml = format!(
            "<?xml version=\"1.0\" encoding=\"utf-16\"?><DataMashup xmlns=\"http://schemas.microsoft.com/DataMashup\">{}</DataMashup>",
            crate::testing::b64(&bin)
        );
        let mut part = vec![0xFF, 0xFE];
        for u in xml.encode_utf16() {
            part.extend(u.to_le_bytes());
        }
        let q = sql_queries(&part);
        assert_eq!(q.len(), 1, "{q:?}");
        assert_eq!(q[0].command, QueryCommand::Sql { text: "SELECT 1".into() });
        assert!(sql_queries(b"<DataMashup>!!!</DataMashup>").is_empty());
        assert!(sql_queries(&[]).is_empty());
    }
}

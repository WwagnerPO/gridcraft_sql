//! Database connections (`xl/connections.xml`, ECMA-376 Part 1 §18.13) and the query tables
//! that bind them to worksheet tables (`xl/queryTables/queryTableN.xml`, §18.12).
//!
//! GridCraft models SQL Server connections read through an OLE DB provider (`<dbPr>`). A file
//! whose connections are all of that kind and each fill a table is read into
//! [`Workbook::connections`] and written back from it; any other file keeps its connections part
//! byte for byte (pass-through), and connections added in GridCraft are appended to it.

use std::fmt::Write as _;

use gridcraft_model::{Connection, DataSource, QueryCommand, SqlAuth, Workbook};

use crate::xml::{El, esc_attr};

pub const CT_CONNECTIONS: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
pub const CT_QUERY_TABLE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
/// OLE DB connection (`ST_ConnectionType` 5).
const TYPE_OLEDB: u32 = 5;
/// `ST_DbCommandType`: SQL text, table name.
const COMMAND_SQL: u32 = 2;
const COMMAND_TABLE: u32 = 3;

// ---------------------------------------------------------------- connection strings

/// A connection-string value, quoted when it holds `;`, quotes or edge spaces.
fn quote_value(v: &str) -> String {
    if v.contains([';', '"', '\'']) || v.starts_with(' ') || v.ends_with(' ') { format!("\"{}\"", v.replace('"', "\"\"")) } else { v.to_string() }
}

/// The OLE DB connection string of a connection. Never holds a password.
pub fn connection_string(c: &Connection) -> String {
    let DataSource::SqlServer { server, database, encrypt, trust_server_certificate, auth } = &c.source;
    let mut s = String::from("Provider=SQLOLEDB.1;Persist Security Info=False");
    if *auth == SqlAuth::Integrated {
        s.push_str(";Integrated Security=SSPI");
    } else if !c.user.is_empty() {
        let _ = write!(s, ";User ID={}", quote_value(&c.user));
    }
    if !database.is_empty() {
        let _ = write!(s, ";Initial Catalog={}", quote_value(database));
    }
    let _ = write!(s, ";Data Source={}", quote_value(server));
    let _ = write!(s, ";Use Encryption for Data={}", if *encrypt { "True" } else { "False" });
    if *trust_server_certificate {
        s.push_str(";Trust Server Certificate=True");
    }
    s
}

/// `key=value;…` pairs, keys lower-cased; quoted values unquoted.
pub fn parse_connection_string(s: &str) -> Vec<(String, String)> {
    let mut out = vec![];
    let mut chars = s.chars().peekable();
    loop {
        let mut key = String::new();
        let mut has_value = false;
        while let Some(&c) = chars.peek() {
            chars.next();
            if c == '=' {
                has_value = true;
                break;
            }
            if c == ';' {
                break;
            }
            key.push(c);
        }
        if !has_value {
            // A piece without `=` names nothing.
            if chars.peek().is_none() {
                break;
            }
            continue;
        }
        while chars.peek() == Some(&' ') {
            chars.next();
        }
        let mut value = String::new();
        match chars.peek().copied() {
            Some(q @ ('"' | '\'')) => {
                chars.next();
                while let Some(c) = chars.next() {
                    if c == q {
                        if chars.peek() == Some(&q) {
                            chars.next();
                            value.push(q);
                        } else {
                            break;
                        }
                    } else {
                        value.push(c);
                    }
                }
                // Up to the separator.
                for c in chars.by_ref() {
                    if c == ';' {
                        break;
                    }
                }
            }
            _ => {
                for c in chars.by_ref() {
                    if c == ';' {
                        break;
                    }
                    value.push(c);
                }
                value = value.trim().to_string();
            }
        }
        let key = key.trim().to_lowercase();
        if !key.is_empty() {
            out.push((key, value));
        }
        if chars.peek().is_none() {
            break;
        }
    }
    out
}

fn truthy(v: &str) -> bool {
    matches!(v.trim().to_ascii_lowercase().as_str(), "true" | "yes" | "1" | "mandatory" | "strict" | "sspi")
}

// ---------------------------------------------------------------- commands

fn quote_identifier(name: &str) -> String {
    format!("\"{}\"", name.replace('"', "\"\""))
}

fn command(c: &Connection) -> (String, u32) {
    match &c.command {
        QueryCommand::Sql { text } => (text.clone(), COMMAND_SQL),
        QueryCommand::Table { schema, name } => {
            let mut parts = vec![];
            if !c.source.database().is_empty() {
                parts.push(quote_identifier(c.source.database()));
            }
            if !schema.is_empty() {
                parts.push(quote_identifier(schema));
            }
            parts.push(quote_identifier(name));
            (parts.join("."), COMMAND_TABLE)
        }
    }
}

/// `"db"."dbo"."Customers"`, `[dbo].[Customers]`, `dbo.Customers` → the dot-separated parts.
fn identifier_parts(text: &str) -> Option<Vec<String>> {
    let mut parts = vec![];
    let mut cur = String::new();
    let mut chars = text.trim().chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '[' | '"' => {
                let close = if c == '[' { ']' } else { '"' };
                loop {
                    match chars.next() {
                        Some(x) if x == close => {
                            if chars.peek() == Some(&close) {
                                chars.next();
                                cur.push(close);
                            } else {
                                break;
                            }
                        }
                        Some(x) => cur.push(x),
                        None => return None,
                    }
                }
            }
            '.' => parts.push(std::mem::take(&mut cur)),
            c => cur.push(c),
        }
    }
    parts.push(cur);
    let parts: Vec<String> = parts.into_iter().map(|p| p.trim().to_string()).collect();
    (!parts.is_empty() && parts.len() <= 3 && parts.last().is_some_and(|p| !p.is_empty())).then_some(parts)
}

// ---------------------------------------------------------------- writing

/// The connection id each query table refers to: (table name, lower case → id), ids from
/// `first_id` in [`Workbook::connections`] order.
pub fn ids(wb: &Workbook, first_id: u32) -> Vec<(String, u32)> {
    wb.connections.iter().enumerate().map(|(i, c)| (c.table.to_lowercase(), first_id.saturating_add(i as u32))).collect()
}

/// The `<connection>` elements of the workbook's connections, ids from `first_id`.
pub fn connection_elements(wb: &Workbook, first_id: u32) -> String {
    let mut s = String::new();
    for (i, c) in wb.connections.iter().enumerate() {
        let id = first_id.saturating_add(i as u32);
        let _ = write!(s, "<connection id=\"{id}\"");
        if c.refresh_minutes > 0 {
            let _ = write!(s, " interval=\"{}\"", c.refresh_minutes);
        }
        let _ = write!(s, " name=\"{}\"", esc_attr(&c.name));
        if !c.description.is_empty() {
            let _ = write!(s, " description=\"{}\"", esc_attr(&c.description));
        }
        let _ = write!(s, " type=\"{TYPE_OLEDB}\" refreshedVersion=\"8\" background=\"1\"");
        if c.refresh_on_open {
            s.push_str(" refreshOnLoad=\"1\"");
        }
        s.push_str(" saveData=\"1\">");
        let (cmd, ty) = command(c);
        let _ = write!(s, "<dbPr connection=\"{}\" command=\"{}\" commandType=\"{ty}\"/>", esc_attr(&connection_string(c)), esc_attr(&cmd));
        s.push_str("</connection>");
    }
    s
}

/// A whole connections part for the workbook's connections.
pub fn connections_part(wb: &Workbook) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<connections xmlns=\"{NS_MAIN}\">{}</connections>",
        connection_elements(wb, 1)
    )
}

/// The query table part binding a table's columns to connection `connection_id`.
pub fn query_table_xml(n: u32, connection_id: u32, columns: &[String]) -> String {
    let mut s = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<queryTable xmlns=\"{NS_MAIN}\" name=\"ExternalData_{n}\" connectionId=\"{connection_id}\" autoFormatId=\"16\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\" applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"0\">"
    );
    let _ = write!(s, "<queryTableRefresh nextId=\"{}\"><queryTableFields count=\"{}\">", columns.len() + 1, columns.len());
    for (i, c) in columns.iter().enumerate() {
        let _ = write!(s, "<queryTableField id=\"{0}\" name=\"{1}\" tableColumnId=\"{0}\"/>", i + 1, esc_attr(c));
    }
    s.push_str("</queryTableFields></queryTableRefresh></queryTable>");
    s
}

/// The highest `<connection id>` in a connections part.
pub fn max_id(xml: &[u8]) -> u32 {
    crate::xml::parse(xml).ok().map(|x| x.kids("connection").filter_map(|c| c.attr_u32("id")).max().unwrap_or(0)).unwrap_or(0)
}

/// A kept connections part with `elements` appended. `None` when the part has no closing tag to
/// append before.
pub fn merge(existing: &[u8], elements: &str) -> Option<Vec<u8>> {
    let text = std::str::from_utf8(existing).ok()?;
    let at = text.rfind("</")?;
    let tail = text.get(at..)?;
    if !tail.trim_end().ends_with("connections>") {
        return None;
    }
    let mut out = String::with_capacity(text.len() + elements.len());
    out.push_str(text.get(..at)?);
    out.push_str(elements);
    out.push_str(tail);
    Some(out.into_bytes())
}

// ---------------------------------------------------------------- reading

/// The connections of a connections part, when GridCraft models every one of them: SQL Server
/// through OLE DB, filling one of the tables in `query_tables` (table name → connection id).
/// `None` keeps the part as it is.
pub fn read(x: &El, query_tables: &[(String, u32)]) -> Option<Vec<Connection>> {
    let mut out = vec![];
    for c in x.kids("connection") {
        let id = c.attr_u32("id")?;
        if c.flag("deleted", false) {
            return None;
        }
        let db = c.child("dbPr")?;
        let kv = parse_connection_string(db.attr("connection")?);
        let get = |k: &str| kv.iter().find(|(key, _)| key == k).map(|(_, v)| v.as_str());
        let provider = get("provider")?.to_ascii_lowercase();
        if !(provider.starts_with("sqloledb") || provider.starts_with("msoledbsql") || provider.starts_with("sqlncli")) {
            return None;
        }
        let integrated = get("integrated security").is_some_and(truthy) || get("trusted_connection").is_some_and(truthy);
        let user = get("user id").or_else(|| get("uid")).unwrap_or("").to_string();
        let server = get("data source").or_else(|| get("server")).filter(|v| !v.trim().is_empty())?.trim().to_string();
        let database = get("initial catalog").or_else(|| get("database")).unwrap_or("").trim().to_string();
        let text = db.attr("command").unwrap_or("");
        let command = match db.attr_u32("commandType").unwrap_or(COMMAND_SQL) {
            COMMAND_SQL if !text.trim().is_empty() => QueryCommand::Sql { text: text.to_string() },
            COMMAND_TABLE => {
                let parts = identifier_parts(text)?;
                match parts.as_slice() {
                    [name] => QueryCommand::Table { schema: String::new(), name: name.clone() },
                    [schema, name] => QueryCommand::Table { schema: schema.clone(), name: name.clone() },
                    [_db, schema, name] => QueryCommand::Table { schema: schema.clone(), name: name.clone() },
                    _ => return None,
                }
            }
            _ => return None,
        };
        let table = query_tables.iter().find(|(_, cid)| *cid == id).map(|(t, _)| t.clone())?;
        out.push(Connection {
            id,
            name: c.attr("name").unwrap_or("").to_string(),
            description: c.attr("description").unwrap_or("").to_string(),
            source: DataSource::SqlServer {
                server,
                database,
                encrypt: get("use encryption for data").or_else(|| get("encrypt")).is_some_and(truthy),
                trust_server_certificate: get("trust server certificate").or_else(|| get("trustservercertificate")).is_some_and(truthy),
                // A `DOMAIN\name` login signs in to Windows; SQL Server logins have no backslash.
                auth: if integrated {
                    SqlAuth::Integrated
                } else if user.contains('\\') {
                    SqlAuth::Windows
                } else {
                    SqlAuth::Sql
                },
            },
            user,
            command,
            table,
            refresh_on_open: c.flag("refreshOnLoad", false),
            refresh_minutes: c.attr_u32("interval").unwrap_or(0).min(32_767),
            refreshed: None,
            rows: 0,
        });
    }
    // Names must be unique for Queries & Connections.
    for i in 0..out.len() {
        let taken = |n: &str, out: &[Connection]| out.iter().take(i).any(|c| c.name.eq_ignore_ascii_case(n));
        if out.get(i).is_some_and(|c| c.name.is_empty() || taken(&c.name, &out)) {
            let base = out.get(i).map(|c| if c.name.is_empty() { "Query".to_string() } else { c.name.clone() }).unwrap_or_default();
            let name = (2..).map(|k| format!("{base} ({k})")).find(|n| !taken(n, &out)).unwrap_or(base);
            if let Some(c) = out.get_mut(i) {
                c.name = name;
            }
        }
    }
    (!out.is_empty()).then_some(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn conn() -> Connection {
        Connection {
            id: 1,
            name: "dbo Customers".into(),
            description: String::new(),
            source: DataSource::SqlServer {
                server: "db;x".into(),
                database: "Shop".into(),
                encrypt: true,
                trust_server_certificate: true,
                auth: SqlAuth::Sql,
            },
            user: "sa".into(),
            command: QueryCommand::Table { schema: "dbo".into(), name: "Order \"Lines\"".into() },
            table: "dbo_Customers".into(),
            refresh_on_open: true,
            refresh_minutes: 15,
            refreshed: None,
            rows: 0,
        }
    }

    #[test]
    fn connection_strings_round_trip() {
        let c = conn();
        let s = connection_string(&c);
        assert!(!s.to_lowercase().contains("password"));
        let kv = parse_connection_string(&s);
        let get = |k: &str| kv.iter().find(|(key, _)| key == k).map(|(_, v)| v.clone());
        assert_eq!(get("data source").as_deref(), Some("db;x"));
        assert_eq!(get("initial catalog").as_deref(), Some("Shop"));
        assert_eq!(get("user id").as_deref(), Some("sa"));
        let kv = parse_connection_string("Provider=X; Data Source = 'a''b' ;Empty=;");
        assert_eq!(kv, vec![("provider".into(), "X".into()), ("data source".into(), "a'b".into()), ("empty".into(), "".into())]);
        assert!(parse_connection_string("").is_empty());
        assert_eq!(parse_connection_string(";;=;\""), vec![]);
    }

    #[test]
    fn parts_read_back() {
        let mut wb = Workbook::new();
        wb.connections.push(conn());
        let mut sql = conn();
        sql.name = "Query1".into();
        sql.table = "Query1".into();
        sql.command = QueryCommand::Sql { text: "SELECT 1 AS \"a<b\"".into() };
        sql.refresh_on_open = false;
        sql.refresh_minutes = 0;
        wb.connections.push(sql);
        let part = connections_part(&wb);
        let x = crate::xml::parse(part.as_bytes()).unwrap();
        let tables = vec![("dbo_Customers".to_string(), 1), ("Query1".to_string(), 2)];
        let back = read(&x, &tables).unwrap();
        assert_eq!(back.len(), 2);
        assert_eq!(back[0], Connection { refreshed: None, rows: 0, ..conn() });
        assert_eq!(back[1].command, QueryCommand::Sql { text: "SELECT 1 AS \"a<b\"".into() });
        // A connection with no table to fill isn't modelled: the part is kept as it is.
        assert!(read(&x, &tables[..1]).is_none());
        assert_eq!(max_id(part.as_bytes()), 2);
        let merged = merge(part.as_bytes(), &connection_elements(&wb, 3)).unwrap();
        assert_eq!(max_id(&merged), 4);
        assert!(merge(b"<x/>", "").is_none());
    }

    #[test]
    fn other_connections_are_left_alone() {
        let mashup = r#"<connections xmlns="x"><connection id="1" name="Query - T" type="5"><dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=T" command="SELECT * FROM [T]"/></connection></connections>"#;
        let x = crate::xml::parse(mashup.as_bytes()).unwrap();
        assert!(read(&x, &[("t".into(), 1)]).is_none());
        let windows = r#"<connections xmlns="x"><connection id="1" name="q" type="5"><dbPr connection="Provider=SQLOLEDB.1;Integrated Security=SSPI;Data Source=s" command="SELECT 1" commandType="2"/></connection></connections>"#;
        let x = crate::xml::parse(windows.as_bytes()).unwrap();
        let c = read(&x, &[("t".into(), 1)]).unwrap();
        assert_eq!(c[0].source.auth(), SqlAuth::Integrated);
        assert!(connection_string(&c[0]).contains("Integrated Security=SSPI"));
    }

    #[test]
    fn identifiers() {
        assert_eq!(identifier_parts("\"Shop\".\"dbo\".\"T\""), Some(vec!["Shop".into(), "dbo".into(), "T".into()]));
        assert_eq!(identifier_parts("[dbo].[Order]]s]"), Some(vec!["dbo".into(), "Order]s".into()]));
        assert_eq!(identifier_parts("a.b.c.d"), None);
        assert_eq!(identifier_parts("\"open"), None);
    }

    #[test]
    fn query_table_part() {
        let x = query_table_xml(1, 7, &["Id".into(), "A&B".into()]);
        let el = crate::xml::parse(x.as_bytes()).unwrap();
        assert_eq!(el.attr_u32("connectionId"), Some(7));
        assert!(x.contains("name=\"A&amp;B\""));
    }
}

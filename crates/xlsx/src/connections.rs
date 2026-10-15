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
pub fn ids(wb: &Workbook, first_id: u32, kept: bool) -> Vec<(String, u32)> {
    let mut next = first_id;
    wb.connections
        .iter()
        .map(|c| {
            // A Power Query connection keeps its id in the kept connections part.
            let id = if kept && c.kept {
                c.id
            } else {
                next = next.saturating_add(1);
                next - 1
            };
            (c.table.to_lowercase(), id)
        })
        .collect()
}

/// The `<connection>` elements of the workbook's connections, ids from `first_id`.
pub fn connection_elements(wb: &Workbook, first_id: u32, kept: bool) -> String {
    let mut s = String::new();
    // Power Query connections stay as they are in a kept part; without one they are written as
    // plain SQL Server connections.
    for (i, c) in wb.connections.iter().filter(|c| !(kept && c.kept)).enumerate() {
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
        connection_elements(wb, 1, false)
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

// ---------------------------------------------------------------- package glue
//
// The readers and writers of the package call these with one line each, so the shared files
// carry as little of the connections code as possible.

/// What writing a workbook tracks for its connections.
#[derive(Default)]
pub struct WriteState {
    /// The connection id of each table a query fills (table name, lower case → id).
    query_ids: Vec<(String, u32)>,
    query_tables: u32,
    /// The workbook keeps a connections part from its file.
    kept: bool,
}

fn is_connections(rel_type: &str) -> bool {
    rel_type.rsplit('/').next().is_some_and(|k| k.eq_ignore_ascii_case("connections"))
}

/// The bytes of the kept connections part, if the workbook has one.
fn kept_part(wb: &Workbook) -> Option<&[u8]> {
    let p = wb.preserved.as_deref()?;
    let r = p.rels.iter().find(|r| !r.external && is_connections(&r.rel_type))?;
    let target = crate::package::normalize(&r.target);
    p.parts.iter().find(|q| crate::package::normalize(&q.name) == target).map(|q| &q.data[..])
}

/// Reads the workbook's connections (and the Power Query definitions that only read SQL
/// Server); the ones GridCraft models are no longer kept as pass-through.
pub(crate) fn read_workbook(cx: &mut crate::read::Ctx<'_>, wb_rels: &[crate::package::Rel], wb: &mut Workbook) {
    let mut mashup = vec![];
    for r in wb_rels.iter().filter(|r| r.kind.eq_ignore_ascii_case("customXml")) {
        if let Ok(Some(b)) = cx.pkg.read_limited(&r.target, 32 << 20) {
            mashup.extend(crate::mashup::sql_queries(&b));
        }
    }
    let Some(r) = crate::read::find_rel(wb_rels, "connections").cloned() else { return };
    let Ok(Some(x)) = cx.pkg.read_xml(&r.target) else { return };
    let Some(list) = read(&x, &cx.query_tables, &mashup) else { return };
    wb.connections = list;
    // A Power Query keeps its connection (and M) in the file, for Excel.
    let keep = wb.connections.iter().any(|c| c.power_query.is_some());
    for c in wb.connections.iter_mut() {
        c.kept = keep;
    }
    if !keep && let Some(p) = wb.preserved.as_mut() {
        let p = std::sync::Arc::make_mut(p);
        p.rels.retain(|k| k.external || !k.rel_type.ends_with("/connections"));
        p.parts.retain(|q| crate::package::normalize(&q.name) != crate::package::normalize(&r.target));
    }
    let rows: Vec<u64> = wb
        .connections
        .iter()
        .map(|c| wb.table(&c.table).and_then(|(si, ti)| wb.sheets.get(si)?.tables.get(ti)?.data_range()).map_or(0, |d| d.height() as u64))
        .collect();
    for (c, n) in wb.connections.iter_mut().zip(rows) {
        c.rows = n;
    }
}

/// A query table (`table` at part `target`) names the connection that fills it.
pub(crate) fn read_table(cx: &mut crate::read::Ctx<'_>, table: &El, target: &str, name: &str) -> Result<(), crate::IoError> {
    if table.attr("tableType") == Some("queryTable")
        && let Some(q) = cx.pkg.rels(target)?.into_iter().find(|r| r.kind == "queryTable")
        && let Ok(Some(qx)) = cx.pkg.read_xml(&q.target)
        && let Some(id) = qx.attr_u32("connectionId")
    {
        cx.query_tables.push((name.to_string(), id));
    }
    Ok(())
}

/// Before the kept parts are written: the ids connections get (GridCraft's follow those a kept
/// connections part already holds).
pub(crate) fn begin_write(wb: &Workbook, out: &mut crate::write::Out) {
    let kept = kept_part(wb);
    out.connections.kept = kept.is_some();
    out.connections.query_ids = ids(wb, kept.map_or(1, |x| max_id(x).saturating_add(1)), kept.is_some());
}

/// After the kept parts: a connections part when the file had none.
pub(crate) fn write_part(wb: &Workbook, out: &mut crate::write::Out, wb_rels: &mut crate::write::Rels) {
    if !wb.connections.is_empty() && !out.connections.kept {
        out.part("xl/connections.xml", Some(CT_CONNECTIONS), connections_part(wb).into_bytes());
        wb_rels.add("connections", "connections.xml");
    }
}

/// Connections added in GridCraft join a kept connections part (`parts` are the kept parts
/// about to be written).
pub(crate) fn merge_kept(wb: &Workbook, rels: &[gridcraft_model::PreservedRel], parts: &mut [gridcraft_model::PreservedPart]) {
    if !wb.connections.iter().any(|c| !c.kept) {
        return;
    }
    let Some(r) = rels.iter().find(|r| !r.external && is_connections(&r.rel_type)) else { return };
    let target = crate::package::normalize(&r.target);
    if let Some(q) = parts.iter_mut().find(|q| crate::package::normalize(&q.name) == target) {
        let first = max_id(&q.data).saturating_add(1);
        if let Some(merged) = merge(&q.data, &connection_elements(wb, first, true)) {
            q.data = merged.into();
        }
    }
}

/// Makes table `n` (its XML just written) a query table when a query fills it: marks the table
/// and its columns, and writes the query table part it points at.
pub(crate) fn bind_table(t: &gridcraft_model::Table, cols: &[String], n: u32, xml: String, out: &mut crate::write::Out) -> String {
    let name = t.name.to_lowercase();
    let Some(cid) = out.connections.query_ids.iter().find(|(tn, _)| *tn == name).map(|(_, id)| *id).filter(|_| t.table_type.is_none()) else {
        return xml;
    };
    out.connections.query_tables += 1;
    let q = out.connections.query_tables;
    out.part(&format!("xl/queryTables/queryTable{q}.xml"), Some(CT_QUERY_TABLE), query_table_xml(q, cid, cols).into_bytes());
    let mut trels = crate::write::Rels::default();
    trels.add("queryTable", &format!("../queryTables/queryTable{q}.xml"));
    out.part(&format!("xl/tables/_rels/table{n}.xml.rels"), None, trels.xml().into_bytes());
    query_table_markup(&xml)
}

/// `tableType="queryTable"` on the table, and each column's `uniqueName` (unless it has one) and
/// `queryTableFieldId`.
fn query_table_markup(xml: &str) -> String {
    let s = xml.replacen("<table xmlns=", "<table tableType=\"queryTable\" xmlns=", 1);
    let mut out = String::with_capacity(s.len() + 64);
    let mut rest = s.as_str();
    const OPEN: &str = "<tableColumn id=\"";
    while let Some(at) = rest.find(OPEN) {
        let (head, tail) = rest.split_at(at + OPEN.len());
        out.push_str(head);
        let Some(end) = tail.find('"') else {
            rest = tail;
            break;
        };
        let (id, after) = tail.split_at(end + 1);
        out.push_str(id);
        let id = id.trim_end_matches('"');
        if !after.starts_with(" uniqueName=") {
            let _ = write!(out, " uniqueName=\"{id}\"");
        }
        let _ = write!(out, " queryTableFieldId=\"{id}\"");
        rest = after;
    }
    out.push_str(rest);
    out
}

// ---------------------------------------------------------------- reading

/// The connections of a connections part, when GridCraft models every one of them: SQL Server
/// through OLE DB, filling one of the tables in `query_tables` (table name → connection id).
/// `None` keeps the part as it is.
pub fn read(x: &El, query_tables: &[(String, u32)], mashup: &[crate::mashup::SqlQuery]) -> Option<Vec<Connection>> {
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
        // A Power Query connection: its query, if it only reads SQL Server.
        if provider.starts_with("microsoft.mashup.oledb") {
            let location = get("location")?;
            let q = mashup.iter().find(|q| q.name == location)?;
            let table = query_tables.iter().find(|(_, cid)| *cid == id).map(|(t, _)| t.clone())?;
            out.push(Connection {
                id,
                name: c.attr("name").unwrap_or(location).to_string(),
                description: c.attr("description").unwrap_or("").to_string(),
                source: DataSource::SqlServer {
                    server: q.server.clone(),
                    database: q.database.clone(),
                    encrypt: true,
                    trust_server_certificate: false,
                    auth: SqlAuth::Sql,
                },
                user: String::new(),
                command: q.command.clone(),
                table,
                refresh_on_open: c.flag("refreshOnLoad", false),
                refresh_minutes: c.attr_u32("interval").unwrap_or(0).min(32_767),
                refreshed: None,
                rows: 0,
                power_query: Some(location.to_string()),
                kept: false,
            });
            continue;
        }
        let power_query = None;
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
            power_query,
            kept: false,
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
            power_query: None,
            kept: false,
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
        let back = read(&x, &tables, &[]).unwrap();
        assert_eq!(back.len(), 2);
        assert_eq!(back[0], Connection { refreshed: None, rows: 0, ..conn() });
        assert_eq!(back[1].command, QueryCommand::Sql { text: "SELECT 1 AS \"a<b\"".into() });
        // A connection with no table to fill isn't modelled: the part is kept as it is.
        assert!(read(&x, &tables[..1], &[]).is_none());
        assert_eq!(max_id(part.as_bytes()), 2);
        let merged = merge(part.as_bytes(), &connection_elements(&wb, 3, true)).unwrap();
        assert_eq!(max_id(&merged), 4);
        assert!(merge(b"<x/>", "").is_none());
    }

    #[test]
    fn other_connections_are_left_alone() {
        let mashup = r#"<connections xmlns="x"><connection id="1" name="Query - T" type="5"><dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=T" command="SELECT * FROM [T]"/></connection></connections>"#;
        let x = crate::xml::parse(mashup.as_bytes()).unwrap();
        assert!(read(&x, &[("t".into(), 1)], &[]).is_none());
        let windows = r#"<connections xmlns="x"><connection id="1" name="q" type="5"><dbPr connection="Provider=SQLOLEDB.1;Integrated Security=SSPI;Data Source=s" command="SELECT 1" commandType="2"/></connection></connections>"#;
        let x = crate::xml::parse(windows.as_bytes()).unwrap();
        let c = read(&x, &[("t".into(), 1)], &[]).unwrap();
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

#[cfg(test)]
mod power_query_tests {
    use super::*;
    use gridcraft_core::{CellRef, RangeRef};
    use gridcraft_model::{Table, TableColumn, TotalsFn};
    use std::io::Read as _;

    /// A workbook as Excel saves a Power Query over SQL Server: a mashup connection, its M in a
    /// DataMashup custom XML part, and a query table bound to the worksheet table.
    fn excel_power_query_file() -> Vec<u8> {
        let mut wb = Workbook::new();
        {
            let sh = wb.sheet_mut(0).unwrap();
            for (c, h) in ["device_id", "os"].iter().enumerate() {
                sh.set_value(CellRef::new(0, c as u32), gridcraft_core::Value::text(*h));
                sh.set_value(CellRef::new(1, c as u32), gridcraft_core::Value::text("x"));
            }
            let col =
                |n: &str| TableColumn { name: n.into(), totals: TotalsFn::None, totals_label: None, formula: None, unique_name: None, xml: None };
            sh.tables.push(Table {
                id: 1,
                name: "Consulta1".into(),
                range: RangeRef::parse("A1:B2").unwrap(),
                header_row: true,
                totals_row: false,
                columns: vec![col("device_id"), col("os")],
                style: "TableStyleMedium2".into(),
                banded_rows: true,
                banded_cols: false,
                first_col: false,
                last_col: false,
                filter_button: true,
                filter: vec![],
                table_type: None,
                connection_id: None,
            });
        }
        let plain = crate::write_xlsx(&wb).unwrap();
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(plain)).unwrap();
        let mut parts: Vec<(String, Vec<u8>)> = vec![];
        for i in 0..zip.len() {
            let mut f = zip.by_index(i).unwrap();
            let mut b = vec![];
            f.read_to_end(&mut b).unwrap();
            parts.push((f.name().to_string(), b));
        }
        let edit = |parts: &mut Vec<(String, Vec<u8>)>, name: &str, f: &dyn Fn(String) -> String| {
            let p = parts.iter_mut().find(|(n, _)| n == name).unwrap();
            p.1 = f(String::from_utf8(p.1.clone()).unwrap()).into_bytes();
        };
        edit(&mut parts, "xl/tables/table1.xml", &|t| t.replacen("<table ", "<table tableType=\"queryTable\" ", 1));
        edit(&mut parts, "xl/_rels/workbook.xml.rels", &|t| {
            t.replace(
                "</Relationships>",
                "<Relationship Id=\"rIdC\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections\" Target=\"connections.xml\"/><Relationship Id=\"rIdX\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/item1.xml\"/></Relationships>",
            )
        });
        edit(&mut parts, "[Content_Types].xml", &|t| {
            t.replace("</Types>", &format!("<Override PartName=\"/xl/connections.xml\" ContentType=\"{CT_CONNECTIONS}\"/><Override PartName=\"/xl/queryTables/queryTable1.xml\" ContentType=\"{CT_QUERY_TABLE}\"/></Types>"))
        });
        parts.push((
            "xl/connections.xml".into(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection id="1" keepAlive="1" name="Consulta - Consulta1" type="5" refreshedVersion="8" background="1" saveData="1"><dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Consulta1;Extended Properties=&quot;&quot;" command="SELECT * FROM [Consulta1]"/></connection></connections>"#.to_vec(),
        ));
        parts.push(("xl/queryTables/queryTable1.xml".into(), query_table_xml(1, 1, &["device_id".into(), "os".into()]).into_bytes()));
        parts.push((
            "xl/tables/_rels/table1.xml.rels".into(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="../queryTables/queryTable1.xml"/></Relationships>"#.to_vec(),
        ));
        let m = "section Section1;\r\nshared Consulta1 = let\r\n    Origen = Sql.Database(\"10.0.0.5\", \"IT\", [Query=\"SELECT * FROM devices\", CreateNavigationProperties=false])\r\nin\r\n    Origen;";
        let inner = crate::testing::zip_parts(&[("Formulas/Section1.m", m.as_bytes())]);
        let mut bin = 0u32.to_le_bytes().to_vec();
        bin.extend((inner.len() as u32).to_le_bytes());
        bin.extend(&inner);
        let xml = format!(
            "<?xml version=\"1.0\" encoding=\"utf-16\"?><DataMashup xmlns=\"http://schemas.microsoft.com/DataMashup\">{}</DataMashup>",
            crate::testing::b64(&bin)
        );
        let mut item = vec![0xFF, 0xFE];
        for u in xml.encode_utf16() {
            item.extend(u.to_le_bytes());
        }
        parts.push(("customXml/item1.xml".into(), item));
        let refs: Vec<(&str, &[u8])> = parts.iter().map(|(n, b)| (n.as_str(), b.as_slice())).collect();
        crate::testing::zip_parts(&refs)
    }

    #[test]
    fn a_power_query_over_sql_server_is_a_query_and_survives_saving() {
        let (wb, _) = crate::read_xlsx(&excel_power_query_file()).unwrap();
        assert_eq!(wb.connections.len(), 1);
        let c = &wb.connections[0];
        assert_eq!((c.source.server(), c.source.database(), c.table.as_str()), ("10.0.0.5", "IT", "Consulta1"));
        assert_eq!(c.command, QueryCommand::Sql { text: "SELECT * FROM devices".into() });
        assert_eq!(c.power_query.as_deref(), Some("Consulta1"));
        // Saved, Excel still finds its Power Query: the mashup connection, the M and the binding.
        let saved = crate::write_xlsx(&wb).unwrap();
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(saved.clone())).unwrap();
        let mut read = |n: &str| {
            let mut s = String::new();
            zip.by_name(n).unwrap().read_to_string(&mut s).unwrap();
            s
        };
        let conns = read("xl/connections.xml");
        assert_eq!(conns.matches("<connection ").count(), 1, "{conns}");
        assert!(conns.contains("Microsoft.Mashup.OleDb.1"));
        assert!(read("xl/queryTables/queryTable1.xml").contains("connectionId=\"1\""));
        assert!(zip.by_name("customXml/item1.xml").is_ok());
        // And reopened, it is still the same query.
        let (again, _) = crate::read_xlsx(&saved).unwrap();
        assert_eq!(again.connections.len(), 1);
        assert_eq!(again.connections[0].power_query.as_deref(), Some("Consulta1"));
    }
}

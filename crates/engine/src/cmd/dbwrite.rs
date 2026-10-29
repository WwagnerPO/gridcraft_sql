//! Writing to a database: Run SQL, Export to SQL Server and Save to Database.
//!
//! Every write runs in one transaction on the server, all or nothing, and never touches the
//! workbook's undo history (a database can't be undone from a spreadsheet). The commands aren't
//! journaled, so replaying a journal or a recorded script never writes to a database again.
//! `dryRun` returns what would be sent without connecting for writing.

use gridcraft_core::{CellRef, RangeRef, Value};
use gridcraft_model::{DataSource, QueryCommand, Sheet, Workbook};
use serde_json::{Value as Json, json};

use super::connections::{credential_for, db_error, find_connection, object_label, provider, refresh_one, source_param};
use super::{CommandSpec, EngineError, Result, Session, bad, bool_param, cmd, has_doc, split_sheet, str_param, target_range, target_sheet};
use crate::datasource::{self, Credential, DbColumn, DbValue, Statement};

pub fn specs() -> Vec<CommandSpec> {
    vec![
        cmd!(
            query "data.sqlExecute",
            "Run SQL",
            ["Data", "Queries & Connections"],
            None,
            "{sql, connection? (its server and database) | server, database, user?, password?, encrypt?, trustServerCertificate?, params?: [values for @P1…], cells?: [\"B2\", …] (values for @P1…)} runs in one transaction → {rowsAffected}",
            has_doc,
            sql_execute
        ),
        cmd!(
            query "data.exportToSqlServer",
            "Export to SQL Server",
            ["Data", "Queries & Connections"],
            None,
            "{table: \"dbo.Sales\", range? (default: the table or region at the active cell; first row = column names), mode?: create (default) | append | replace (deletes the table's rows first), connection? | server, database, user?, password?, …, dryRun?} → {table, mode, rows, columns, statements}",
            has_doc,
            export
        ),
        cmd!(
            query "data.saveToDatabase",
            "Save to Database",
            ["Data", "Queries & Connections"],
            None,
            "{table? | connection? (default: the table at the active cell), dryRun?} sends the rows added, changed and deleted since the last load or refresh (matched by primary key), then refreshes → {inserted, updated, deleted}",
            has_doc,
            save
        ),
    ]
}

/// Most cells a loaded query remembers for Save to Database; larger loads can't be saved back.
pub const MAX_SNAPSHOT_CELLS: usize = 2_000_000;
/// SQL Server takes at most 2,100 parameters per statement.
const MAX_PARAMS: usize = 2_000;

/// The rows a query loaded, kept so Save to Database can tell what changed.
#[derive(Clone, Debug)]
pub struct Snapshot {
    pub doc: u64,
    pub connection: u32,
    pub columns: Vec<String>,
    pub rows: Vec<Vec<DbValue>>,
}

/// Remembers what connection `conn` just loaded into the active workbook.
pub(crate) fn remember(s: &mut Session, conn: u32, columns: &[String], rows: &[Vec<DbValue>]) {
    let Some(doc) = s.active().map(|d| d.uid) else { return };
    s.data.snapshots.retain(|x| !(x.doc == doc && x.connection == conn));
    if rows.len().saturating_mul(columns.len().max(1)) <= MAX_SNAPSHOT_CELLS {
        s.data.snapshots.push(Snapshot { doc, connection: conn, columns: columns.to_vec(), rows: rows.to_vec() });
    }
}

/// Where a write goes: the named connection's (or the one at the active cell's) server, else the
/// `server`/`database` params.
fn target(s: &mut Session, p: &Json, cmd: &str) -> Result<(DataSource, Credential)> {
    let by_connection = str_param(p, "connection").is_some() || (str_param(p, "server").is_none() && find_connection(s, p, cmd).is_ok());
    if by_connection {
        let ci = find_connection(s, p, cmd)?;
        let source = s.doc()?.wb.connections.get(ci).map(|c| c.source.clone()).ok_or_else(|| bad(cmd, "no such query"))?;
        let credential = credential_for(s, &source, p)?;
        return Ok((source, credential));
    }
    let source = source_param(cmd, p)?;
    let credential = credential_for(s, &source, p)?;
    Ok((source, credential))
}

// ---------------------------------------------------------------- values

/// What a column holds, for converting worksheet values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Kind {
    Text,
    Number,
    Bool,
    Date,
    DateTime,
    Time,
}

fn kind_of_sql(t: &str) -> Kind {
    match t.to_ascii_lowercase().as_str() {
        "bit" => Kind::Bool,
        "date" => Kind::Date,
        "time" => Kind::Time,
        "datetime" | "datetime2" | "smalldatetime" | "datetimeoffset" => Kind::DateTime,
        "tinyint" | "smallint" | "int" | "bigint" | "float" | "real" | "decimal" | "numeric" | "money" | "smallmoney" => Kind::Number,
        _ => Kind::Text,
    }
}

fn kind_of_value(v: &DbValue) -> Option<Kind> {
    Some(match v {
        DbValue::Null => return None,
        DbValue::Bool(_) => Kind::Bool,
        DbValue::Number(_) => Kind::Number,
        DbValue::Text(_) => Kind::Text,
        DbValue::Date(_) => Kind::Date,
        DbValue::DateTime(_) => Kind::DateTime,
        DbValue::Time(_) => Kind::Time,
    })
}

/// A worksheet cell as a database value for a column of `kind`. Errors name the cell.
fn cell_value(wb: &Workbook, sh: &Sheet, at: CellRef, kind: Kind) -> Result<DbValue> {
    let v = sh.value(at);
    Ok(match (&v, kind) {
        (Value::Empty, _) => DbValue::Null,
        (Value::Error(e), _) => return Err(EngineError::Other(format!("Cell {} holds the error {}; fix it before writing.", at.a1(), e))),
        (Value::Text(t), _) if t.is_empty() => DbValue::Null,
        (Value::Number(n), Kind::Date) => DbValue::Date(*n),
        (Value::Number(n), Kind::DateTime) => DbValue::DateTime(*n),
        (Value::Number(n), Kind::Time) => DbValue::Time(*n),
        (Value::Number(n), Kind::Number | Kind::Bool) => DbValue::Number(*n),
        (Value::Bool(b), Kind::Bool | Kind::Number) => DbValue::Bool(*b),
        (_, Kind::Text) => DbValue::Text(crate::display::cell_text(wb, sh, at)),
        (Value::Text(t), _) => DbValue::Text(t.to_string()),
        (other, _) => DbValue::Text(other.display()),
    })
}

/// The cell's value with a date kind when its number format shows a date or a time.
fn natural_value(wb: &Workbook, sh: &Sheet, at: CellRef) -> DbValue {
    let v = sh.value(at);
    match v {
        Value::Empty => DbValue::Null,
        Value::Number(n) => {
            let code = wb.styles.get(sh.style_id(at)).num_fmt.0.clone();
            match gridcraft_numfmt::NumberFormat::parse_cached(&code).kind() {
                gridcraft_numfmt::FormatKind::Time => DbValue::Time(n),
                gridcraft_numfmt::FormatKind::Date if n.fract() == 0.0 => DbValue::Date(n),
                gridcraft_numfmt::FormatKind::Date => DbValue::DateTime(n),
                _ => DbValue::Number(n),
            }
        }
        Value::Bool(b) => DbValue::Bool(b),
        Value::Text(t) if t.is_empty() => DbValue::Null,
        _ => DbValue::Text(crate::display::cell_text(wb, sh, at)),
    }
}

fn same(a: &DbValue, b: &DbValue) -> bool {
    let num = |v: &DbValue| match v {
        DbValue::Number(n) | DbValue::Date(n) | DbValue::DateTime(n) | DbValue::Time(n) => Some(*n),
        DbValue::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        _ => None,
    };
    match (a, b) {
        (DbValue::Null, DbValue::Null) => true,
        (DbValue::Null, DbValue::Text(t)) | (DbValue::Text(t), DbValue::Null) => t.is_empty(),
        (DbValue::Text(x), DbValue::Text(y)) => x == y,
        _ => match (num(a), num(b)) {
            (Some(x), Some(y)) => (x - y).abs() <= 1e-9 * x.abs().max(y.abs()).max(1.0),
            _ => false,
        },
    }
}

/// A primary-key value as a map key.
fn key_part(v: &DbValue) -> String {
    match v {
        DbValue::Null => "\u{0}null".into(),
        DbValue::Text(t) => format!("t:{t}"),
        DbValue::Bool(b) => format!("n:{}", if *b { 1.0 } else { 0.0 }),
        DbValue::Number(n) | DbValue::Date(n) | DbValue::DateTime(n) | DbValue::Time(n) => format!("n:{n}"),
    }
}

fn quote(name: &str) -> String {
    format!("[{}]", name.replace(']', "]]"))
}

fn qualified(schema: &str, name: &str) -> String {
    if schema.is_empty() { quote(name) } else { format!("{}.{}", quote(schema), quote(name)) }
}

/// `INSERT INTO t (cols) VALUES (…),(…)` statements, as many rows each as the parameter limit
/// allows (and at most 1,000, SQL Server's row-constructor limit).
fn inserts(table: &str, columns: &[String], rows: Vec<Vec<DbValue>>) -> Vec<Statement> {
    let w = columns.len().max(1);
    let per = (MAX_PARAMS / w).clamp(1, 1000);
    let cols = columns.iter().map(|c| quote(c)).collect::<Vec<_>>().join(", ");
    let mut out = vec![];
    for chunk in rows.chunks(per) {
        let mut sql = format!("INSERT INTO {table} ({cols}) VALUES ");
        let mut params = Vec::with_capacity(chunk.len() * w);
        for (ri, row) in chunk.iter().enumerate() {
            if ri > 0 {
                sql.push_str(", ");
            }
            sql.push('(');
            for ci in 0..w {
                if ci > 0 {
                    sql.push_str(", ");
                }
                params.push(row.get(ci).cloned().unwrap_or(DbValue::Null));
                sql.push_str(&format!("@P{}", params.len()));
            }
            sql.push(')');
        }
        out.push(Statement::new(sql, params));
    }
    out
}

// ---------------------------------------------------------------- Run SQL

fn sql_execute(s: &mut Session, p: &Json) -> Result<Json> {
    const C: &str = "data.sqlExecute";
    let sql = str_param(p, "sql").map(str::trim).filter(|v| !v.is_empty()).ok_or_else(|| bad(C, "missing `sql`"))?.to_string();
    if sql.len() > 1_000_000 {
        return Err(bad(C, "`sql` is too long"));
    }
    let mut params: Vec<DbValue> = vec![];
    if let Some(list) = p.get("params").and_then(Json::as_array) {
        for v in list.iter().take(MAX_PARAMS) {
            params.push(match v {
                Json::Null => DbValue::Null,
                Json::Bool(b) => DbValue::Bool(*b),
                Json::Number(n) => DbValue::Number(n.as_f64().unwrap_or(0.0)),
                Json::String(t) => DbValue::Text(t.clone()),
                other => DbValue::Text(other.to_string()),
            });
        }
    }
    if let Some(cells) = p.get("cells").and_then(Json::as_array) {
        let d = s.doc()?;
        let sh = d.wb.active().ok_or(EngineError::NoDocument)?;
        for c in cells.iter().take(MAX_PARAMS) {
            let t = c.as_str().ok_or_else(|| bad(C, "`cells` holds cell references"))?;
            let at = CellRef::parse(split_sheet(t).1).ok_or_else(|| bad(C, format!("not a cell: `{t}`")))?;
            params.push(natural_value(&d.wb, sh, at));
        }
    }
    let (source, credential) = target(s, p, C)?;
    let provider = provider(s)?;
    let n = provider.execute(&source, &credential, &[Statement::new(sql, params)]).map_err(db_error)?;
    Ok(json!({"rowsAffected": n.iter().sum::<u64>()}))
}

// ---------------------------------------------------------------- Export

/// The range an export reads: `range`, else the table or the region around the active cell.
fn export_range(s: &Session, p: &Json) -> Result<(usize, RangeRef)> {
    let d = s.doc()?;
    let si = target_sheet(s, p)?;
    let sh = d.wb.sheet(si).ok_or(EngineError::NoDocument)?;
    if str_param(p, "range").is_some() {
        return Ok((si, target_range(s, p)?));
    }
    let at = d.selection.active;
    if let Some(t) = sh.table_at(at) {
        let r = match t.totals_range() {
            Some(_) => RangeRef::new(t.range.start, CellRef::new(t.range.end.row.saturating_sub(1), t.range.end.col)),
            None => t.range,
        };
        return Ok((si, r));
    }
    Ok((si, crate::selection::current_region(sh, at)))
}

fn sql_type_for(kind: Option<Kind>, values: &[DbValue]) -> String {
    match kind {
        Some(Kind::Bool) => "bit".into(),
        Some(Kind::Date) => "date".into(),
        Some(Kind::DateTime) => "datetime2(3)".into(),
        Some(Kind::Time) => "time(3)".into(),
        Some(Kind::Number) => {
            let ints = values.iter().all(|v| match v {
                DbValue::Number(n) => n.fract() == 0.0 && n.abs() < 9.0e15,
                _ => true,
            });
            if ints { "bigint".into() } else { "float".into() }
        }
        Some(Kind::Text) | None => {
            let longest = values.iter().map(|v| if let DbValue::Text(t) = v { t.chars().count() } else { 0 }).max().unwrap_or(0);
            match longest {
                0..=255 => "nvarchar(255)".into(),
                256..=4000 => "nvarchar(4000)".into(),
                _ => "nvarchar(max)".into(),
            }
        }
    }
}

fn export(s: &mut Session, p: &Json) -> Result<Json> {
    const C: &str = "data.exportToSqlServer";
    let table = str_param(p, "table").ok_or_else(|| bad(C, "missing `table` (for example dbo.Sales)"))?;
    let (schema, name) = datasource::parse_object_name(table).ok_or_else(|| bad(C, format!("not a table name: `{table}`")))?;
    let schema = if schema.is_empty() { "dbo".to_string() } else { schema };
    let mode = str_param(p, "mode").unwrap_or("create").to_ascii_lowercase();
    if !matches!(mode.as_str(), "create" | "append" | "replace") {
        return Err(bad(C, "`mode` is create, append or replace"));
    }
    let (si, r) = export_range(s, p)?;
    if r.height() < 2 {
        return Err(EngineError::Other("Select the data to export, with column names in its first row.".into()));
    }
    if r.count() > datasource::MAX_CELLS as u64 {
        return Err(EngineError::Other("The range is too large to export.".into()));
    }
    let d = s.doc()?;
    let sh = d.wb.sheet(si).ok_or(EngineError::NoDocument)?;
    let headers: Vec<String> = (r.start.col..=r.end.col).map(|c| sh.value(CellRef::new(r.start.row, c)).display().trim().to_string()).collect();
    if let Some(i) = headers.iter().position(|h| h.is_empty()) {
        return Err(EngineError::Other(format!("Column {} has no name in the first row.", i + 1)));
    }
    let mut seen = std::collections::HashSet::new();
    if let Some(h) = headers.iter().find(|h| !seen.insert(h.to_lowercase())) {
        return Err(EngineError::Other(format!("The column name `{h}` appears twice.")));
    }
    let data_rows: Vec<u32> =
        (r.start.row + 1..=r.end.row).filter(|row| (r.start.col..=r.end.col).any(|c| !sh.value(CellRef::new(*row, c)).is_empty())).collect();
    let object = qualified(&schema, &name);
    let (source, credential) = target(s, p, C)?;
    let provider = provider(s)?;
    // The kind of each column: the target's types when it exists, else what the cells hold.
    let target_cols: Option<Vec<DbColumn>> =
        if mode == "create" { None } else { Some(provider.columns(&source, &credential, &schema, &name).map_err(db_error)?) };
    let d = s.doc()?;
    let sh = d.wb.sheet(si).ok_or(EngineError::NoDocument)?;
    let mut columns: Vec<String> = vec![];
    let mut kinds: Vec<Kind> = vec![];
    let mut create = None;
    match &target_cols {
        Some(tc) => {
            for h in &headers {
                let c = tc.iter().find(|c| c.name.eq_ignore_ascii_case(h)).ok_or_else(|| {
                    EngineError::Other(format!(
                        "The table {object} has no column `{h}`. Columns: {}.",
                        tc.iter().map(|c| c.name.as_str()).collect::<Vec<_>>().join(", ")
                    ))
                })?;
                if c.generated {
                    return Err(EngineError::Other(format!("The server fills the column `{}` itself; leave it out of the export.", c.name)));
                }
                columns.push(c.name.clone());
                kinds.push(kind_of_sql(&c.sql_type));
            }
        }
        None => {
            let mut defs = vec![];
            for (ci, h) in headers.iter().enumerate() {
                let col = r.start.col + ci as u32;
                let values: Vec<DbValue> = data_rows.iter().map(|row| natural_value(&d.wb, sh, CellRef::new(*row, col))).collect();
                let mut kind: Option<Kind> = None;
                for v in &values {
                    match (kind, kind_of_value(v)) {
                        (_, None) => {}
                        (None, k) => kind = k,
                        (Some(a), Some(b)) if a == b => {}
                        (Some(Kind::Date), Some(Kind::DateTime)) | (Some(Kind::DateTime), Some(Kind::Date)) => kind = Some(Kind::DateTime),
                        _ => kind = Some(Kind::Text),
                    }
                }
                defs.push(format!("{} {} NULL", quote(h), sql_type_for(kind, &values)));
                columns.push(h.clone());
                kinds.push(kind.unwrap_or(Kind::Text));
            }
            create = Some(format!("CREATE TABLE {object} ({})", defs.join(", ")));
        }
    }
    let mut rows = Vec::with_capacity(data_rows.len());
    for row in &data_rows {
        let mut out = Vec::with_capacity(columns.len());
        for (ci, kind) in kinds.iter().enumerate() {
            out.push(cell_value(&d.wb, sh, CellRef::new(*row, r.start.col + ci as u32), *kind)?);
        }
        rows.push(out);
    }
    let n = rows.len();
    let mut statements = vec![];
    if let Some(sql) = &create {
        statements.push(Statement::new(sql.clone(), vec![]));
    }
    if mode == "replace" {
        statements.push(Statement::new(format!("DELETE FROM {object}"), vec![]));
    }
    statements.extend(inserts(&object, &columns, rows));
    let summary = json!({
        "table": object_label(&schema, &name),
        "mode": mode,
        "range": r.a1(),
        "rows": n,
        "columns": columns.iter().zip(&kinds).map(|(c, k)| json!({"name": c, "kind": format!("{k:?}").to_lowercase()})).collect::<Vec<_>>(),
        "create": create,
        "statements": statements.len(),
    });
    if bool_param(p, "dryRun").unwrap_or(false) {
        return Ok(summary);
    }
    provider.execute(&source, &credential, &statements).map_err(db_error)?;
    Ok(summary)
}

// ---------------------------------------------------------------- Save to Database

fn save(s: &mut Session, p: &Json) -> Result<Json> {
    const C: &str = "data.saveToDatabase";
    let ci = find_connection(s, p, C)?;
    let (conn, snap, si, ti, at) = {
        let d = s.doc()?;
        let conn = d.wb.connections.get(ci).cloned().ok_or_else(|| bad(C, "no such query"))?;
        if matches!(conn.command, QueryCommand::Sql { .. }) {
            return Err(EngineError::Other("Only a query that loads a whole table can be saved back; this one runs a SQL statement.".into()));
        }
        let snap = s.data.snapshots.iter().find(|x| x.doc == d.uid && x.connection == conn.id).cloned().ok_or_else(|| {
            EngineError::Other("Refresh the query before saving changes: GridCraft compares the table with the rows it loaded.".into())
        })?;
        let (si, ti) = d.wb.table(&conn.table).ok_or_else(|| EngineError::Other(format!("The table `{}` is gone.", conn.table)))?;
        let sh = d.wb.sheet(si).ok_or(EngineError::NoDocument)?;
        let t = sh.tables.get(ti).ok_or_else(|| EngineError::Other("The table is gone.".into()))?;
        // Sheet column of each loaded column.
        let mut at: Vec<u32> = vec![];
        for c in &snap.columns {
            let i = t
                .column_index(c)
                .ok_or_else(|| EngineError::Other(format!("The column `{c}` was renamed or removed. Refresh the query, then edit again.")))?;
            at.push(t.range.start.col + i as u32);
        }
        (conn, snap, si, ti, at)
    };
    let QueryCommand::Table { schema, name } = &conn.command else {
        return Err(EngineError::Other("Only a query that loads a whole table can be saved back.".into()));
    };
    let source = conn.source.clone();
    let credential = credential_for(s, &source, p)?;
    let provider = provider(s)?;
    let db_cols = provider.columns(&source, &credential, schema, name).map_err(db_error)?;
    let info = |c: &str| db_cols.iter().find(|x| x.name.eq_ignore_ascii_case(c));
    let pk: Vec<usize> = snap.columns.iter().enumerate().filter(|(_, c)| info(c).is_some_and(|x| x.primary_key)).map(|(i, _)| i).collect();
    if pk.is_empty() || pk.len() != db_cols.iter().filter(|c| c.primary_key).count() {
        return Err(EngineError::Other(format!(
            "The table {} has no primary key among the loaded columns, so its rows can't be matched. Changes can only be saved to a table with a primary key.",
            object_label(schema, name)
        )));
    }
    let writable: Vec<usize> = (0..snap.columns.len()).filter(|i| snap.columns.get(*i).and_then(|c| info(c)).is_some_and(|x| !x.generated)).collect();
    let kinds: Vec<Kind> = (0..snap.columns.len())
        .map(|i| {
            snap.rows
                .iter()
                .find_map(|r| r.get(i).and_then(kind_of_value))
                .unwrap_or_else(|| snap.columns.get(i).and_then(|c| info(c)).map_or(Kind::Text, |x| kind_of_sql(&x.sql_type)))
        })
        .collect();
    let key = |row: &[DbValue]| pk.iter().map(|i| key_part(row.get(*i).unwrap_or(&DbValue::Null))).collect::<Vec<_>>().join("\u{1}");
    let mut original: std::collections::HashMap<String, &Vec<DbValue>> = std::collections::HashMap::new();
    for r in &snap.rows {
        original.insert(key(r), r);
    }
    let d = s.doc()?;
    let sh = d.wb.sheet(si).ok_or(EngineError::NoDocument)?;
    let t = sh.tables.get(ti).ok_or_else(|| EngineError::Other("The table is gone.".into()))?;
    let object = qualified(schema, name);
    let mut seen = std::collections::HashSet::new();
    let (mut inserts_rows, mut statements) = (vec![], vec![]);
    let (mut updated, mut deleted) = (0usize, 0usize);
    let where_pk = |row: &[DbValue], first: usize| {
        let mut params = vec![];
        let clause = pk
            .iter()
            .filter_map(|i| {
                let col = snap.columns.get(*i)?;
                params.push(row.get(*i).cloned().unwrap_or(DbValue::Null));
                Some(format!("{} = @P{}", quote(col), first + params.len()))
            })
            .collect::<Vec<_>>()
            .join(" AND ");
        (clause, params)
    };
    if let Some(dr) = t.data_range() {
        for r in dr.start.row..=dr.end.row {
            let mut row = Vec::with_capacity(at.len());
            for (i, col) in at.iter().enumerate() {
                row.push(cell_value(&d.wb, sh, CellRef::new(r, *col), kinds.get(i).copied().unwrap_or(Kind::Text))?);
            }
            if row.iter().all(|v| matches!(v, DbValue::Null)) {
                continue;
            }
            let k = key(&row);
            match original.get(&k) {
                Some(old) => {
                    if !seen.insert(k) {
                        return Err(EngineError::Other(format!("Row {} repeats the primary key of another row.", r + 1)));
                    }
                    let changed: Vec<usize> = writable
                        .iter()
                        .copied()
                        .filter(|i| !pk.contains(i) && !same(row.get(*i).unwrap_or(&DbValue::Null), old.get(*i).unwrap_or(&DbValue::Null)))
                        .collect();
                    if changed.is_empty() {
                        continue;
                    }
                    let mut params: Vec<DbValue> = vec![];
                    let set = changed
                        .iter()
                        .filter_map(|i| {
                            params.push(row.get(*i).cloned().unwrap_or(DbValue::Null));
                            Some(format!("{} = @P{}", quote(snap.columns.get(*i)?), params.len()))
                        })
                        .collect::<Vec<_>>()
                        .join(", ");
                    let (clause, keys) = where_pk(old, params.len());
                    params.extend(keys);
                    statements.push(Statement { sql: format!("UPDATE {object} SET {set} WHERE {clause}"), params, expect_rows: Some(1) });
                    updated += 1;
                }
                None => inserts_rows.push(row),
            }
        }
    }
    for r in &snap.rows {
        if !seen.contains(&key(r)) {
            let (clause, params) = where_pk(r, 0);
            statements.push(Statement { sql: format!("DELETE FROM {object} WHERE {clause}"), params, expect_rows: Some(1) });
            deleted += 1;
        }
    }
    // New rows write every writable column (an identity key is left to the server).
    let inserted = inserts_rows.len();
    let cols: Vec<String> = writable.iter().filter_map(|i| snap.columns.get(*i).cloned()).collect();
    let rows: Vec<Vec<DbValue>> =
        inserts_rows.into_iter().map(|row| writable.iter().map(|i| row.get(*i).cloned().unwrap_or(DbValue::Null)).collect()).collect();
    statements.extend(inserts(&object, &cols, rows));
    let summary = json!({"table": object_label(schema, name), "inserted": inserted, "updated": updated, "deleted": deleted});
    if bool_param(p, "dryRun").unwrap_or(false) || statements.is_empty() {
        return Ok(summary);
    }
    provider.execute(&source, &credential, &statements).map_err(db_error)?;
    // Load what the server now holds (identity values, defaults, triggers).
    refresh_one(s, ci)?;
    Ok(summary)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::datasource::{DataProvider, DbError, DbObject, QueryResult};
    use std::sync::{Arc, Mutex};

    /// A one-table database held in memory; `execute` applies INSERT/UPDATE/DELETE of the shapes
    /// the commands send, so tests can check the result.
    struct MemDb {
        rows: Mutex<Vec<Vec<DbValue>>>,
        log: Mutex<Vec<Statement>>,
    }

    impl MemDb {
        fn new() -> Arc<MemDb> {
            Arc::new(MemDb {
                rows: Mutex::new(vec![
                    vec![DbValue::Number(1.0), DbValue::Text("Ada".into()), DbValue::Number(10.0)],
                    vec![DbValue::Number(2.0), DbValue::Text("Grace".into()), DbValue::Number(20.0)],
                    vec![DbValue::Number(3.0), DbValue::Text("Linus".into()), DbValue::Number(30.0)],
                ]),
                log: Mutex::new(vec![]),
            })
        }
    }

    impl DataProvider for MemDb {
        fn databases(&self, _: &DataSource, _: &Credential) -> std::result::Result<Vec<String>, DbError> {
            Ok(vec!["Shop".into()])
        }
        fn objects(&self, _: &DataSource, _: &Credential) -> std::result::Result<Vec<DbObject>, DbError> {
            Ok(vec![])
        }
        fn query(&self, _: &DataSource, _: &Credential, _: &QueryCommand, max: usize) -> std::result::Result<QueryResult, DbError> {
            let rows = self.rows.lock().map_err(|_| DbError::Other("poisoned".into()))?.clone();
            Ok(QueryResult {
                columns: vec!["Id".into(), "Name".into(), "Score".into()],
                truncated: rows.len() > max,
                rows: rows.into_iter().take(max).collect(),
            })
        }
        fn columns(&self, _: &DataSource, _: &Credential, _: &str, name: &str) -> std::result::Result<Vec<DbColumn>, DbError> {
            let col = |n: &str, t: &str, pk: bool, generated: bool| DbColumn {
                name: n.into(),
                sql_type: t.into(),
                nullable: !pk,
                primary_key: pk,
                generated,
            };
            if name == "NoKey" {
                return Ok(vec![col("Id", "int", false, false), col("Name", "nvarchar", false, false), col("Score", "float", false, false)]);
            }
            Ok(vec![col("Id", "int", true, true), col("Name", "nvarchar", false, false), col("Score", "float", false, false)])
        }
        fn execute(&self, _: &DataSource, _: &Credential, statements: &[Statement]) -> std::result::Result<Vec<u64>, DbError> {
            let mut rows = self.rows.lock().map_err(|_| DbError::Other("poisoned".into()))?;
            let mut out = vec![];
            for st in statements {
                self.log.lock().map_err(|_| DbError::Other("poisoned".into()))?.push(st.clone());
                let n = if st.sql.starts_with("UPDATE") {
                    // UPDATE … SET [Name] = @P1 … WHERE [Id] = @Pn: the key is the last param.
                    let id = st.params.last().cloned();
                    let hit = rows.iter_mut().find(|r| r.first() == id.as_ref());
                    match hit {
                        Some(r) => {
                            for (i, col) in ["[Name]", "[Score]"].iter().enumerate() {
                                if let Some(pos) = st.sql.find(&format!("{col} = @P")) {
                                    let num: usize = st.sql[pos + col.len() + " = @P".len()..]
                                        .chars()
                                        .take_while(|c| c.is_ascii_digit())
                                        .collect::<String>()
                                        .parse()
                                        .unwrap_or(0);
                                    if let Some(v) = num.checked_sub(1).and_then(|i| st.params.get(i)) {
                                        r[i + 1] = v.clone();
                                    }
                                }
                            }
                            1
                        }
                        None => 0,
                    }
                } else if st.sql.starts_with("DELETE") {
                    let id = st.params.first().cloned();
                    let before = rows.len();
                    rows.retain(|r| r.first() != id.as_ref());
                    (before - rows.len()) as u64
                } else if st.sql.starts_with("INSERT") {
                    let next =
                        rows.iter().filter_map(|r| if let Some(DbValue::Number(n)) = r.first() { Some(*n) } else { None }).fold(0.0, f64::max) + 1.0;
                    rows.push(vec![DbValue::Number(next), st.params[0].clone(), st.params[1].clone()]);
                    1
                } else {
                    0
                };
                if let Some(want) = st.expect_rows
                    && n != want
                {
                    return Err(DbError::Other("conflict".into()));
                }
                out.push(n);
            }
            Ok(out)
        }
    }

    fn session(db: &Arc<MemDb>) -> Session {
        let mut s = Session::new();
        s.execute("file.new", json!({})).unwrap();
        s.set_data_provider(db.clone());
        s.set_credential("db", Credential { user: "sa".into(), password: "pw".into() });
        s
    }

    #[test]
    fn save_sends_only_what_changed() {
        let db = MemDb::new();
        let mut s = session(&db);
        s.execute("data.fromSqlServer", json!({"server": "db", "database": "Shop", "table": "dbo.People"})).unwrap();
        // Edit Grace's score, delete Linus's row, add a new person (no Id: the server assigns it).
        s.execute("cell.set", json!({"cell": "C3", "input": "25"})).unwrap();
        s.execute("edit.clearContents", json!({"range": "A4:C4"})).unwrap();
        s.execute("cell.set", json!({"cell": "B5", "input": "Margaret"})).unwrap();
        s.execute("cell.set", json!({"cell": "C5", "input": "40"})).unwrap();
        s.execute("table.resize", json!({"table": "dbo_People", "range": "A1:C5"})).unwrap();
        s.execute("selection.set", json!({"range": "A2"})).unwrap();
        let dry = s.execute("data.saveToDatabase", json!({"dryRun": true})).unwrap();
        assert_eq!((dry["inserted"].as_u64(), dry["updated"].as_u64(), dry["deleted"].as_u64()), (Some(1), Some(1), Some(1)), "{dry}");
        assert!(db.log.lock().unwrap().is_empty(), "a dry run writes nothing");
        s.execute("data.saveToDatabase", json!({})).unwrap();
        let rows = db.rows.lock().unwrap().clone();
        assert_eq!(rows.len(), 3);
        assert_eq!(rows[1][2], DbValue::Number(25.0));
        assert_eq!(rows[2][1], DbValue::Text("Margaret".into()));
        // The update names only the changed column; deletes and updates must hit one row.
        let log = db.log.lock().unwrap().clone();
        let upd = log.iter().find(|x| x.sql.starts_with("UPDATE")).unwrap();
        assert_eq!(upd.sql, "UPDATE [dbo].[People] SET [Score] = @P1 WHERE [Id] = @P2");
        assert_eq!(upd.expect_rows, Some(1));
        let ins = log.iter().find(|x| x.sql.starts_with("INSERT")).unwrap();
        assert_eq!(ins.sql, "INSERT INTO [dbo].[People] ([Name], [Score]) VALUES (@P1, @P2)");
        // Saving reloads the table: the new row has its Id now, and saving again sends nothing.
        let sh = s.doc().unwrap().wb.active().unwrap();
        // (The fake hands out max + 1, so the freed Id 3 comes back.)
        assert_eq!(sh.value(CellRef::parse("A4").unwrap()), Value::Number(3.0));
        assert_eq!(sh.value(CellRef::parse("B4").unwrap()).display(), "Margaret");
        let n = db.log.lock().unwrap().len();
        let again = s.execute("data.saveToDatabase", json!({"table": "dbo_People"})).unwrap();
        assert_eq!(again["updated"], 0);
        assert_eq!(db.log.lock().unwrap().len(), n);
        // Not journaled: replaying a script never writes again.
        assert!(!s.journal.iter().any(|(id, _)| id == "data.saveToDatabase"));
    }

    #[test]
    fn save_needs_a_snapshot_and_a_key() {
        let db = MemDb::new();
        let mut s = session(&db);
        s.execute("data.fromSqlServer", json!({"server": "db", "database": "Shop", "table": "dbo.NoKey"})).unwrap();
        let e = s.execute("data.saveToDatabase", json!({})).unwrap_err().to_string();
        assert!(e.contains("primary key"), "{e}");
        // After reopening (no snapshot), a refresh is needed first.
        s.data.snapshots.clear();
        let e = s.execute("data.saveToDatabase", json!({})).unwrap_err().to_string();
        assert!(e.contains("Refresh the query"), "{e}");
        let mut q = session(&db);
        q.execute("data.fromSqlServer", json!({"server": "db", "database": "Shop", "sql": "SELECT 1"})).unwrap();
        assert!(q.execute("data.saveToDatabase", json!({})).unwrap_err().to_string().contains("SQL statement"));
    }

    #[test]
    fn export_creates_a_table_and_inserts() {
        let db = MemDb::new();
        let mut s = session(&db);
        s.execute("range.setValues", json!({"range": "A1", "values": [["Region", "Units", "Price", "Since", "Active"], ["North", 3, 1.5, 45000, true], ["South", 4, 2.25, 45001, false]]})).unwrap();
        s.execute("home.numberFormat", json!({"range": "D2:D3", "format": "m/d/yyyy"})).unwrap();
        s.execute("selection.set", json!({"range": "B2"})).unwrap();
        let r = s.execute("data.exportToSqlServer", json!({"server": "db", "database": "Shop", "table": "Sales", "dryRun": true})).unwrap();
        assert_eq!(r["rows"], 2);
        assert_eq!(
            r["create"],
            "CREATE TABLE [dbo].[Sales] ([Region] nvarchar(255) NULL, [Units] bigint NULL, [Price] float NULL, [Since] date NULL, [Active] bit NULL)"
        );
        assert!(db.log.lock().unwrap().is_empty());
        s.execute("data.exportToSqlServer", json!({"server": "db", "database": "Shop", "table": "Sales"})).unwrap();
        let log = db.log.lock().unwrap().clone();
        assert!(log[0].sql.starts_with("CREATE TABLE"));
        assert_eq!(
            log[1].sql,
            "INSERT INTO [dbo].[Sales] ([Region], [Units], [Price], [Since], [Active]) VALUES (@P1, @P2, @P3, @P4, @P5), (@P6, @P7, @P8, @P9, @P10)"
        );
        assert_eq!(log[1].params[3], DbValue::Date(45000.0));
        // Appending maps names to the table's columns; an unknown one is refused.
        let e =
            s.execute("data.exportToSqlServer", json!({"server": "db", "database": "Shop", "table": "dbo.People", "mode": "append"})).unwrap_err();
        assert!(e.to_string().contains("no column `Region`"), "{e}");
    }

    #[test]
    fn statements_split_at_the_parameter_limit() {
        let rows: Vec<Vec<DbValue>> = (0..2500).map(|i| vec![DbValue::Number(i as f64)]).collect();
        let st = inserts("[t]", &["a".into()], rows);
        assert_eq!(st.len(), 3);
        assert!(st.iter().all(|x| x.params.len() <= 1000));
        let wide: Vec<Vec<DbValue>> = (0..10).map(|_| vec![DbValue::Null; 300]).collect();
        assert!(inserts("[t]", &vec!["c".to_string(); 300], wide).iter().all(|x| x.params.len() <= MAX_PARAMS));
    }

    #[test]
    fn run_sql_with_cell_params() {
        let db = MemDb::new();
        let mut s = session(&db);
        s.execute("cell.set", json!({"cell": "B2", "input": "Grace"})).unwrap();
        let r = s.execute(
            "data.sqlExecute",
            json!({"server": "db", "database": "Shop", "sql": "DELETE FROM dbo.People WHERE Name = @P1", "cells": ["B2"]}),
        );
        assert!(r.is_ok(), "{r:?}");
        let log = db.log.lock().unwrap().clone();
        assert_eq!(log[0].params, vec![DbValue::Text("Grace".into())]);
        assert!(s.execute("data.sqlExecute", json!({"server": "db"})).is_err());
    }

    #[test]
    fn values_compare_loosely_where_sheets_do() {
        assert!(same(&DbValue::Null, &DbValue::Text(String::new())));
        assert!(same(&DbValue::Number(0.1 + 0.2), &DbValue::Number(0.3)));
        assert!(same(&DbValue::Date(45000.0), &DbValue::Number(45000.0)));
        assert!(!same(&DbValue::Text("a".into()), &DbValue::Text("A".into())));
    }
}

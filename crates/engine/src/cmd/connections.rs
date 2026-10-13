//! Get Data › From Database › From SQL Server Database, the Navigator, Refresh and Queries &
//! Connections.
//!
//! The engine asks the session's [`DataProvider`](crate::datasource::DataProvider) for rows and
//! loads them into a worksheet table bound to a [`Connection`]. Refresh fetches them again. Logins
//! are kept in the session, never in the workbook or the journal (`password` params are secret).

use std::sync::Arc;

use gridcraft_core::{CellRef, RangeRef, Value};
use gridcraft_model::*;
use serde_json::{Value as Json, json};

use super::*;
use crate::datasource::{self, Credential, DbValue, QueryResult};

pub fn specs() -> Vec<CommandSpec> {
    vec![
        cmd!(
            "data.fromSqlServer",
            "From SQL Server Database",
            ["Data", "Get & Transform Data", "Get Data", "From Database"],
            None,
            "{server: \"host\" | \"host,port\" | \"host\\\\instance\", database?, table?: \"dbo.Customers\" | tables?: [...] | sql?: \"SELECT …\" (needs database), user?, password? (kept for the session, never saved), encrypt?: true, trustServerCertificate?: false, name?: query name} → loads each into a new sheet as a table bound to a connection → {loaded: [{sheet, table, connection, rows, cols, truncated}]}",
            has_doc,
            from_sql_server
        ),
        cmd!(
            query "data.sqlServerNavigator",
            "Navigator",
            [],
            None,
            "{server, database?, user?, password?, encrypt?, trustServerCertificate?} → {databases: [...]} without a database, else {objects: [{name: \"dbo.Customers\", schema, table, view}]}",
            always,
            navigator
        ),
        cmd!(
            query "data.sqlServerPreview",
            "Preview",
            [],
            None,
            "{server, database?, table | sql, rows?: 100 (max 1000), user?, password?, encrypt?, trustServerCertificate?} → {columns, rows (display text), truncated}",
            always,
            preview
        ),
        cmd!(
            "data.refresh",
            "Refresh",
            ["Data", "Queries & Connections", "Refresh All"],
            None,
            "{table? | connection?} (default: the table at the active cell) fetches the rows again → {table, rows, cols}",
            has_doc,
            refresh
        ),
        cmd!(
            query "data.dataSourceSettings",
            "Data Source Settings",
            ["Data", "Get & Transform Data", "Get Data"],
            None,
            "{} → [{server, user}] logins kept this session; {server, user, password} sets one; {server, clear: true} forgets it",
            always,
            data_source_settings
        ),
        cmd!(
            "data.connectionProperties",
            "Connection Properties",
            ["Data", "Queries & Connections", "Properties"],
            None,
            "{connection, name?, description?, refreshOnOpen?, refreshMinutes?: 0..32767, sql?: new SQL statement, table?: \"dbo.X\" new source table, user?} → the connection",
            has_doc,
            properties
        ),
        cmd!(
            "data.deleteConnection",
            "Delete Query",
            ["Data", "Queries & Connections"],
            None,
            "{connection} removes the query; its table keeps the rows last loaded",
            has_doc,
            delete_connection
        ),
    ]
}

/// The provider, or the error that says this host has none.
pub(crate) fn provider(s: &Session) -> Result<Arc<dyn datasource::DataProvider>> {
    s.data_provider().ok_or_else(|| EngineError::Other(datasource::NOT_AVAILABLE.into()))
}

pub(crate) fn db_error(e: datasource::DbError) -> EngineError {
    EngineError::Other(e.to_string())
}

/// The source a call names (`server`, `database`, `encrypt`, `trustServerCertificate`).
pub(crate) fn source_param(cmd: &str, p: &Json) -> Result<DataSource> {
    let server = str_param(p, "server").map(str::trim).filter(|v| !v.is_empty()).ok_or_else(|| bad(cmd, "missing `server`"))?;
    if server.len() > 256 {
        return Err(bad(cmd, "`server` is too long"));
    }
    let database = str_param(p, "database").map(str::trim).unwrap_or("");
    if database.len() > 128 {
        return Err(bad(cmd, "`database` is too long"));
    }
    Ok(DataSource::SqlServer {
        server: server.to_string(),
        database: database.to_string(),
        encrypt: bool_param(p, "encrypt").unwrap_or(true),
        trust_server_certificate: bool_param(p, "trustServerCertificate").unwrap_or(false),
    })
}

/// Stores `user`/`password` when given, then returns the session's login for the source's server.
pub(crate) fn credential_for(s: &mut Session, source: &DataSource, p: &Json) -> Result<Credential> {
    if let (Some(user), Some(password)) = (str_param(p, "user"), str_param(p, "password")) {
        s.set_credential(source.server(), Credential { user: user.to_string(), password: password.to_string() });
    }
    s.credential(source.server())
        .cloned()
        .ok_or_else(|| EngineError::Other(format!("{} No login is stored for server `{}`.", datasource::CREDENTIALS_REQUIRED, source.server())))
}

/// The command a call names: `sql`, or `table` (`dbo.Customers`).
fn command_param(cmd: &str, p: &Json, key: &str) -> Result<Option<QueryCommand>> {
    if let Some(sql) = str_param(p, "sql").map(str::trim).filter(|v| !v.is_empty()) {
        if sql.len() > 1_000_000 {
            return Err(bad(cmd, "`sql` is too long"));
        }
        return Ok(Some(QueryCommand::Sql { text: sql.to_string() }));
    }
    match str_param(p, key) {
        Some(t) => {
            let (schema, name) = datasource::parse_object_name(t).ok_or_else(|| bad(cmd, format!("not a table name: `{t}`")))?;
            Ok(Some(QueryCommand::Table { schema, name }))
        }
        None => Ok(None),
    }
}

// ---------------------------------------------------------------- values

/// Number formats of loaded dates and times (the built-in date, date-time and time formats).
const DATE_FORMAT: &str = "m/d/yyyy";
const DATETIME_FORMAT: &str = "m/d/yyyy h:mm";
const TIME_FORMAT: &str = "h:mm:ss";

fn cell_value(v: &DbValue) -> Option<(Value, Option<&'static str>)> {
    Some(match v {
        DbValue::Null => return None,
        DbValue::Bool(b) => (Value::Bool(*b), None),
        DbValue::Number(n) => (Value::number(*n), None),
        DbValue::Text(t) => (Value::text(clip(t)), None),
        DbValue::Date(n) => (Value::number(*n), Some(DATE_FORMAT)),
        DbValue::DateTime(n) => (Value::number(*n), Some(DATETIME_FORMAT)),
        DbValue::Time(n) => (Value::number(*n), Some(TIME_FORMAT)),
    })
}

/// Text cut to what a cell holds.
fn clip(t: &str) -> &str {
    match t.char_indices().nth(datasource::MAX_TEXT) {
        Some((i, _)) => t.get(..i).unwrap_or(t),
        None => t,
    }
}

/// A value as the preview shows it.
fn display(v: &DbValue) -> Json {
    match cell_value(v) {
        None => Json::Null,
        Some((Value::Bool(b), _)) => json!(if b { "TRUE" } else { "FALSE" }),
        Some((value, Some(code))) => {
            json!(gridcraft_numfmt::text_function(&value, code, gridcraft_core::date::DateSystem::D1900).unwrap_or_else(|_| value.display()))
        }
        Some((value, None)) => json!(value.display()),
    }
}

/// Column names made unique (case-insensitively) and never blank, as table headers must be.
pub(crate) fn header_names(columns: &[String]) -> Vec<String> {
    let mut taken = std::collections::HashSet::new();
    columns
        .iter()
        .enumerate()
        .map(|(i, c)| {
            let base = if c.trim().is_empty() { format!("Column{}", i + 1) } else { clip(c).to_string() };
            let mut n = base.clone();
            let mut k = 2;
            while !taken.insert(n.to_lowercase()) {
                n = format!("{base}{k}");
                k += 1;
            }
            n
        })
        .collect()
}

/// The result cut to the sheet: at most [`datasource::MAX_ROWS`] rows, the sheet's columns and
/// [`datasource::MAX_CELLS`] cells.
pub(crate) fn fit(mut r: QueryResult) -> QueryResult {
    r.columns.truncate(gridcraft_core::MAX_COLS as usize);
    let w = r.columns.len().max(1);
    let max_rows = datasource::MAX_ROWS.min(datasource::MAX_CELLS / w);
    if r.rows.len() > max_rows {
        r.rows.truncate(max_rows);
        r.truncated = true;
    }
    r
}

/// Writes the header and rows at `at` (cells left blank stay empty) and returns the range.
fn write_rows(cx: &mut Ctx, si: usize, at: CellRef, headers: &[String], r: &QueryResult) -> Result<RangeRef> {
    let w = headers.len().max(1) as u32;
    let h = r.rows.len().max(1) as u32;
    let end_row = at.row.checked_add(h).filter(|e| *e < gridcraft_core::MAX_ROWS);
    let end_col = at.col.checked_add(w - 1).filter(|e| *e < gridcraft_core::MAX_COLS);
    let (Some(end_row), Some(end_col)) = (end_row, end_col) else {
        return Err(EngineError::Other("The rows don't fit on the worksheet.".into()));
    };
    let base = StyleId::default();
    let mut formats: Vec<(&'static str, StyleId)> = vec![];
    for (i, name) in headers.iter().enumerate() {
        let c = CellRef::new(at.row, at.col + i as u32);
        cx.sheet_mut(si)?.set_value(c, Value::text(name.as_str()));
        cx.touch(si, c);
    }
    for (ri, row) in r.rows.iter().enumerate() {
        for (ci, v) in row.iter().take(headers.len()).enumerate() {
            let Some((value, code)) = cell_value(v) else { continue };
            let c = CellRef::new(at.row + 1 + ri as u32, at.col + ci as u32);
            let style = match code {
                None => base,
                Some(code) => match formats.iter().find(|(k, _)| *k == code) {
                    Some((_, id)) => *id,
                    None => {
                        let mut st = cx.wb.styles.get(base).clone();
                        st.num_fmt = NumFmt::new(code);
                        let id = cx.wb.styles.intern(st);
                        formats.push((code, id));
                        id
                    }
                },
            };
            cx.sheet_mut(si)?.set_cell(c, Cell { value, formula: None, style });
        }
    }
    Ok(RangeRef::new(at, CellRef::new(end_row, end_col)))
}

/// A table name from a query name: invalid characters become `_`, unique in the workbook.
fn table_name_for(wb: &Workbook, query: &str) -> String {
    let mut base: String = query.chars().map(|c| if c.is_alphanumeric() || c == '_' || c == '.' { c } else { '_' }).take(200).collect();
    if !base.starts_with(|c: char| c.is_alphabetic() || c == '_') || CellRef::parse(&base).is_some() {
        base = format!("_{base}");
    }
    let free = |n: &str| wb.table(n).is_none() && !wb.names.iter().any(|d| d.name.eq_ignore_ascii_case(n));
    if free(&base) {
        return base;
    }
    (2..).map(|k| format!("{base}_{k}")).find(|n| free(n)).unwrap_or(base)
}

/// A sheet name from a query name: unique, valid, else the next `SheetN`.
fn sheet_name_for(wb: &Workbook, query: &str) -> String {
    let base: String = query.chars().filter(|c| !matches!(c, ':' | '\\' | '/' | '?' | '*' | '[' | ']')).take(31).collect();
    let base = base.trim_matches('\'').trim().to_string();
    if wb.check_sheet_name(&base, None).is_ok() {
        return base;
    }
    for k in 2..1000 {
        let suffix = format!(" ({k})");
        let n: String = base.chars().take(31usize.saturating_sub(suffix.chars().count())).collect::<String>() + &suffix;
        if wb.check_sheet_name(&n, None).is_ok() {
            return n;
        }
    }
    wb.next_sheet_name()
}

/// A query name unique among the workbook's connections.
fn query_name_for(wb: &Workbook, base: &str) -> String {
    let free = |n: &str| !wb.connections.iter().any(|c| c.name.eq_ignore_ascii_case(n));
    if free(base) {
        return base.to_string();
    }
    (2..).map(|k| format!("{base} ({k})")).find(|n| free(n)).unwrap_or_else(|| base.to_string())
}

fn next_query_number(wb: &Workbook) -> String {
    (1..).map(|n| format!("Query{n}")).find(|n| !wb.connections.iter().any(|c| c.name.eq_ignore_ascii_case(n))).unwrap_or_else(|| "Query".into())
}

/// Runs nested commands as part of the calling command: their undo steps and journal entries are
/// dropped, so the caller's one undo step covers them.
fn quietly<R>(s: &mut Session, f: impl FnOnce(&mut Session) -> R) -> R {
    let undo = s.active().map(|d| d.undo.len());
    let journal = s.journal.len();
    let r = f(s);
    if let (Some(n), Some(d)) = (undo, s.active_mut()) {
        d.undo.truncate(n);
    }
    s.journal.truncate(journal);
    r
}

// ---------------------------------------------------------------- load

fn from_sql_server(s: &mut Session, p: &Json) -> Result<Json> {
    const C: &str = "data.fromSqlServer";
    if str_param(p, "server").is_none_or(|v| v.trim().is_empty()) {
        s.ui_requests.push(crate::UiRequest::Dialog("sqlServer".into(), json!({})));
        return ok();
    }
    let source = source_param(C, p)?;
    let mut commands = vec![];
    if let Some(c) = command_param(C, p, "table")? {
        commands.push(c);
    }
    if let Some(list) = p.get("tables").and_then(Json::as_array) {
        for t in list.iter().take(100) {
            let t = t.as_str().ok_or_else(|| bad(C, "`tables` holds table names"))?;
            let (schema, name) = datasource::parse_object_name(t).ok_or_else(|| bad(C, format!("not a table name: `{t}`")))?;
            commands.push(QueryCommand::Table { schema, name });
        }
    }
    if commands.is_empty() {
        return Err(bad(C, "give `table` (see data.sqlServerNavigator) or `sql`"));
    }
    if commands.iter().any(|c| matches!(c, QueryCommand::Sql { .. })) && source.database().is_empty() {
        return Err(bad(C, "a SQL statement needs a `database`"));
    }
    if s.doc()?.wb.protected_structure {
        return Err(EngineError::Other("The workbook is protected and cannot be changed.".into()));
    }
    let provider = provider(s)?;
    let credential = credential_for(s, &source, p)?;
    let mut loaded = vec![];
    let mut truncated_any = false;
    for (i, command) in commands.iter().enumerate() {
        let result = fit(provider.query(&source, &credential, command, datasource::MAX_ROWS + 1).map_err(db_error)?);
        truncated_any |= result.truncated;
        let given = str_param(p, "name").map(str::trim).filter(|n| !n.is_empty() && i == 0).map(str::to_string);
        let r = load_new(s, &source, &credential.user, command.clone(), given, &result)?;
        loaded.push(r);
    }
    if truncated_any {
        s.ui_requests.push(crate::UiRequest::Message(format!(
            "The query returned more rows than a worksheet holds. The first {} rows were loaded.",
            datasource::MAX_ROWS
        )));
    }
    Ok(json!({"loaded": loaded}))
}

/// Loads a result into a new sheet as a table and adds its connection.
fn load_new(s: &mut Session, source: &DataSource, user: &str, command: QueryCommand, name: Option<String>, r: &QueryResult) -> Result<Json> {
    let d = s.doc()?;
    let query = query_name_for(&d.wb, &name.or_else(|| command.object_name()).unwrap_or_else(|| next_query_number(&d.wb)));
    let sheet = sheet_name_for(&d.wb, &query);
    let table = table_name_for(&d.wb, &query);
    let headers = header_names(&r.columns);
    if headers.is_empty() {
        return Err(EngineError::Other("The query returned no columns.".into()));
    }
    let table_id = d.wb.next_object_id();
    let conn_id = d.wb.connections.iter().map(|c| c.id).max().unwrap_or(0).saturating_add(1);
    let now = gridcraft_calc::now_serial();
    let rows = r.rows.len();
    let (si, range) = edit(s, |cx| {
        let at = (cx.wb.active_sheet + 1).min(cx.wb.sheets.len());
        cx.wb.sheets.insert(at, Arc::new(Sheet::new(sheet.clone())));
        for n in cx.wb.names.iter_mut() {
            if let Some(sc) = n.scope.as_mut()
                && *sc >= at
            {
                *sc += 1;
            }
        }
        cx.wb.active_sheet = at;
        let range = write_rows(cx, at, CellRef::new(0, 0), &headers, r)?;
        cx.sheet_mut(at)?.tables.push(new_table(table_id, &table, range, &headers));
        cx.wb.connections.push(Connection {
            id: conn_id,
            name: query.clone(),
            description: String::new(),
            source: source.clone(),
            user: user.to_string(),
            command,
            table: table.clone(),
            refresh_on_open: false,
            refresh_minutes: 0,
            refreshed: Some(now),
            rows: rows as u64,
        });
        *cx.sel = crate::Selection::at(CellRef::new(0, 0));
        cx.structural = true;
        Ok((at, range))
    })?;
    if let Some(d) = s.active_mut() {
        d.sheet_selections.insert(si.min(d.sheet_selections.len()), crate::Selection::default());
    }
    super::dbwrite::remember(s, conn_id, &r.columns, &r.rows);
    let cols = range.width();
    quietly(s, |s| {
        let _ = s.execute("home.autofitColumnWidth", json!({"cols": RangeRef::cols(0, cols - 1).a1()}));
    });
    Ok(json!({"sheet": sheet, "table": table, "connection": query, "rows": rows, "cols": cols, "truncated": r.truncated}))
}

fn new_table(id: u32, name: &str, range: RangeRef, headers: &[String]) -> Table {
    Table {
        id,
        name: name.to_string(),
        range,
        header_row: true,
        totals_row: false,
        columns: headers.iter().map(|n| column(n)).collect(),
        style: "TableStyleMedium2".into(),
        banded_rows: true,
        banded_cols: false,
        first_col: false,
        last_col: false,
        filter_button: true,
        filter: vec![],
        table_type: None,
        connection_id: None,
    }
}

fn column(name: &str) -> TableColumn {
    TableColumn { name: name.to_string(), totals: TotalsFn::None, totals_label: None, formula: None, unique_name: None, xml: None }
}

// ---------------------------------------------------------------- navigator, preview

fn navigator(s: &mut Session, p: &Json) -> Result<Json> {
    const C: &str = "data.sqlServerNavigator";
    let source = source_param(C, p)?;
    let provider = provider(s)?;
    let credential = credential_for(s, &source, p)?;
    if source.database().is_empty() {
        let dbs = provider.databases(&source, &credential).map_err(db_error)?;
        return Ok(json!({"databases": dbs}));
    }
    let objects = provider.objects(&source, &credential).map_err(db_error)?;
    Ok(json!({
        "objects": objects
            .iter()
            .map(|o| json!({"name": object_label(&o.schema, &o.name), "schema": o.schema, "table": o.name, "view": o.view}))
            .collect::<Vec<_>>()
    }))
}

/// `dbo.Customers`, bracketing parts that need it (`[sales].[Order Details]`).
pub fn object_label(schema: &str, name: &str) -> String {
    let part = |x: &str| {
        if !x.is_empty() && x.chars().all(|c| c.is_alphanumeric() || c == '_') { x.to_string() } else { format!("[{}]", x.replace(']', "]]")) }
    };
    if schema.is_empty() { part(name) } else { format!("{}.{}", part(schema), part(name)) }
}

fn preview(s: &mut Session, p: &Json) -> Result<Json> {
    const C: &str = "data.sqlServerPreview";
    let source = source_param(C, p)?;
    let command = command_param(C, p, "table")?.ok_or_else(|| bad(C, "give `table` or `sql`"))?;
    if matches!(command, QueryCommand::Sql { .. }) && source.database().is_empty() {
        return Err(bad(C, "a SQL statement needs a `database`"));
    }
    let n = f64_param(p, "rows").filter(|n| n.is_finite()).unwrap_or(100.0).clamp(1.0, 1000.0) as usize;
    let provider = provider(s)?;
    let credential = credential_for(s, &source, p)?;
    let r = fit(provider.query(&source, &credential, &command, n).map_err(db_error)?);
    Ok(json!({
        "columns": r.columns,
        "rows": r.rows.iter().map(|row| row.iter().map(display).collect::<Vec<_>>()).collect::<Vec<_>>(),
        "truncated": r.truncated,
    }))
}

// ---------------------------------------------------------------- refresh

/// The connection a call names (`connection`, or `table`), else the one whose table holds the
/// active cell.
pub(crate) fn find_connection(s: &Session, p: &Json, cmd: &str) -> Result<usize> {
    let d = s.doc()?;
    let conns = &d.wb.connections;
    if let Some(n) = str_param(p, "connection") {
        return conns.iter().position(|c| c.name.eq_ignore_ascii_case(n)).ok_or_else(|| bad(cmd, format!("no query named `{n}`")));
    }
    let table = match str_param(p, "table") {
        Some(t) => t.to_string(),
        None => {
            d.wb.active()
                .and_then(|sh| sh.table_at(d.selection.active))
                .map(|t| t.name.clone())
                .ok_or_else(|| bad(cmd, "select a cell in a table that was loaded from a database"))?
        }
    };
    conns.iter().position(|c| c.table.eq_ignore_ascii_case(&table)).ok_or_else(|| bad(cmd, format!("table `{table}` isn't loaded from a database")))
}

fn refresh(s: &mut Session, p: &Json) -> Result<Json> {
    let ci = find_connection(s, p, "data.refresh")?;
    refresh_one(s, ci)
}

/// Fetches connection `ci`'s rows again and rewrites its table.
pub(crate) fn refresh_one(s: &mut Session, ci: usize) -> Result<Json> {
    let conn = s.doc()?.wb.connections.get(ci).cloned().ok_or_else(|| bad("data.refresh", "no such query"))?;
    let (si, ti) =
        s.doc()?.wb.table(&conn.table).ok_or_else(|| EngineError::Other(format!("The table `{}` of query `{}` is gone.", conn.table, conn.name)))?;
    let provider = provider(s)?;
    let credential = s.credential(conn.source.server()).cloned().ok_or_else(|| {
        EngineError::Other(format!("{} No login is stored for server `{}`.", datasource::CREDENTIALS_REQUIRED, conn.source.server()))
    })?;
    let r = fit(provider.query(&conn.source, &credential, &conn.command, datasource::MAX_ROWS + 1).map_err(db_error)?);
    let headers = header_names(&r.columns);
    if headers.is_empty() {
        return Err(EngineError::Other("The query returned no columns.".into()));
    }
    let now = gridcraft_calc::now_serial();
    let rows = r.rows.len();
    let range = edit(s, |cx| {
        let sh = cx.wb.sheet(si).ok_or(EngineError::NoDocument)?;
        let t = sh.tables.get(ti).cloned().ok_or_else(|| EngineError::Other("The table is gone.".into()))?;
        let old = t.range;
        let start = old.start;
        let h = rows.max(1) as u32 + t.totals_row as u32;
        let new_end = CellRef::new(
            start.row.saturating_add(h).min(gridcraft_core::MAX_ROWS - 1),
            start.col.saturating_add(headers.len() as u32 - 1).min(gridcraft_core::MAX_COLS - 1),
        );
        let new = RangeRef::new(start, new_end);
        // Rows that would land on cells outside the table: refuse rather than overwrite them.
        if sh.tables.iter().enumerate().any(|(i, o)| i != ti && o.range.intersects(&new)) {
            return Err(EngineError::Other("The refreshed rows would overlap another table.".into()));
        }
        let blocked = sh.cells.iter_range(new).any(|(c, x)| !old.contains(c) && (x.formula.is_some() || !x.value.is_empty()));
        if blocked {
            return Err(EngineError::Other(
                "The refreshed rows would overwrite cells next to the table. Move those cells, then refresh again.".into(),
            ));
        }
        // The totals row keeps its cells, moved to the new last row.
        let totals: Vec<(u32, Cell)> = match t.totals_range() {
            Some(tr) => sh.cells.iter_range(tr).map(|(c, x)| (c.col - start.col, x.clone())).collect(),
            None => vec![],
        };
        let filtered: Vec<u32> =
            t.data_range().map(|dr| (dr.start.row..=dr.end.row).filter(|r| sh.is_row_filtered(*r)).collect()).unwrap_or_default();
        let sheet = cx.sheet_mut(si)?;
        for (c, _) in sheet.take_cells(old) {
            cx.touch(si, c);
        }
        let sheet = cx.sheet_mut(si)?;
        for r in filtered {
            if let Some(line) = sheet.rows.get_mut(&r) {
                line.hidden = false;
            }
        }
        write_rows(cx, si, start, &headers, &r)?;
        let range = RangeRef::new(start, new_end);
        let sheet = cx.sheet_mut(si)?;
        for (off, cell) in totals {
            if off < headers.len() as u32 {
                sheet.set_cell(CellRef::new(new_end.row, start.col + off), cell);
            }
        }
        if let Some(t) = sheet.tables.get_mut(ti) {
            let mut cols = vec![];
            for n in &headers {
                cols.push(t.columns.iter().find(|c| c.name.eq_ignore_ascii_case(n)).cloned().unwrap_or_else(|| column(n)));
            }
            t.columns = cols;
            t.range = range;
            t.filter.clear();
        }
        if let Some(c) = cx.wb.connections.get_mut(ci) {
            c.refreshed = Some(now);
            c.rows = rows as u64;
        }
        cx.structural = true;
        Ok(range)
    })?;
    super::dbwrite::remember(s, conn.id, &r.columns, &r.rows);
    if r.truncated {
        s.ui_requests.push(crate::UiRequest::Message(format!(
            "The query returned more rows than a worksheet holds. The first {} rows were loaded.",
            datasource::MAX_ROWS
        )));
    }
    Ok(json!({"connection": conn.name, "table": conn.table, "range": range.a1(), "rows": rows, "cols": headers.len()}))
}

/// Refreshes every connection (Refresh All): (refreshed, errors).
pub(crate) fn refresh_connections(s: &mut Session) -> Result<(usize, Vec<Json>)> {
    let n = s.doc()?.wb.connections.len();
    let mut ok = 0;
    let mut errors = vec![];
    for ci in 0..n {
        let name = s.doc()?.wb.connections.get(ci).map(|c| c.name.clone()).unwrap_or_default();
        match refresh_one(s, ci) {
            Ok(_) => ok += 1,
            Err(e) => errors.push(json!({"connection": name, "error": e.to_string()})),
        }
    }
    Ok((ok, errors))
}

// ---------------------------------------------------------------- settings and properties

fn data_source_settings(s: &mut Session, p: &Json) -> Result<Json> {
    const C: &str = "data.dataSourceSettings";
    let Some(server) = str_param(p, "server").map(str::trim).filter(|v| !v.is_empty()) else {
        return Ok(Json::Array(s.credential_servers().into_iter().map(|(server, user)| json!({"server": server, "user": user})).collect()));
    };
    if bool_param(p, "clear").unwrap_or(false) {
        return Ok(json!({"cleared": s.clear_credential(server)}));
    }
    let user = str_param(p, "user").ok_or_else(|| bad(C, "missing `user`"))?;
    let password = str_param(p, "password").ok_or_else(|| bad(C, "missing `password`"))?;
    s.set_credential(server, Credential { user: user.to_string(), password: password.to_string() });
    Ok(json!({"server": datasource::credential_key(server), "user": user}))
}

fn properties(s: &mut Session, p: &Json) -> Result<Json> {
    const C: &str = "data.connectionProperties";
    let ci = find_connection(s, p, C)?;
    let new_name = str_param(p, "name").map(str::trim).filter(|n| !n.is_empty()).map(str::to_string);
    if let Some(n) = &new_name {
        let taken = s.doc()?.wb.connections.iter().enumerate().any(|(i, c)| i != ci && c.name.eq_ignore_ascii_case(n));
        if taken || n.chars().count() > 255 {
            return Err(EngineError::Other("A query with this name already exists. Enter a unique name.".into()));
        }
    }
    let command = command_param(C, p, "table")?;
    let minutes = f64_param(p, "refreshMinutes").filter(|m| m.is_finite()).map(|m| m.clamp(0.0, 32_767.0) as u32);
    edit(s, |cx| {
        let c = cx.wb.connections.get_mut(ci).ok_or_else(|| bad(C, "no such query"))?;
        if let Some(n) = new_name {
            c.name = n;
        }
        if let Some(d) = str_param(p, "description") {
            c.description = d.chars().take(255).collect();
        }
        if let Some(b) = bool_param(p, "refreshOnOpen") {
            c.refresh_on_open = b;
        }
        if let Some(m) = minutes {
            c.refresh_minutes = m;
        }
        if let Some(u) = str_param(p, "user") {
            c.user = u.to_string();
        }
        if let Some(cmd) = command {
            if matches!(cmd, QueryCommand::Sql { .. }) && c.source.database().is_empty() {
                return Err(bad(C, "a SQL statement needs a database"));
            }
            c.command = cmd;
        }
        Ok(connection_json(c))
    })
}

fn delete_connection(s: &mut Session, p: &Json) -> Result<Json> {
    let ci = find_connection(s, p, "data.deleteConnection")?;
    edit(s, |cx| {
        if ci < cx.wb.connections.len() {
            cx.wb.connections.remove(ci);
        }
        Ok(Json::Null)
    })
}

pub(crate) fn connection_json(c: &Connection) -> Json {
    let (kind, command) = match &c.command {
        QueryCommand::Table { schema, name } => ("table", object_label(schema, name)),
        QueryCommand::Sql { text } => ("sql", text.clone()),
    };
    json!({
        "name": c.name,
        "description": c.description,
        "server": c.source.server(),
        "database": c.source.database(),
        "user": c.user,
        kind: command,
        "table": c.table,
        "refreshOnOpen": c.refresh_on_open,
        "refreshMinutes": c.refresh_minutes,
        "refreshed": c.refreshed,
        "rows": c.rows,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::datasource::{DataProvider, DbError, DbObject};
    use std::sync::Mutex;

    /// A provider serving fixed rows, counting calls and checking the login.
    struct Fake {
        rows: Mutex<Vec<Vec<DbValue>>>,
        calls: Mutex<Vec<String>>,
    }

    impl Fake {
        fn new() -> Arc<Fake> {
            Arc::new(Fake {
                rows: Mutex::new(vec![
                    vec![DbValue::Number(1.0), DbValue::Text("Ada".into()), DbValue::Date(45000.0), DbValue::Bool(true)],
                    vec![DbValue::Number(2.0), DbValue::Text("=1+1".into()), DbValue::Null, DbValue::Bool(false)],
                ]),
                calls: Mutex::new(vec![]),
            })
        }
        fn check(c: &Credential) -> std::result::Result<(), DbError> {
            if c.password == "secret" { Ok(()) } else { Err(DbError::Credentials("Login failed.".into())) }
        }
    }

    impl DataProvider for Fake {
        fn databases(&self, _: &DataSource, c: &Credential) -> std::result::Result<Vec<String>, DbError> {
            Fake::check(c)?;
            Ok(vec!["master".into(), "Shop".into()])
        }
        fn objects(&self, _: &DataSource, c: &Credential) -> std::result::Result<Vec<DbObject>, DbError> {
            Fake::check(c)?;
            Ok(vec![DbObject { schema: "dbo".into(), name: "Customers".into(), view: false }])
        }
        fn query(&self, _: &DataSource, c: &Credential, cmd: &QueryCommand, max: usize) -> std::result::Result<QueryResult, DbError> {
            Fake::check(c)?;
            self.calls.lock().map_err(|_| DbError::Other("poisoned".into()))?.push(format!("{cmd:?}"));
            let rows = self.rows.lock().map_err(|_| DbError::Other("poisoned".into()))?.clone();
            let truncated = rows.len() > max;
            Ok(QueryResult {
                columns: vec!["Id".into(), "Name".into(), "Since".into(), "Active".into()],
                rows: rows.into_iter().take(max).collect(),
                truncated,
            })
        }
    }

    fn session(fake: &Arc<Fake>) -> Session {
        let mut s = Session::new();
        s.execute("file.new", json!({})).unwrap();
        s.set_data_provider(fake.clone());
        s
    }

    fn login() -> Json {
        json!({"server": "db.local", "database": "Shop", "table": "dbo.Customers", "user": "sa", "password": "secret"})
    }

    #[test]
    fn loads_a_table_into_a_new_sheet() {
        let fake = Fake::new();
        let mut s = session(&fake);
        let r = s.execute("data.fromSqlServer", login()).unwrap();
        assert_eq!(r["loaded"][0]["sheet"], "dbo Customers");
        assert_eq!(r["loaded"][0]["table"], "dbo_Customers");
        assert_eq!(r["loaded"][0]["rows"], 2);
        let d = s.doc().unwrap();
        let sh = d.wb.active().unwrap();
        assert_eq!(sh.name, "dbo Customers");
        assert_eq!(sh.value(CellRef::parse("B1").unwrap()).display(), "Name");
        assert_eq!(sh.value(CellRef::parse("A2").unwrap()), Value::Number(1.0));
        // Text stays text, even when it looks like a formula.
        assert_eq!(sh.value(CellRef::parse("B3").unwrap()), Value::text("=1+1"));
        assert!(sh.cell(CellRef::parse("B3").unwrap()).unwrap().formula.is_none());
        // Dates get a date format; NULL stays empty.
        let c2 = sh.cell(CellRef::parse("C2").unwrap()).unwrap();
        assert_eq!(d.wb.styles.get(c2.style).num_fmt.0.as_ref(), "m/d/yyyy");
        assert_eq!(sh.value(CellRef::parse("C3").unwrap()), Value::Empty);
        assert_eq!(sh.tables[0].range.a1(), "A1:D3");
        let c = &d.wb.connections[0];
        assert_eq!((c.name.as_str(), c.table.as_str(), c.user.as_str()), ("dbo Customers", "dbo_Customers", "sa"));
        // One undo step, and the journal keeps no password.
        assert_eq!(d.undo.len(), 1);
        let (id, params) = s.journal.last().unwrap();
        assert_eq!(id, "data.fromSqlServer");
        assert!(params.get("password").is_none());
        s.execute("edit.undo", json!({})).unwrap();
        assert!(s.doc().unwrap().wb.connections.is_empty());
        assert_eq!(s.doc().unwrap().wb.sheets.len(), 1);
    }

    #[test]
    fn needs_a_login_and_a_provider() {
        let fake = Fake::new();
        let mut s = session(&fake);
        let e = s.execute("data.fromSqlServer", json!({"server": "db.local", "table": "dbo.Customers"})).unwrap_err().to_string();
        assert!(e.contains(datasource::CREDENTIALS_REQUIRED), "{e}");
        let e =
            s.execute("data.fromSqlServer", json!({"server": "db.local", "table": "t", "user": "sa", "password": "nope"})).unwrap_err().to_string();
        assert!(e.contains(datasource::CREDENTIALS_REQUIRED), "{e}");
        let mut bare = Session::new();
        bare.execute("file.new", json!({})).unwrap();
        let e = bare.execute("data.fromSqlServer", login()).unwrap_err().to_string();
        assert!(e.contains(datasource::NOT_AVAILABLE), "{e}");
        // No server: the menu opens the dialog; headless callers get an error.
        assert!(bare.execute_headless("data.fromSqlServer", json!({})).is_err());
        bare.execute("data.fromSqlServer", json!({})).unwrap();
        assert!(matches!(bare.take_ui_requests().as_slice(), [crate::UiRequest::Dialog(n, _)] if n == "sqlServer"));
        // A SQL statement needs a database.
        let e = s.execute("data.fromSqlServer", json!({"server": "x", "sql": "SELECT 1", "user": "sa", "password": "secret"})).unwrap_err();
        assert!(e.to_string().contains("database"));
    }

    #[test]
    fn navigator_and_preview() {
        let fake = Fake::new();
        let mut s = session(&fake);
        let r = s.execute("data.sqlServerNavigator", json!({"server": "db.local", "user": "sa", "password": "secret"})).unwrap();
        assert_eq!(r["databases"][1], "Shop");
        let r = s.execute("data.sqlServerNavigator", json!({"server": "DB.local ", "database": "Shop"})).unwrap();
        assert_eq!(r["objects"][0]["name"], "dbo.Customers");
        let r = s.execute("data.sqlServerPreview", json!({"server": "db.local", "database": "Shop", "table": "dbo.Customers", "rows": 1})).unwrap();
        assert_eq!(r["rows"][0], json!(["1", "Ada", "3/15/2023", "TRUE"]));
        assert_eq!(r["truncated"], true);
        // Data Source Settings lists the login without its password, and forgets it.
        let list = s.execute("data.dataSourceSettings", json!({})).unwrap();
        assert_eq!(list, json!([{"server": "db.local", "user": "sa"}]));
        s.execute("data.dataSourceSettings", json!({"server": "db.local", "clear": true})).unwrap();
        assert!(s.execute("data.sqlServerNavigator", json!({"server": "db.local"})).is_err());
    }

    #[test]
    fn refresh_rewrites_the_table() {
        let fake = Fake::new();
        let mut s = session(&fake);
        s.execute("data.fromSqlServer", login()).unwrap();
        fake.rows.lock().unwrap().push(vec![DbValue::Number(3.0), DbValue::Text("Grace".into()), DbValue::Null, DbValue::Null]);
        // A formula next to the table that reads it.
        s.execute("cell.set", json!({"cell": "F1", "input": "=ROWS(dbo_Customers)"})).unwrap();
        let r = s.execute("data.refresh", json!({"table": "dbo_Customers"})).unwrap();
        assert_eq!(r["rows"], 3);
        let d = s.doc().unwrap();
        let sh = d.wb.active().unwrap();
        assert_eq!(sh.tables[0].range.a1(), "A1:D4");
        assert_eq!(sh.value(CellRef::parse("B4").unwrap()).display(), "Grace");
        assert_eq!(sh.value(CellRef::parse("F1").unwrap()), Value::Number(3.0));
        assert_eq!(d.wb.connections[0].rows, 3);
        // Shrinking clears the rows that are gone.
        fake.rows.lock().unwrap().truncate(1);
        s.execute("selection.set", json!({"range": "A2"})).unwrap();
        s.execute("data.refresh", json!({})).unwrap();
        let sh = s.doc().unwrap().wb.active().unwrap();
        assert_eq!(sh.tables[0].range.a1(), "A1:D2");
        assert_eq!(sh.value(CellRef::parse("A3").unwrap()), Value::Empty);
        // A cell in the way stops the refresh instead of being overwritten.
        fake.rows.lock().unwrap().push(vec![DbValue::Number(9.0)]);
        s.execute("cell.set", json!({"cell": "A3", "input": "keep me"})).unwrap();
        assert!(s.execute("data.refresh", json!({"table": "dbo_Customers"})).is_err());
        assert_eq!(s.doc().unwrap().wb.active().unwrap().value(CellRef::parse("A3").unwrap()).display(), "keep me");
        // Refresh All reports it among its errors and still succeeds.
        let r = s.execute("data.refreshAll", json!({})).unwrap();
        assert_eq!(r["errors"].as_array().unwrap().len(), 1);
    }

    #[test]
    fn properties_rename_and_delete() {
        let fake = Fake::new();
        let mut s = session(&fake);
        s.execute("data.fromSqlServer", login()).unwrap();
        let r = s
            .execute(
                "data.connectionProperties",
                json!({"connection": "dbo Customers", "name": "Clients", "refreshOnOpen": true, "sql": "SELECT * FROM x"}),
            )
            .unwrap();
        assert_eq!(r["name"], "Clients");
        assert_eq!(r["sql"], "SELECT * FROM x");
        assert!(s.doc().unwrap().wb.connections[0].refresh_on_open);
        // Renaming the table keeps its query.
        s.execute("table.rename", json!({"table": "dbo_Customers", "name": "Clients"})).unwrap();
        assert_eq!(s.doc().unwrap().wb.connections[0].table, "Clients");
        s.execute("data.refresh", json!({"connection": "Clients"})).unwrap();
        let q = s.execute("data.queriesConnections", json!({})).unwrap();
        assert_eq!(q["connections"][0]["name"], "Clients");
        s.execute("data.deleteConnection", json!({"connection": "Clients"})).unwrap();
        assert!(s.doc().unwrap().wb.connections.is_empty());
        assert!(s.doc().unwrap().wb.table("Clients").is_some());
    }

    #[test]
    fn queries_survive_xlsx() {
        let fake = Fake::new();
        let mut s = session(&fake);
        s.execute("data.fromSqlServer", login()).unwrap();
        s.execute("data.connectionProperties", json!({"connection": "dbo Customers", "refreshOnOpen": true, "refreshMinutes": 30.0})).unwrap();
        let bytes = gridcraft_xlsx::write_xlsx(&s.doc().unwrap().wb).unwrap();
        let (wb, _) = gridcraft_xlsx::read_xlsx(&bytes).unwrap();
        assert_eq!(wb.connections.len(), 1);
        let c = &wb.connections[0];
        assert_eq!((c.name.as_str(), c.table.as_str(), c.user.as_str()), ("dbo Customers", "dbo_Customers", "sa"));
        assert_eq!(c.command, QueryCommand::Table { schema: "dbo".into(), name: "Customers".into() });
        assert_eq!(c.source.server(), "db.local");
        assert_eq!(c.source.database(), "Shop");
        assert!(c.refresh_on_open);
        assert_eq!(c.refresh_minutes, 30);
        assert_eq!(c.rows, 2);
        // The password never reaches the file.
        let text = String::from_utf8_lossy(&bytes).to_string();
        assert!(!text.contains("secret"));
        // Saved again from what was read, it stays one connection (not kept twice).
        let again = gridcraft_xlsx::write_xlsx(&wb).unwrap();
        let (wb2, _) = gridcraft_xlsx::read_xlsx(&again).unwrap();
        assert_eq!(wb2.connections.len(), 1);
        assert!(wb2.preserved.as_ref().is_none_or(|p| p.parts.iter().all(|q| !q.name.contains("connections"))));
        // And it refreshes after reopening, once signed in.
        let mut s2 = Session::new();
        s2.set_data_provider(fake.clone());
        s2.add_document(crate::DocState::new(wb2, None, "Book".into()));
        s2.set_credential("db.local", Credential { user: "sa".into(), password: "secret".into() });
        let r = s2.execute("data.refreshAll", json!({})).unwrap();
        assert_eq!(r["refreshed"], 1, "{r}");
    }

    #[test]
    fn names_are_made_valid_and_unique() {
        let wb = Workbook::new();
        assert_eq!(table_name_for(&wb, "dbo Order Details"), "dbo_Order_Details");
        assert_eq!(table_name_for(&wb, "1st"), "_1st");
        assert_eq!(table_name_for(&wb, "A1"), "_A1");
        assert_eq!(sheet_name_for(&wb, "a/b:c"), "abc");
        assert_eq!(sheet_name_for(&wb, "Sheet1"), "Sheet1 (2)");
        assert_eq!(header_names(&["a".into(), "A".into(), "".into()]), vec!["a", "A2", "Column3"]);
        assert_eq!(object_label("sales", "Order Details"), "sales.[Order Details]");
        assert_eq!(object_label("", "x]y"), "[x]]y]");
    }

    #[test]
    fn hostile_params_never_panic() {
        let fake = Fake::new();
        let mut s = session(&fake);
        for id in [
            "data.fromSqlServer",
            "data.sqlServerNavigator",
            "data.sqlServerPreview",
            "data.refresh",
            "data.dataSourceSettings",
            "data.connectionProperties",
            "data.deleteConnection",
        ] {
            for p in [
                json!({}),
                json!({"server": ""}),
                json!({"server": "x".repeat(10_000)}),
                json!({"server": "x", "table": "[unclosed"}),
                json!({"server": "x", "tables": [1, 2]}),
                json!({"server": "x", "rows": -1, "table": "t"}),
                json!({"connection": "none"}),
                json!({"server": "x", "database": "d", "sql": "SELECT 1", "user": "sa", "password": "secret", "refreshMinutes": u64::MAX}),
            ] {
                let r = s.execute(id, p.clone());
                assert!(!matches!(r, Err(EngineError::Internal(..))), "{id} {p}: {r:?}");
            }
        }
    }
}

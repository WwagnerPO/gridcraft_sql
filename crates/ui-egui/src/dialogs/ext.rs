//! The database and macro dialogs (Get Data › From SQL Server, Data Source Settings, Query
//! Properties, the database write dialogs, Developer › Macros). `dialogs.rs` reaches them through
//! [`open`], [`width`], [`owns`] and [`show`], one line each.

use serde_json::{Value as Json, json};

use super::Dialog;
use super::Field::*;
use crate::SheetApp;

/// The dialogs drawn here.
const NAMES: &[&str] = &["sqlServer", "existingConnections", "dataSourceSettings", "saveToDatabase", "exportToSqlServer", "sqlExecute", "macros"];

impl Dialog {
    /// The sign-in step of the SQL Server dialog on its own: Refresh asks for a login with it,
    /// then runs `retry` (a refresh command) with `params` again. An empty `retry` only stores it.
    pub fn login(server: &str, user: &str, retry: &str, params: Json) -> Option<Dialog> {
        Some(Dialog::custom(
            "sqlServer",
            "SQL Server database",
            json!({"stage": "credentials", "server": server, "user": user, "retry": retry, "retryParams": params, "encrypt": true}),
        ))
    }
}

/// Opens dialog `name` when it is one of these: `Some(dialog)` (`Some(None)` when it can't open,
/// after saying why); `None` for every other dialog.
pub(super) fn open(app: &mut SheetApp, name: &str, params: &Json) -> Option<Option<Dialog>> {
    let p = |k: &str| params.get(k).and_then(Json::as_str).unwrap_or("").to_string();
    Some(match name {
        "macros" => {
            let mut d = Dialog::custom("macros", "Macro", json!({}));
            // What `view.macros` lists (recorded scripts, the VBA project), as the upstream dialog.
            d.result = app.session.run("view.macros", json!({})).ok();
            Some(d)
        }
        "sqlServer" => Some(Dialog::custom(
            "sqlServer",
            "SQL Server database",
            json!({"stage": "source", "server": p("server"), "database": p("database"), "encrypt": true, "trust": false}),
        )),
        "dataSourceSettings" => Some(Dialog::custom("dataSourceSettings", "Data Source Settings", json!({}))),
        "existingConnections" => Some(Dialog::custom("existingConnections", "Existing Connections", json!({}))),
        // Write dialogs take their values back from `params` after a sign-in.
        "saveToDatabase" => Some(Dialog::custom("saveToDatabase", "Save Changes to Database", params.clone())),
        "exportToSqlServer" => Some(Dialog::custom("exportToSqlServer", "Export to SQL Server", params.clone())),
        "sqlExecute" => Some(Dialog::custom("sqlExecute", "Run SQL", params.clone())),
        "connectionProperties" => connection_properties(app, params),
        _ => return None,
    })
}

fn connection_properties(app: &mut SheetApp, params: &Json) -> Option<Dialog> {
    let name = params.get("connection").and_then(Json::as_str).map(str::to_string).or_else(|| {
        let doc = app.session.active()?;
        let t = doc.wb.active()?.table_at(doc.selection.active)?;
        doc.wb.connections.iter().find(|c| c.table.eq_ignore_ascii_case(&t.name)).map(|c| c.name.clone())
    });
    let conn = name.and_then(|n| app.session.active()?.wb.connections.iter().find(|c| c.name.eq_ignore_ascii_case(&n)).cloned());
    let Some(c) = conn else {
        app.message =
            Some(("Connection Properties".into(), "Select a cell in a table loaded from a database, or open Queries & Connections.".into()));
        return None;
    };
    let sql = match &c.command {
        gridcraft_engine::model::QueryCommand::Sql { text } => text.clone(),
        _ => String::new(),
    };
    let mut fields = vec![
        Text { key: "name", label: "Query name:" },
        Text { key: "description", label: "Description:" },
        Check { key: "refreshOnOpen", label: "Refresh data when opening the file" },
        Number { key: "refreshMinutes", label: "Refresh every (minutes, 0 = never):" },
    ];
    if !sql.is_empty() {
        fields.push(Text { key: "sql", label: "SQL statement:" });
    }
    Some(Dialog::form(
        "connectionProperties",
        "Query Properties",
        "data.connectionProperties",
        fields,
        json!({
            "connection": c.name,
            "name": c.name,
            "description": c.description,
            "refreshOnOpen": c.refresh_on_open,
            "refreshMinutes": c.refresh_minutes,
            "sql": sql,
        }),
    ))
}

/// The window width of these dialogs.
pub(super) fn width(d: &Dialog) -> Option<f32> {
    match d.name.as_str() {
        "sqlServer" if matches!(d.values.get("stage").and_then(Json::as_str), Some("navigator" | "preview")) => Some(680.0),
        "existingConnections" => Some(500.0),
        "sqlServer" | "dataSourceSettings" | "saveToDatabase" | "exportToSqlServer" | "sqlExecute" => Some(440.0),
        _ => None,
    }
}

/// Whether dialog `name` is drawn here.
pub(super) fn owns(name: &str) -> bool {
    NAMES.contains(&name)
}

/// Draws one of these dialogs.
pub(super) fn show(app: &mut SheetApp, ui: &mut egui::Ui, d: &mut Dialog, open: &mut bool) {
    match d.name.as_str() {
        "sqlServer" => crate::get_data::sql_server(app, ui, d, open),
        "dataSourceSettings" => crate::get_data::data_source_settings(app, ui, d, open),
        "existingConnections" => crate::get_data::existing_connections(app, ui, d, open),
        "saveToDatabase" => crate::get_data::save_to_database(app, ui, d, open),
        "exportToSqlServer" => crate::get_data::export_to_sql_server(app, ui, d, open),
        "sqlExecute" => crate::get_data::run_sql(app, ui, d, open),
        "macros" => crate::macros::dialog(app, ui, d, open),
        _ => {}
    }
}

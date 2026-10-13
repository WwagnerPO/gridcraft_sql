//! Get Data › From Database › From SQL Server Database: the source, sign-in, Navigator and
//! preview steps, Data Source Settings, the login prompt Refresh uses, and the Queries &
//! Connections pane. Every step acts through engine commands (`data.sqlServerNavigator`,
//! `data.sqlServerPreview`, `data.fromSqlServer`, `data.dataSourceSettings`, `data.refresh`…).

use egui::{Color32, RichText, Ui, vec2};
use gridcraft_engine::datasource::{CREDENTIALS_REQUIRED, ENCRYPTION_FAILED};
use serde_json::{Value as Json, json};

use crate::SheetApp;
use crate::dialogs::Dialog;

fn s(d: &Dialog, key: &str) -> String {
    d.values.get(key).and_then(Json::as_str).unwrap_or("").to_string()
}

fn b(d: &Dialog, key: &str, default: bool) -> bool {
    d.values.get(key).and_then(Json::as_bool).unwrap_or(default)
}

fn set(d: &mut Dialog, key: &str, v: Json) {
    d.values.insert(key.into(), v);
}

fn tr(app: &SheetApp, text: &'static str) -> &'static str {
    app.ui.language.tr(text)
}

fn error_label(ui: &mut Ui, d: &Dialog) {
    let e = s(d, "error");
    if !e.is_empty() {
        ui.add_space(4.0);
        ui.label(RichText::new(e).color(Color32::from_rgb(0xC4, 0x2B, 0x1C)));
    }
}

/// A row of buttons laid out from the right, only as tall as the buttons.
fn button_row<R>(ui: &mut Ui, add: impl FnOnce(&mut Ui) -> R) -> R {
    ui.allocate_ui_with_layout(vec2(ui.available_width(), 24.0), egui::Layout::right_to_left(egui::Align::Center), add).inner
}

fn primary(ui: &mut Ui, text: &str) -> bool {
    ui.add(egui::Button::new(RichText::new(text).color(Color32::WHITE)).fill(crate::theme::Tokens::get(ui.ctx()).accent).min_size(vec2(72.0, 0.0)))
        .clicked()
}

fn secondary(ui: &mut Ui, text: &str) -> bool {
    ui.add(egui::Button::new(text).min_size(vec2(72.0, 0.0))).clicked()
}

/// The source params every step sends.
fn source(d: &Dialog) -> Json {
    json!({
        "server": s(d, "server").trim(),
        "database": s(d, "database").trim(),
        "encrypt": b(d, "encrypt", true),
        "trustServerCertificate": b(d, "trust", false),
    })
}

/// The message of an engine error, without the command prefix.
fn message(e: &str) -> String {
    crate::clean_error(e)
}

/// Runs the step after the source (and sign-in): the Navigator, or the statement's preview.
fn connect(app: &mut SheetApp, d: &mut Dialog) {
    set(d, "error", json!(""));
    let mut p = source(d);
    let sql = s(d, "sql");
    let r = if sql.trim().is_empty() {
        app.run("data.sqlServerNavigator", p)
    } else {
        p["sql"] = json!(sql);
        p["rows"] = json!(50);
        app.run("data.sqlServerPreview", p)
    };
    match r {
        Ok(v) => {
            if sql.trim().is_empty() {
                set(d, "databases", v.get("databases").cloned().unwrap_or(Json::Null));
                set(d, "objects", v.get("objects").cloned().unwrap_or(json!([])));
                set(d, "stage", json!("navigator"));
            } else {
                set(d, "preview", v);
                set(d, "stage", json!("preview"));
            }
        }
        Err(e) if e.contains(CREDENTIALS_REQUIRED) => {
            let why = message(&e).replace(CREDENTIALS_REQUIRED, "").trim().to_string();
            set(d, "error", json!(if why.starts_with("No login is stored") { String::new() } else { why }));
            set(d, "stage", json!("credentials"));
        }
        Err(e) if e.contains(ENCRYPTION_FAILED) => {
            set(d, "encryptionFailed", json!(true));
            set(d, "error", json!(message(&e)));
            // The offer to trust the certificate shows on the step the user is on.
        }
        Err(e) => set(d, "error", json!(message(&e))),
    }
}

/// After a failed TLS handshake: offer to trust the server's certificate and connect again.
/// SQL Server checks the certificate even for an unencrypted connection (the login is always
/// encrypted), so trusting it is the way through for an expired or self-signed one.
fn certificate_help(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog) {
    if !b(d, "encryptionFailed", false) || b(d, "trust", false) {
        return;
    }
    ui.add_space(4.0);
    ui.label(tr(app, "You can trust this server's certificate (only on a network you trust), or ask the server's administrator to renew it."));
    if primary(ui, tr(app, "Trust the certificate and connect")) {
        set(d, "trust", json!(true));
        set(d, "encryptionFailed", json!(false));
        if s(d, "stage") == "credentials" && !s(d, "password").is_empty() {
            let (server, user, password) = (s(d, "server"), s(d, "user"), s(d, "password"));
            let _ = app.run("data.dataSourceSettings", json!({"server": server.trim(), "user": user.trim(), "password": password}));
        }
        connect(app, d);
    }
}

/// The SQL Server database dialog: `stage` is source → credentials → navigator | preview.
pub fn sql_server(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, open: &mut bool) {
    match s(d, "stage").as_str() {
        "credentials" => credentials(app, ui, d, open),
        "navigator" => navigator(app, ui, d, open),
        "preview" => statement_preview(app, ui, d, open),
        _ => source_step(app, ui, d, open),
    }
}

fn source_step(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, open: &mut bool) {
    let mut server = s(d, "server");
    let mut database = s(d, "database");
    let mut sql = s(d, "sql");
    let (mut encrypt, mut trust) = (b(d, "encrypt", true), b(d, "trust", false));
    ui.label(tr(app, "Server"));
    let r = ui.add(egui::TextEdit::singleline(&mut server).desired_width(f32::INFINITY).hint_text("host, host,port or host\\instance"));
    if s(d, "focused").is_empty() {
        r.request_focus();
        set(d, "focused", json!("1"));
    }
    ui.add_space(4.0);
    ui.label(tr(app, "Database (optional)"));
    ui.add(egui::TextEdit::singleline(&mut database).desired_width(f32::INFINITY));
    ui.add_space(4.0);
    egui::CollapsingHeader::new(tr(app, "Advanced options")).default_open(!sql.is_empty()).show(ui, |ui| {
        ui.label(tr(app, "SQL statement (optional, requires database)"));
        ui.add(egui::TextEdit::multiline(&mut sql).code_editor().desired_rows(5).desired_width(f32::INFINITY));
        ui.checkbox(&mut encrypt, tr(app, "Encrypt connection"));
        ui.checkbox(&mut trust, tr(app, "Trust the server certificate"));
    });
    set(d, "server", json!(server));
    set(d, "database", json!(database));
    set(d, "sql", json!(sql));
    set(d, "encrypt", json!(encrypt));
    set(d, "trust", json!(trust));
    error_label(ui, d);
    certificate_help(app, ui, d);
    ui.add_space(8.0);
    button_row(ui, |ui| {
        if secondary(ui, tr(app, "Cancel")) {
            *open = false;
        }
        if primary(ui, tr(app, "OK")) {
            if server.trim().is_empty() {
                set(d, "error", json!(tr(app, "Enter a server name.")));
            } else if !sql.trim().is_empty() && database.trim().is_empty() {
                set(d, "error", json!(tr(app, "A SQL statement needs a database.")));
            } else {
                connect(app, d);
            }
        }
    });
}

fn credentials(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, open: &mut bool) {
    let server = s(d, "server");
    ui.label(RichText::new(format!("{} {}", tr(app, "Sign in to"), server.trim())).strong());
    ui.label(tr(app, "Use a SQL Server login. GridCraft keeps it for this session and never saves it in the workbook."));
    ui.add_space(6.0);
    let mut user = s(d, "user");
    let mut password = s(d, "password");
    let mut show = b(d, "showPassword", false);
    egui::Grid::new("sql_login").num_columns(2).spacing(vec2(8.0, 6.0)).show(ui, |ui| {
        ui.label(tr(app, "User name"));
        ui.add(egui::TextEdit::singleline(&mut user).desired_width(220.0));
        ui.end_row();
        ui.label(tr(app, "Password"));
        ui.add(egui::TextEdit::singleline(&mut password).password(!show).desired_width(220.0));
        ui.end_row();
        ui.label("");
        ui.checkbox(&mut show, tr(app, "Show password"));
        ui.end_row();
    });
    set(d, "user", json!(user));
    set(d, "password", json!(password));
    set(d, "showPassword", json!(show));
    error_label(ui, d);
    certificate_help(app, ui, d);
    ui.add_space(8.0);
    let enter = ui.input(|i| i.key_pressed(egui::Key::Enter));
    ui.horizontal(|ui| {
        if !s(d, "retry").is_empty() {
            // A login asked for by Refresh: no source step to go back to.
        } else if ui.button(tr(app, "Back")).clicked() {
            set(d, "stage", json!("source"));
            set(d, "error", json!(""));
        }
        button_row(ui, |ui| {
            if secondary(ui, tr(app, "Cancel")) {
                *open = false;
            }
            if (primary(ui, tr(app, "Connect")) || enter) && !user.trim().is_empty() {
                if let Err(e) = app.run("data.dataSourceSettings", json!({"server": server.trim(), "user": user.trim(), "password": password})) {
                    set(d, "error", json!(message(&e)));
                    return;
                }
                set(d, "password", json!(""));
                let retry = s(d, "retry");
                if retry.is_empty() {
                    connect(app, d);
                } else {
                    *open = false;
                    let params = d.values.get("retryParams").cloned().unwrap_or(json!({}));
                    refresh_with_login(app, &retry, params);
                }
            }
        });
    });
}

fn navigator(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, open: &mut bool) {
    let databases: Vec<String> = d
        .values
        .get("databases")
        .and_then(Json::as_array)
        .map(|a| a.iter().filter_map(|x| x.as_str().map(str::to_string)).collect())
        .unwrap_or_default();
    let objects: Vec<Json> = d.values.get("objects").and_then(Json::as_array).cloned().unwrap_or_default();
    let mut selected: Vec<String> = d
        .values
        .get("selected")
        .and_then(Json::as_array)
        .map(|a| a.iter().filter_map(|x| x.as_str().map(str::to_string)).collect())
        .unwrap_or_default();
    let mut multi = b(d, "multi", false);
    let mut search = s(d, "search");
    let mut focus = s(d, "focus");
    let mut pick_db: Option<String> = None;
    let mut show_preview: Option<String> = None;
    // A fixed height: the separator between the columns would otherwise run down the screen.
    let body = vec2(ui.available_width(), 400.0);
    ui.allocate_ui_with_layout(body, egui::Layout::left_to_right(egui::Align::Min), |ui| {
        ui.set_min_height(body.y);
        ui.vertical(|ui| {
            ui.set_width(240.0);
            let current = s(d, "database");
            if !databases.is_empty() || !current.is_empty() {
                ui.label(tr(app, "Database"));
                egui::ComboBox::from_id_salt("sql_db")
                    .width(230.0)
                    .selected_text(if current.is_empty() { "—".to_string() } else { current.clone() })
                    .show_ui(ui, |ui| {
                        for db in &databases {
                            if ui.selectable_label(*db == current, db).clicked() {
                                pick_db = Some(db.clone());
                            }
                        }
                    });
                ui.add_space(4.0);
            }
            ui.add(egui::TextEdit::singleline(&mut search).hint_text(tr(app, "Search")).desired_width(230.0));
            if ui.checkbox(&mut multi, tr(app, "Select multiple items")).changed() && !multi {
                selected.truncate(1);
            }
            egui::ScrollArea::vertical().id_salt("sql_objects").max_height(340.0).show(ui, |ui| {
                if objects.is_empty() {
                    let hint = if current.is_empty() { "Choose a database." } else { "This database has no tables or views you can read." };
                    ui.label(RichText::new(tr(app, hint)).italics());
                }
                let q = search.to_lowercase();
                for o in &objects {
                    let name = o["name"].as_str().unwrap_or("").to_string();
                    if !q.is_empty() && !name.to_lowercase().contains(&q) {
                        continue;
                    }
                    let kind = if o["view"].as_bool().unwrap_or(false) { "view" } else { "table" };
                    let is_sel = selected.contains(&name);
                    let label = format!("{} {name}", if kind == "view" { "◇" } else { "▦" });
                    let r = if multi {
                        let mut on = is_sel;
                        let r = ui.checkbox(&mut on, label);
                        if r.changed() {
                            if on {
                                selected.push(name.clone());
                            } else {
                                selected.retain(|x| *x != name);
                            }
                        }
                        r
                    } else {
                        let r = ui.add(egui::Button::selectable(is_sel, label).min_size(vec2(230.0, 20.0)));
                        if r.clicked() {
                            selected = vec![name.clone()];
                        }
                        r
                    };
                    if r.clicked() && focus != name {
                        focus = name.clone();
                        show_preview = Some(name.clone());
                    }
                }
            });
        });
        ui.separator();
        ui.vertical(|ui| {
            ui.set_min_width(360.0);
            if focus.is_empty() {
                ui.label(RichText::new(tr(app, "Select a table or view to preview its data.")).italics());
            } else {
                ui.label(RichText::new(&focus).strong());
                error_label(ui, d);
                if let Some(p) = d.values.get("preview") {
                    preview_grid(ui, p, "sql_nav_preview");
                }
            }
        });
    });
    set(d, "multi", json!(multi));
    set(d, "search", json!(search));
    set(d, "focus", json!(focus));
    set(d, "selected", json!(selected));
    if let Some(db) = pick_db {
        set(d, "database", json!(db));
        set(d, "focus", json!(""));
        set(d, "selected", json!([]));
        set(d, "preview", Json::Null);
        connect(app, d);
    }
    if let Some(name) = show_preview {
        let mut p = source(d);
        p["table"] = json!(name);
        p["rows"] = json!(50);
        match app.run("data.sqlServerPreview", p) {
            Ok(v) => {
                set(d, "preview", v);
                set(d, "error", json!(""));
            }
            Err(e) => {
                set(d, "preview", Json::Null);
                set(d, "error", json!(message(&e)));
            }
        }
    }
    ui.add_space(8.0);
    button_row(ui, |ui| {
        if secondary(ui, tr(app, "Cancel")) {
            *open = false;
        }
        let selected: Vec<String> = d
            .values
            .get("selected")
            .and_then(Json::as_array)
            .map(|a| a.iter().filter_map(|x| x.as_str().map(str::to_string)).collect())
            .unwrap_or_default();
        let load = ui.add_enabled_ui(!selected.is_empty(), |ui| primary(ui, tr(app, "Load"))).inner;
        if load {
            let mut p = source(d);
            p["tables"] = json!(selected);
            load_query(app, d, p, open);
        }
    });
}

fn statement_preview(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, open: &mut bool) {
    ui.label(RichText::new(format!("{} · {}", s(d, "server").trim(), s(d, "database").trim())).strong());
    if let Some(p) = d.values.get("preview") {
        preview_grid(ui, p, "sql_statement_preview");
    }
    error_label(ui, d);
    ui.add_space(8.0);
    ui.horizontal(|ui| {
        if ui.button(tr(app, "Back")).clicked() {
            set(d, "stage", json!("source"));
            set(d, "error", json!(""));
        }
        button_row(ui, |ui| {
            if secondary(ui, tr(app, "Cancel")) {
                *open = false;
            }
            if primary(ui, tr(app, "Load")) {
                let mut p = source(d);
                p["sql"] = json!(s(d, "sql"));
                load_query(app, d, p, open);
            }
        });
    });
}

fn load_query(app: &mut SheetApp, d: &mut Dialog, p: Json, open: &mut bool) {
    match app.run("data.fromSqlServer", p) {
        Ok(r) => {
            *open = false;
            let n: u64 = r["loaded"].as_array().map(|a| a.iter().filter_map(|x| x["rows"].as_u64()).sum()).unwrap_or(0);
            app.toast = Some((format!("{n} {}", tr(app, "rows loaded.")), crate::now_ms()));
        }
        Err(e) => set(d, "error", json!(message(&e))),
    }
}

/// The first rows of a preview result (`{columns, rows}`), scrolling both ways.
fn preview_grid(ui: &mut Ui, p: &Json, id: &str) {
    let columns: Vec<String> = p["columns"].as_array().map(|a| a.iter().map(|c| c.as_str().unwrap_or("").to_string()).collect()).unwrap_or_default();
    let rows = p["rows"].as_array().cloned().unwrap_or_default();
    egui::ScrollArea::both().id_salt(id).max_height(320.0).max_width(ui.available_width()).show(ui, |ui| {
        egui::Grid::new(id).striped(true).num_columns(columns.len().max(1)).spacing(vec2(12.0, 3.0)).show(ui, |ui| {
            for c in &columns {
                ui.label(RichText::new(c).strong());
            }
            ui.end_row();
            for row in &rows {
                for v in row.as_array().into_iter().flatten() {
                    let t = v.as_str().map(str::to_string).unwrap_or_default();
                    let short: String = t.chars().take(40).collect();
                    let cell = if v.is_null() { RichText::new("null").italics().weak() } else { RichText::new(short) };
                    ui.label(cell);
                }
                ui.end_row();
            }
        });
    });
    if rows.is_empty() {
        ui.label(RichText::new("—").weak());
    }
}

/// Data Source Settings: the logins kept this session, to clear or replace.
pub fn data_source_settings(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, open: &mut bool) {
    let list = app.session.credential_servers();
    if list.is_empty() {
        ui.label(RichText::new(tr(app, "No database logins are kept in this session.")).italics());
    }
    let mut clear: Option<String> = None;
    let mut edit: Option<String> = None;
    egui::Grid::new("dss").striped(true).num_columns(3).spacing(vec2(12.0, 4.0)).show(ui, |ui| {
        for (server, user) in &list {
            ui.label(RichText::new(server).strong());
            ui.label(user);
            ui.horizontal(|ui| {
                if ui.small_button(tr(app, "Edit Permissions…")).clicked() {
                    edit = Some(server.clone());
                }
                if ui.small_button(tr(app, "Clear Permissions")).clicked() {
                    clear = Some(server.clone());
                }
            });
            ui.end_row();
        }
    });
    if let Some(server) = clear {
        app.run_or_alert("data.dataSourceSettings", json!({"server": server, "clear": true}));
    }
    if let Some(server) = edit {
        *open = false;
        let user = list.iter().find(|(s, _)| *s == server).map(|(_, u)| u.clone()).unwrap_or_default();
        app.dialog = Dialog::login(&server, &user, "", json!({}));
    }
    let _ = d;
    ui.add_space(8.0);
    button_row(ui, |ui| {
        if secondary(ui, tr(app, "Close")) {
            *open = false;
        }
    });
}

/// The servers of the connections a refresh command reaches (`data.refresh` the one named or at
/// the active cell; `data.refreshAll` all of them).
fn servers_for(app: &SheetApp, id: &str, params: &Json) -> Vec<String> {
    let Some(doc) = app.session.active() else { return vec![] };
    let conns = &doc.wb.connections;
    let pick: Vec<&gridcraft_engine::model::Connection> = if id == "data.refreshAll" {
        conns.iter().collect()
    } else if let Some(n) = params.get("connection").and_then(Json::as_str) {
        conns.iter().filter(|c| c.name.eq_ignore_ascii_case(n)).collect()
    } else {
        let table = params
            .get("table")
            .and_then(Json::as_str)
            .map(str::to_string)
            .or_else(|| doc.wb.active().and_then(|sh| sh.table_at(doc.selection.active)).map(|t| t.name.clone()));
        conns.iter().filter(|c| table.as_deref().is_some_and(|t| c.table.eq_ignore_ascii_case(t))).collect()
    };
    let mut out: Vec<String> = vec![];
    for c in pick {
        let s = c.source.server().to_string();
        if !out.iter().any(|x| x.eq_ignore_ascii_case(&s)) {
            out.push(s);
        }
    }
    out
}

/// Runs a refresh command, first asking for the login of a server that has none (or whose login
/// was refused), then running it again. Query errors Refresh All collects are shown together.
pub fn refresh_with_login(app: &mut SheetApp, id: &str, params: Json) {
    let servers = servers_for(app, id, &params);
    let user_of = |app: &SheetApp, server: &str| {
        app.session
            .active()
            .and_then(|d| d.wb.connections.iter().find(|c| c.source.server().eq_ignore_ascii_case(server)).map(|c| c.user.clone()))
            .unwrap_or_default()
    };
    if let Some(server) = servers.iter().find(|s| app.session.credential(s).is_none()) {
        let user = user_of(app, server);
        app.dialog = Dialog::login(server, &user, id, params);
        return;
    }
    let r = app.run(id, params.clone());
    let refused = |e: &str| e.contains(CREDENTIALS_REQUIRED);
    match r {
        Err(e) if refused(&e) => {
            if let Some(server) = servers.first() {
                let user = user_of(app, server);
                app.dialog = Dialog::login(server, &user, id, params);
                if let Some(d) = app.dialog.as_mut() {
                    d.values.insert("error".into(), json!(message(&e).replace(CREDENTIALS_REQUIRED, "").trim()));
                }
            }
        }
        Err(e) => app.message = Some((tr(app, "Refresh").into(), message(&e))),
        Ok(v) => {
            let errors: Vec<String> = v["errors"]
                .as_array()
                .map(|a| {
                    a.iter()
                        .map(|x| {
                            let name = x["connection"].as_str().or(x["pivot"].as_str()).unwrap_or("");
                            format!("{name}: {}", message(x["error"].as_str().unwrap_or("")))
                        })
                        .collect()
                })
                .unwrap_or_default();
            if let Some(first) = v["errors"].as_array().and_then(|a| a.iter().find(|x| x["error"].as_str().is_some_and(refused)))
                && let Some(name) = first["connection"].as_str()
            {
                let server =
                    app.session.active().and_then(|d| d.wb.connections.iter().find(|c| c.name == name).map(|c| c.source.server().to_string()));
                if let Some(server) = server {
                    let user = user_of(app, &server);
                    app.dialog = Dialog::login(&server, &user, id, params);
                    return;
                }
            }
            if !errors.is_empty() {
                app.message = Some((tr(app, "Refresh All").into(), errors.join("\n")));
            }
        }
    }
}

/// The Queries & Connections pane.
pub fn queries_pane(app: &mut SheetApp, ui: &mut Ui) {
    let conns: Vec<gridcraft_engine::model::Connection> = app.session.active().map(|d| d.wb.connections.clone()).unwrap_or_default();
    if conns.is_empty() {
        ui.label(RichText::new(tr(app, "This workbook has no queries. Use Get Data to load data from a database.")).italics());
        return;
    }
    ui.label(format!("{} {}", conns.len(), tr(app, if conns.len() == 1 { "query" } else { "queries" })));
    ui.add_space(4.0);
    for c in conns {
        ui.group(|ui| {
            ui.set_width(ui.available_width());
            ui.label(RichText::new(&c.name).strong());
            let when = c
                .refreshed
                .and_then(|t| {
                    gridcraft_engine::numfmt::text_function(
                        &gridcraft_engine::core::Value::number(t),
                        "m/d/yyyy h:mm",
                        gridcraft_engine::core::date::DateSystem::D1900,
                    )
                    .ok()
                })
                .unwrap_or_default();
            ui.label(RichText::new(format!("{} {}", c.rows, tr(app, "rows loaded."))).small());
            if !when.is_empty() {
                ui.label(RichText::new(format!("{} {when} (UTC)", tr(app, "Last refreshed"))).small().weak());
            }
            ui.label(RichText::new(format!("{} · {}", c.source.server(), c.source.database())).small().weak());
            ui.horizontal(|ui| {
                if ui.small_button(tr(app, "Refresh")).clicked() {
                    refresh_with_login(app, "data.refresh", json!({"connection": c.name}));
                }
                if ui.small_button(tr(app, "Properties…")).clicked() {
                    app.open_dialog("connectionProperties", json!({"connection": c.name}));
                }
                if ui.small_button(tr(app, "Delete")).clicked() {
                    app.run_or_alert("data.deleteConnection", json!({"connection": c.name}));
                }
            });
        });
    }
}

/// Refreshes the active workbook's queries marked "Refresh data when opening the file".
pub fn refresh_on_open(app: &mut SheetApp) {
    let names: Vec<String> =
        app.session.active().map(|d| d.wb.connections.iter().filter(|c| c.refresh_on_open).map(|c| c.name.clone()).collect()).unwrap_or_default();
    for name in names {
        refresh_with_login(app, "data.refresh", json!({"connection": name}));
        if app.dialog.is_some() {
            break;
        }
    }
}

/// Refreshes queries with a refresh interval when it has passed (the active workbook, signed in,
/// nothing being edited) and wakes the interface for the next one.
pub fn tick(app: &mut SheetApp, ctx: &egui::Context) {
    let Some(doc) = app.session.active() else { return };
    let uid = doc.uid;
    let timed: Vec<(String, u32, String)> = doc
        .wb
        .connections
        .iter()
        .filter(|c| c.refresh_minutes > 0)
        .map(|c| (c.name.clone(), c.refresh_minutes, c.source.server().to_string()))
        .collect();
    if timed.is_empty() {
        return;
    }
    let now = crate::now_ms();
    let mut next_due = f64::INFINITY;
    for (name, minutes, server) in timed {
        let period = minutes as f64 * 60_000.0;
        let last = match app.query_timers.iter().find(|(u, n, _)| *u == uid && *n == name) {
            Some((_, _, t)) => *t,
            None => {
                app.query_timers.push((uid, name.clone(), now));
                now
            }
        };
        let busy = app.dialog.is_some() || app.editor.is_some() || app.session.credential(&server).is_none();
        if now - last >= period && !busy {
            if let Err(e) = app.run("data.refresh", json!({"connection": name})) {
                app.toast = Some((format!("{name}: {}", message(&e)), now));
            }
            if let Some(t) = app.query_timers.iter_mut().find(|(u, n, _)| *u == uid && *n == name) {
                t.2 = now;
            }
            next_due = next_due.min(period);
        } else {
            next_due = next_due.min((period - (now - last)).max(1_000.0));
        }
    }
    if next_due.is_finite() {
        ctx.request_repaint_after(std::time::Duration::from_millis(next_due as u64));
    }
}

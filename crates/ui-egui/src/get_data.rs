//! Get Data › From Database › From SQL Server Database: the source, sign-in, Navigator and
//! preview steps, Data Source Settings, the login prompt Refresh uses, and the Queries &
//! Connections pane. Every step acts through engine commands (`data.sqlServerNavigator`,
//! `data.sqlServerPreview`, `data.fromSqlServer`, `data.dataSourceSettings`, `data.refresh`…).

use egui::{Color32, RichText, Ui, vec2};
use gridcraft_engine::datasource::{CREDENTIALS_REQUIRED, ENCRYPTION_FAILED};
use serde_json::{Value as Json, json};

use crate::SheetApp;
use crate::dialogs::Dialog;

/// Keeps database logins between sessions (the system's credential store). Keys are server names
/// as typed; the store decides how to normalise them.
pub trait LoginStore {
    /// The remembered (user, password) for `server`.
    fn load(&self, server: &str) -> Option<(String, String)>;
    fn save(&self, server: &str, user: &str, password: &str) -> Result<(), String>;
    fn forget(&self, server: &str);
}

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
        "auth": auth(d),
    })
}

/// How the dialog signs in: `sql` (default), `windows` or `integrated`.
fn auth(d: &Dialog) -> String {
    match s(d, "auth").as_str() {
        a @ ("windows" | "integrated") => a.to_string(),
        _ => "sql".to_string(),
    }
}

/// The sign-in choices, as Excel's Windows / Database tabs.
fn auth_choice(app: &SheetApp, ui: &mut Ui, d: &mut Dialog) {
    let mut a = auth(d);
    ui.horizontal_wrapped(|ui| {
        ui.radio_value(&mut a, "sql".to_string(), tr(app, "SQL Server login"));
        ui.radio_value(&mut a, "windows".to_string(), tr(app, "Windows (another account)"));
        ui.radio_value(&mut a, "integrated".to_string(), tr(app, "Windows (current user)"));
    });
    set(d, "auth", json!(a));
}

/// The message of an engine error, without the command prefix.
fn message(e: &str) -> String {
    crate::clean_error(e)
}

/// Puts a login remembered on this computer into the session when it has none for `server`.
pub fn recall_login(app: &mut SheetApp, server: &str) {
    let server = server.trim();
    if server.is_empty() || app.session.credential(server).is_some() {
        return;
    }
    if let Some((user, password)) = app.services.logins.as_ref().and_then(|l| l.load(server)) {
        app.session.set_credential(server, gridcraft_engine::datasource::Credential { user, password });
    }
}

/// Runs the step after the source (and sign-in): the Navigator, or the statement's preview.
fn connect(app: &mut SheetApp, d: &mut Dialog) {
    set(d, "error", json!(""));
    recall_login(app, &s(d, "server"));
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
    // A refresh or write asking again keeps the connection's way of signing in.
    if s(d, "retry").is_empty() {
        auth_choice(app, ui, d);
    }
    let integrated = auth(d) == "integrated";
    ui.label(tr(
        app,
        if integrated {
            "GridCraft signs in as the Windows account you're using; no password is asked."
        } else {
            "GridCraft keeps the login for this session and never saves it in the workbook."
        },
    ));
    ui.add_space(6.0);
    let mut user = s(d, "user");
    let mut password = s(d, "password");
    let mut show = b(d, "showPassword", false);
    egui::Grid::new("sql_login").num_columns(2).spacing(vec2(8.0, 6.0)).show(ui, |ui| {
        if integrated {
            return;
        }
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
    let mut remember = b(d, "remember", true);
    if !integrated && app.services.logins.is_some() {
        ui.checkbox(&mut remember, tr(app, "Remember on this computer"));
    }
    set(d, "remember", json!(remember));
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
            if (primary(ui, tr(app, "Connect")) || enter) && (integrated || !user.trim().is_empty()) {
                if integrated {
                    // Nothing to keep: the Windows session is the login.
                } else if let Err(e) = app.run("data.dataSourceSettings", json!({"server": server.trim(), "user": user.trim(), "password": password}))
                {
                    set(d, "error", json!(message(&e)));
                    return;
                } else if let Some(store) = app.services.logins.as_ref() {
                    // Kept in the system's credential store, or forgotten when unticked.
                    if remember {
                        if let Err(e) = store.save(server.trim(), user.trim(), &password) {
                            app.toast = Some((format!("{} {e}", tr(app, "The login couldn't be remembered:")), crate::now_ms()));
                        }
                    } else {
                        store.forget(server.trim());
                    }
                }
                set(d, "password", json!(""));
                let retry = s(d, "retry");
                if retry.is_empty() {
                    connect(app, d);
                } else {
                    *open = false;
                    let params = d.values.get("retryParams").cloned().unwrap_or(json!({}));
                    match retry.strip_prefix("dialog:") {
                        // A write dialog asked for the login: it comes back as it was.
                        Some(name) => app.open_dialog(name, params),
                        None => refresh_with_login(app, &retry, params),
                    }
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
        if let Some(store) = app.services.logins.as_ref() {
            store.forget(&server);
        }
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
    for c in pick.into_iter().filter(|c| c.source.auth() != gridcraft_engine::model::SqlAuth::Integrated) {
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
    for s in &servers {
        recall_login(app, s);
    }
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
                if matches!(c.command, gridcraft_engine::model::QueryCommand::Table { .. }) && ui.small_button(tr(app, "Save changes…")).clicked() {
                    app.open_dialog("saveToDatabase", json!({"table": c.table}));
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
        let last = match app.ext.query_timers.iter().find(|(u, n, _)| *u == uid && *n == name) {
            Some((_, _, t)) => *t,
            None => {
                app.ext.query_timers.push((uid, name.clone(), now));
                now
            }
        };
        recall_login(app, &server);
        let integrated = app
            .session
            .active()
            .is_some_and(|d| d.wb.connections.iter().any(|c| c.name == name && c.source.auth() == gridcraft_engine::model::SqlAuth::Integrated));
        let busy = app.dialog.is_some() || app.editor.is_some() || (!integrated && app.session.credential(&server).is_none());
        if now - last >= period && !busy {
            if let Err(e) = app.run("data.refresh", json!({"connection": name})) {
                app.toast = Some((format!("{name}: {}", message(&e)), now));
            }
            if let Some(t) = app.ext.query_timers.iter_mut().find(|(u, n, _)| *u == uid && *n == name) {
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

// ---------------------------------------------------------------- writing to the database

/// The queries' servers to write to, or a server typed in: the dialog's `target` params.
fn target_picker(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog) -> Json {
    let conns: Vec<(String, String, String)> = app
        .session
        .active()
        .map(|doc| doc.wb.connections.iter().map(|c| (c.name.clone(), c.source.server().to_string(), c.source.database().to_string())).collect())
        .unwrap_or_default();
    let mut pick = s(d, "connection");
    if pick.is_empty() && !b(d, "picked", false) {
        pick = conns.first().map(|c| c.0.clone()).unwrap_or_default();
    }
    ui.label(tr(app, "Database"));
    let shown =
        conns.iter().find(|c| c.0 == pick).map(|c| format!("{} · {} ({})", c.1, c.2, c.0)).unwrap_or_else(|| tr(app, "Another server…").to_string());
    egui::ComboBox::from_id_salt("db_target").width(380.0).selected_text(shown).show_ui(ui, |ui| {
        for c in &conns {
            if ui.selectable_label(pick == c.0, format!("{} · {} ({})", c.1, c.2, c.0)).clicked() {
                pick = c.0.clone();
            }
        }
        if ui.selectable_label(pick.is_empty(), tr(app, "Another server…")).clicked() {
            pick.clear();
        }
    });
    set(d, "picked", json!(true));
    set(d, "connection", json!(pick));
    if pick.is_empty() {
        let (mut server, mut database) = (s(d, "server"), s(d, "database"));
        let mut trust = b(d, "trust", false);
        egui::Grid::new("db_target_grid").num_columns(2).show(ui, |ui| {
            ui.label(tr(app, "Server"));
            ui.add(egui::TextEdit::singleline(&mut server).desired_width(280.0));
            ui.end_row();
            ui.label(tr(app, "Database"));
            ui.add(egui::TextEdit::singleline(&mut database).desired_width(280.0));
            ui.end_row();
        });
        ui.checkbox(&mut trust, tr(app, "Trust the server certificate"));
        auth_choice(app, ui, d);
        set(d, "server", json!(server));
        set(d, "database", json!(database));
        set(d, "trust", json!(trust));
        json!({"server": server.trim(), "database": database.trim(), "trustServerCertificate": trust, "auth": auth(d)})
    } else {
        json!({"connection": pick})
    }
}

/// The server a write dialog's params reach (for the login prompt).
fn server_of(app: &SheetApp, p: &Json) -> String {
    if let Some(n) = p.get("connection").and_then(Json::as_str) {
        return app
            .session
            .active()
            .and_then(|d| d.wb.connections.iter().find(|c| c.name == n).map(|c| c.source.server().to_string()))
            .unwrap_or_default();
    }
    p.get("server").and_then(Json::as_str).unwrap_or("").to_string()
}

/// Runs a write command; a missing or refused login opens the sign-in step, then this dialog
/// again. `Ok` carries the result.
fn write(app: &mut SheetApp, d: &mut Dialog, id: &str, p: Json, open: &mut bool) -> Option<Json> {
    let server = server_of(app, &p);
    recall_login(app, &server);
    match app.run(id, p.clone()) {
        Ok(v) => {
            set(d, "error", json!(""));
            Some(v)
        }
        Err(e) if e.contains(CREDENTIALS_REQUIRED) => {
            let server = server_of(app, &p);
            let user = app
                .session
                .active()
                .and_then(|doc| doc.wb.connections.iter().find(|c| c.source.server().eq_ignore_ascii_case(&server)).map(|c| c.user.clone()))
                .unwrap_or_default();
            let mut back = d.values.clone();
            back.insert("error".into(), json!(""));
            *open = false;
            app.dialog = Dialog::login(&server, &user, &format!("dialog:{}", d.name()), Json::Object(back));
            None
        }
        Err(e) => {
            set(d, "error", json!(message(&e)));
            None
        }
    }
}

/// Save Changes to Database: what will be sent, then Save.
pub fn save_to_database(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, open: &mut bool) {
    let table = s(d, "table");
    let p = if table.is_empty() { json!({}) } else { json!({"table": table}) };
    if d.values.get("plan").is_none_or(Json::is_null) && s(d, "error").is_empty() {
        let mut dry = p.clone();
        dry["dryRun"] = json!(true);
        if let Some(v) = write(app, d, "data.saveToDatabase", dry, open) {
            set(d, "plan", v);
        }
        if !*open {
            return;
        }
    }
    let plan = d.values.get("plan").cloned().unwrap_or(Json::Null);
    if !plan.is_null() {
        let n = |k: &str| plan[k].as_u64().unwrap_or(0);
        ui.label(RichText::new(plan["table"].as_str().unwrap_or("")).strong());
        egui::Grid::new("save_plan").num_columns(2).show(ui, |ui| {
            for (label, k) in [("Rows to add", "inserted"), ("Rows to change", "updated"), ("Rows to delete", "deleted")] {
                ui.label(tr(app, label));
                ui.label(RichText::new(n(k).to_string()).strong());
                ui.end_row();
            }
        });
        if n("inserted") + n("updated") + n("deleted") == 0 {
            ui.label(RichText::new(tr(app, "Nothing has changed since the last refresh.")).italics());
        } else {
            ui.label(tr(app, "Everything is sent in one transaction: all of it or nothing. Undo in GridCraft doesn't undo it on the server."));
        }
    }
    error_label(ui, d);
    ui.add_space(8.0);
    button_row(ui, |ui| {
        if secondary(ui, tr(app, "Cancel")) {
            *open = false;
        }
        let ready = plan["inserted"].as_u64().unwrap_or(0) + plan["updated"].as_u64().unwrap_or(0) + plan["deleted"].as_u64().unwrap_or(0) > 0;
        if ui.add_enabled_ui(ready, |ui| primary(ui, tr(app, "Save"))).inner
            && let Some(v) = write(app, d, "data.saveToDatabase", p.clone(), open)
        {
            *open = false;
            let total = v["inserted"].as_u64().unwrap_or(0) + v["updated"].as_u64().unwrap_or(0) + v["deleted"].as_u64().unwrap_or(0);
            app.toast = Some((format!("{total} {}", tr(app, "rows saved to the database.")), crate::now_ms()));
        }
    });
}

/// Export to SQL Server: target, table name, mode; Preview shows the plan, Export writes it.
pub fn export_to_sql_server(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, open: &mut bool) {
    let mut target = target_picker(app, ui, d);
    ui.add_space(4.0);
    let mut table = s(d, "table");
    ui.label(tr(app, "Table name (schema.table)"));
    ui.add(egui::TextEdit::singleline(&mut table).desired_width(380.0).hint_text("dbo.Sales"));
    let mut mode = d.values.get("mode").and_then(Json::as_str).unwrap_or("create").to_string();
    ui.horizontal(|ui| {
        ui.radio_value(&mut mode, "create".to_string(), tr(app, "New table"));
        ui.radio_value(&mut mode, "append".to_string(), tr(app, "Add rows"));
        ui.radio_value(&mut mode, "replace".to_string(), tr(app, "Replace all rows"));
    });
    let mut range = s(d, "range");
    ui.label(tr(app, "Range (blank: the table or region at the active cell)"));
    ui.add(egui::TextEdit::singleline(&mut range).desired_width(380.0));
    if s(d, "table") != table || d.values.get("mode").and_then(Json::as_str) != Some(mode.as_str()) || s(d, "range") != range {
        set(d, "plan", Json::Null);
    }
    set(d, "table", json!(table));
    set(d, "mode", json!(mode));
    set(d, "range", json!(range));
    target["table"] = json!(table.trim());
    target["mode"] = json!(mode);
    if !range.trim().is_empty() {
        target["range"] = json!(range.trim());
    }
    if let Some(plan) = d.values.get("plan").filter(|x| !x.is_null()).cloned() {
        ui.separator();
        ui.label(format!("{} {} → {}", plan["rows"], tr(app, "rows"), plan["table"].as_str().unwrap_or("")));
        let cols: Vec<String> = plan["columns"]
            .as_array()
            .map(|a| a.iter().map(|c| format!("{} ({})", c["name"].as_str().unwrap_or(""), c["kind"].as_str().unwrap_or(""))).collect())
            .unwrap_or_default();
        ui.label(RichText::new(cols.join(", ")).small());
        if mode == "replace" {
            ui.label(RichText::new(tr(app, "Every row now in the table is deleted first.")).color(Color32::from_rgb(0xC4, 0x2B, 0x1C)));
        }
    }
    error_label(ui, d);
    ui.add_space(8.0);
    button_row(ui, |ui| {
        if secondary(ui, tr(app, "Cancel")) {
            *open = false;
        }
        let planned = d.values.get("plan").is_some_and(|x| !x.is_null());
        if planned {
            if primary(ui, tr(app, "Export"))
                && let Some(v) = write(app, d, "data.exportToSqlServer", target.clone(), open)
            {
                *open = false;
                app.toast = Some((format!("{} {} {}", v["rows"], tr(app, "rows exported to"), v["table"].as_str().unwrap_or("")), crate::now_ms()));
            }
        } else if primary(ui, tr(app, "Preview")) {
            let mut dry = target.clone();
            dry["dryRun"] = json!(true);
            if let Some(v) = write(app, d, "data.exportToSqlServer", dry, open) {
                set(d, "plan", v);
            }
        }
    });
}

/// Run SQL: target, statement, optional cells for @P1…; asks before running.
pub fn run_sql(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, open: &mut bool) {
    let mut target = target_picker(app, ui, d);
    ui.add_space(4.0);
    let mut sql = s(d, "sql");
    ui.label(tr(app, "SQL statement"));
    ui.add(
        egui::TextEdit::multiline(&mut sql)
            .code_editor()
            .desired_rows(6)
            .desired_width(f32::INFINITY)
            .hint_text("UPDATE dbo.Prices SET Price = @P1 WHERE Code = @P2"),
    );
    let mut cells = s(d, "cells");
    ui.label(tr(app, "Cells for @P1, @P2… (optional, for example B2, C2)"));
    ui.add(egui::TextEdit::singleline(&mut cells).desired_width(380.0));
    if s(d, "sql") != sql || s(d, "cells") != cells {
        set(d, "confirm", json!(false));
    }
    set(d, "sql", json!(sql));
    set(d, "cells", json!(cells));
    target["sql"] = json!(sql);
    let refs: Vec<String> = cells.split([',', ';', ' ']).map(str::trim).filter(|x| !x.is_empty()).map(str::to_string).collect();
    if !refs.is_empty() {
        target["cells"] = json!(refs);
    }
    let confirming = b(d, "confirm", false);
    if confirming {
        ui.label(RichText::new(tr(app, "This runs on the server in one transaction and can't be undone from GridCraft. Run it?")).strong());
    }
    error_label(ui, d);
    ui.add_space(8.0);
    button_row(ui, |ui| {
        if secondary(ui, tr(app, "Cancel")) {
            *open = false;
        }
        if confirming {
            if primary(ui, tr(app, "Run"))
                && let Some(v) = write(app, d, "data.sqlExecute", target.clone(), open)
            {
                *open = false;
                app.toast = Some((format!("{} {}", v["rowsAffected"], tr(app, "rows affected.")), crate::now_ms()));
            }
        } else if ui.add_enabled_ui(!sql.trim().is_empty(), |ui| primary(ui, tr(app, "Run…"))).inner {
            set(d, "confirm", json!(true));
        }
    });
}

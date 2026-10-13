//! Get Data › From SQL Server Database end to end against a fake database: the source step,
//! sign-in, Navigator, Load, then Refresh asking for the login again, and the Queries &
//! Connections pane. `GRIDCRAFT_SNAPSHOT_DIR` also saves PNGs of the steps (needs a GPU adapter).

use std::sync::{Arc, Mutex};

use egui::vec2;
use egui_kittest::kittest::Queryable;
use gridcraft_engine::Session;
use gridcraft_engine::datasource::{Credential, DataProvider, DbColumn, DbError, DbObject, DbValue, QueryResult, Statement};
use gridcraft_engine::model::{DataSource, QueryCommand};
use gridcraft_ui_egui::{LoginStore, SheetApp};
use serde_json::{Value as Json, json};

type Harness = egui_kittest::Harness<'static, SheetApp>;

#[derive(Default)]
struct FakeDb {
    queries: Mutex<usize>,
    written: Mutex<Vec<Statement>>,
}

impl FakeDb {
    fn login(c: &Credential) -> Result<(), DbError> {
        if c.user == "reader" && c.password == "pw" { Ok(()) } else { Err(DbError::Credentials("Login failed for user.".into())) }
    }
}

impl DataProvider for FakeDb {
    fn databases(&self, _: &DataSource, c: &Credential) -> Result<Vec<String>, DbError> {
        FakeDb::login(c)?;
        Ok(vec!["Shop".into()])
    }
    fn objects(&self, _: &DataSource, c: &Credential) -> Result<Vec<DbObject>, DbError> {
        FakeDb::login(c)?;
        Ok(vec![
            DbObject { schema: "dbo".into(), name: "Customers".into(), view: false },
            DbObject { schema: "dbo".into(), name: "Orders".into(), view: false },
            DbObject { schema: "sales".into(), name: "Top Customers".into(), view: true },
        ])
    }
    fn columns(&self, _: &DataSource, c: &Credential, _: &str, _: &str) -> Result<Vec<DbColumn>, DbError> {
        FakeDb::login(c)?;
        let col = |n: &str, t: &str, pk: bool| DbColumn { name: n.into(), sql_type: t.into(), nullable: !pk, primary_key: pk, generated: false };
        Ok(vec![col("Id", "int", true), col("Name", "nvarchar", false), col("Since", "date", false), col("Active", "bit", false)])
    }
    fn execute(&self, _: &DataSource, c: &Credential, st: &[Statement]) -> Result<Vec<u64>, DbError> {
        FakeDb::login(c)?;
        self.written.lock().map_err(|_| DbError::Other("poisoned".into()))?.extend(st.iter().cloned());
        Ok(st.iter().map(|x| x.expect_rows.unwrap_or(1)).collect())
    }
    fn query(&self, _: &DataSource, c: &Credential, _: &QueryCommand, max: usize) -> Result<QueryResult, DbError> {
        FakeDb::login(c)?;
        *self.queries.lock().map_err(|_| DbError::Other("poisoned".into()))? += 1;
        let rows: Vec<Vec<DbValue>> = (1..=5)
            .map(|i| {
                vec![DbValue::Number(i as f64), DbValue::Text(format!("Customer {i}")), DbValue::Date(45_000.0 + i as f64), DbValue::Bool(i % 2 == 0)]
            })
            .collect();
        Ok(QueryResult {
            columns: vec!["Id".into(), "Name".into(), "Since".into(), "Active".into()],
            truncated: rows.len() > max,
            rows: rows.into_iter().take(max).collect(),
        })
    }
}

fn harness(db: Arc<FakeDb>) -> Harness {
    let mut session = Session::new();
    session.new_workbook();
    session.set_data_provider(db);
    let builder = egui_kittest::Harness::builder().with_step_dt(1.0 / 60.0).with_size(vec2(1200.0, 800.0));
    let builder = if std::env::var_os("GRIDCRAFT_SNAPSHOT_DIR").is_some() { builder.wgpu() } else { builder };
    let mut h = builder.build_ui_state(
        |ui, app: &mut SheetApp| {
            app.logic(ui.ctx());
            app.ui(ui);
        },
        SheetApp::new(session, Default::default()),
    );
    h.run_steps(4);
    h
}

/// Frames until the layout stops moving (a dialog step can change the window's size).
fn settle(h: &mut Harness) {
    h.run_steps(8);
}

fn snapshot(h: &mut Harness, name: &str) {
    let Some(dir) = std::env::var_os("GRIDCRAFT_SNAPSHOT_DIR") else { return };
    if let Ok(img) = h.render() {
        let _ = img.save(std::path::Path::new(&dir).join(format!("{name}.png")));
    }
}

fn stage(h: &Harness) -> String {
    h.state().dialog.as_ref().and_then(|d| d.values.get("stage")).and_then(Json::as_str).unwrap_or("").to_string()
}

fn set(h: &mut Harness, key: &str, v: Json) {
    if let Some(d) = h.state_mut().dialog.as_mut() {
        d.values.insert(key.into(), v);
    }
}

#[test]
fn load_a_table_then_refresh_it() {
    let db = Arc::new(FakeDb::default());
    let mut h = harness(db.clone());
    h.state_mut().open_dialog("sqlServer", json!({}));
    settle(&mut h);
    assert_eq!(stage(&h), "source");
    set(&mut h, "server", json!("db.example"));
    settle(&mut h);
    snapshot(&mut h, "1-source");
    h.get_by_label("OK").click();
    settle(&mut h);
    // No login is kept for the server yet: sign in.
    assert_eq!(stage(&h), "credentials");
    set(&mut h, "user", json!("reader"));
    set(&mut h, "password", json!("wrong"));
    settle(&mut h);
    h.get_by_label("Connect").click();
    settle(&mut h);
    assert_eq!(stage(&h), "credentials", "a refused login asks again");
    set(&mut h, "password", json!("pw"));
    settle(&mut h);
    snapshot(&mut h, "2-credentials");
    h.get_by_label("Connect").click();
    settle(&mut h);
    // No database named: pick one in the Navigator.
    assert_eq!(stage(&h), "navigator");
    set(&mut h, "database", json!("Shop"));
    h.state_mut().dialog.as_mut().map(|d| d.values.insert("objects".into(), json!([])));
    settle(&mut h);
    // Re-list for the chosen database, as picking it in the list does.
    let objects = h.state_mut().run("data.sqlServerNavigator", json!({"server": "db.example", "database": "Shop"})).unwrap();
    set(&mut h, "objects", objects["objects"].clone());
    settle(&mut h);
    h.get_by_label("▦ dbo.Customers").click();
    settle(&mut h);
    let preview = h.state().dialog.as_ref().and_then(|d| d.values.get("preview").cloned()).unwrap_or(Json::Null);
    assert_eq!(preview["rows"][0][1], "Customer 1", "{preview}");
    snapshot(&mut h, "3-navigator");
    h.get_by_label("Load").click();
    settle(&mut h);
    assert!(h.state().dialog.is_none(), "Load closes the dialog: {:?}", h.state().dialog.as_ref().map(|d| d.values.clone()));
    {
        let d = h.state().session.active().unwrap();
        assert_eq!(d.wb.connections.len(), 1);
        assert_eq!(d.wb.active().unwrap().name, "dbo Customers");
        assert_eq!(d.wb.active().unwrap().tables[0].range.a1(), "A1:D6");
    }
    snapshot(&mut h, "4-loaded");

    // A new session has no login: Refresh All asks for it, then refreshes.
    h.state_mut().session.clear_credential("db.example");
    let before = *db.queries.lock().unwrap();
    h.state_mut().run_or_alert("data.refreshAll", json!({}));
    settle(&mut h);
    assert_eq!(stage(&h), "credentials");
    assert_eq!(h.state().dialog.as_ref().unwrap().values["user"], "reader", "the query's user is filled in");
    set(&mut h, "password", json!("pw"));
    settle(&mut h);
    h.get_by_label("Connect").click();
    settle(&mut h);
    assert!(h.state().dialog.is_none());
    assert_eq!(*db.queries.lock().unwrap(), before + 1);

    // The Queries & Connections pane lists the query and refreshes it.
    h.state_mut().grid.pane = Some("queries".into());
    settle(&mut h);
    snapshot(&mut h, "5-queries-pane");
    h.get_by_label("Refresh").click();
    settle(&mut h);
    assert_eq!(*db.queries.lock().unwrap(), before + 2);
}

#[test]
fn a_sql_statement_shows_its_preview() {
    let db = Arc::new(FakeDb::default());
    let mut h = harness(db);
    h.state_mut().session.set_credential("db.example", Credential { user: "reader".into(), password: "pw".into() });
    h.state_mut().open_dialog("sqlServer", json!({"server": "db.example", "database": "Shop"}));
    settle(&mut h);
    set(&mut h, "sql", json!("SELECT * FROM dbo.Customers WHERE Active = 1"));
    settle(&mut h);
    h.get_by_label("OK").click();
    settle(&mut h);
    assert_eq!(stage(&h), "preview");
    snapshot(&mut h, "6-statement-preview");
    h.get_by_label("Load").click();
    settle(&mut h);
    let d = h.state().session.active().unwrap();
    assert_eq!(d.wb.connections[0].name, "Query1");
    assert!(matches!(d.wb.connections[0].command, QueryCommand::Sql { .. }));
}

/// A server whose certificate has expired: refused until the user trusts it.
struct ExpiredCert;

impl DataProvider for ExpiredCert {
    fn databases(&self, s: &DataSource, c: &Credential) -> Result<Vec<String>, DbError> {
        let DataSource::SqlServer { trust_server_certificate, .. } = s;
        if !trust_server_certificate {
            return Err(DbError::Encryption("The server's certificate expired on 2026-03-11.".into()));
        }
        FakeDb::login(c)?;
        Ok(vec!["Shop".into()])
    }
    fn objects(&self, s: &DataSource, c: &Credential) -> Result<Vec<DbObject>, DbError> {
        self.databases(s, c).map(|_| vec![])
    }
    fn query(&self, s: &DataSource, c: &Credential, _: &QueryCommand, _: usize) -> Result<QueryResult, DbError> {
        self.databases(s, c).map(|_| QueryResult::default())
    }
}

#[test]
fn an_expired_certificate_can_be_trusted() {
    let mut session = Session::new();
    session.new_workbook();
    session.set_data_provider(Arc::new(ExpiredCert));
    session.set_credential("old.example", Credential { user: "reader".into(), password: "pw".into() });
    let mut h = egui_kittest::Harness::builder().with_step_dt(1.0 / 60.0).with_size(vec2(1200.0, 800.0)).build_ui_state(
        |ui, app: &mut SheetApp| {
            app.logic(ui.ctx());
            app.ui(ui);
        },
        SheetApp::new(session, Default::default()),
    );
    h.state_mut().open_dialog("sqlServer", json!({"server": "old.example"}));
    settle(&mut h);
    h.get_by_label("OK").click();
    settle(&mut h);
    let err = h.state().dialog.as_ref().and_then(|d| d.values.get("error").cloned()).unwrap_or(Json::Null);
    assert!(err.as_str().unwrap_or("").contains("expired on 2026-03-11"), "{err}");
    h.get_by_label("Trust the certificate and connect").click();
    settle(&mut h);
    assert_eq!(stage(&h), "navigator");
    assert_eq!(h.state().dialog.as_ref().unwrap().values["trust"], true);
}

#[test]
fn write_dialogs_preview_then_write() {
    let db = Arc::new(FakeDb::default());
    let mut h = harness(db.clone());
    h.state_mut().session.set_credential("db.example", Credential { user: "reader".into(), password: "pw".into() });
    h.state_mut().run("data.fromSqlServer", json!({"server": "db.example", "database": "Shop", "table": "dbo.Customers"})).unwrap();
    h.state_mut().run("cell.set", json!({"cell": "B3", "input": "Renamed"})).unwrap();
    h.state_mut().run("selection.set", json!({"range": "A2"})).unwrap();
    // Save Changes: the plan first, then Save.
    h.state_mut().open_dialog("saveToDatabase", json!({}));
    settle(&mut h);
    let plan = h.state().dialog.as_ref().unwrap().values["plan"].clone();
    assert_eq!(plan["updated"], 1, "{plan}");
    snapshot(&mut h, "7-save-changes");
    assert!(db.written.lock().unwrap().is_empty(), "the plan writes nothing");
    h.get_by_label("Save").click();
    settle(&mut h);
    assert!(h.state().dialog.is_none());
    assert!(db.written.lock().unwrap().iter().any(|x| x.sql.starts_with("UPDATE")));

    // Export: Preview, then Export.
    db.written.lock().unwrap().clear();
    h.state_mut().open_dialog("exportToSqlServer", json!({"table": "dbo.Copy"}));
    settle(&mut h);
    h.get_by_label("Preview").click();
    settle(&mut h);
    let plan = h.state().dialog.as_ref().unwrap().values["plan"].clone();
    assert_eq!(plan["rows"], 5, "{plan}");
    snapshot(&mut h, "8-export");
    h.get_by_label("Export").click();
    settle(&mut h);
    assert!(h.state().dialog.is_none());
    assert!(db.written.lock().unwrap()[0].sql.starts_with("CREATE TABLE [dbo].[Copy]"));

    // Run SQL asks before running.
    db.written.lock().unwrap().clear();
    h.state_mut().open_dialog("sqlExecute", json!({"sql": "EXEC dbo.CloseMonth"}));
    settle(&mut h);
    h.get_by_label("Run…").click();
    settle(&mut h);
    assert!(db.written.lock().unwrap().is_empty());
    snapshot(&mut h, "9-run-sql");
    h.get_by_label("Run").click();
    settle(&mut h);
    assert_eq!(db.written.lock().unwrap()[0].sql, "EXEC dbo.CloseMonth");
}

#[test]
fn a_write_without_a_login_asks_for_it_and_comes_back() {
    let db = Arc::new(FakeDb::default());
    let mut h = harness(db.clone());
    h.state_mut().open_dialog("sqlExecute", json!({"sql": "EXEC x", "connection": "", "picked": true, "server": "db.example", "database": "Shop"}));
    settle(&mut h);
    h.get_by_label("Run…").click();
    settle(&mut h);
    h.get_by_label("Run").click();
    settle(&mut h);
    assert_eq!(stage(&h), "credentials");
    set(&mut h, "user", json!("reader"));
    set(&mut h, "password", json!("pw"));
    settle(&mut h);
    h.get_by_label("Connect").click();
    settle(&mut h);
    let d = h.state().dialog.as_ref().expect("Run SQL comes back");
    assert_eq!(d.name(), "sqlExecute");
    assert_eq!(d.values["sql"], "EXEC x");
}

/// A credential store in memory.
#[derive(Clone, Default)]
struct MemLogins(Arc<Mutex<Vec<(String, String, String)>>>);

impl LoginStore for MemLogins {
    fn load(&self, server: &str) -> Option<(String, String)> {
        self.0.lock().ok()?.iter().find(|x| x.0 == server).map(|x| (x.1.clone(), x.2.clone()))
    }
    fn save(&self, server: &str, user: &str, password: &str) -> Result<(), String> {
        let mut v = self.0.lock().map_err(|_| "poisoned".to_string())?;
        v.retain(|x| x.0 != server);
        v.push((server.into(), user.into(), password.into()));
        Ok(())
    }
    fn forget(&self, server: &str) {
        if let Ok(mut v) = self.0.lock() {
            v.retain(|x| x.0 != server);
        }
    }
}

#[test]
fn a_remembered_login_is_used_after_a_restart() {
    let db = Arc::new(FakeDb::default());
    let logins = MemLogins::default();
    let mut h = harness(db.clone());
    h.state_mut().services.logins = Some(Box::new(logins.clone()));
    h.state_mut().open_dialog("sqlServer", json!({"server": "db.example", "database": "Shop"}));
    settle(&mut h);
    h.get_by_label("OK").click();
    settle(&mut h);
    assert_eq!(stage(&h), "credentials");
    set(&mut h, "user", json!("reader"));
    set(&mut h, "password", json!("pw"));
    settle(&mut h);
    h.get_by_label("Connect").click();
    settle(&mut h);
    assert_eq!(logins.load("db.example"), Some(("reader".into(), "pw".into())), "Remember is on by default");
    h.state_mut().dialog = None;
    h.state_mut().run("data.fromSqlServer", json!({"server": "db.example", "database": "Shop", "table": "dbo.Customers"})).unwrap();
    // A new session (the app restarted): Refresh All finds the login without asking.
    h.state_mut().session.clear_credential("db.example");
    let before = *db.queries.lock().unwrap();
    h.state_mut().run_or_alert("data.refreshAll", json!({}));
    settle(&mut h);
    assert!(h.state().dialog.is_none(), "no sign-in asked");
    assert_eq!(*db.queries.lock().unwrap(), before + 1);
    // Clear Permissions forgets it on this computer too.
    h.state_mut().open_dialog("dataSourceSettings", json!({}));
    settle(&mut h);
    h.get_by_label("Clear Permissions").click();
    settle(&mut h);
    assert_eq!(logins.load("db.example"), None);
}

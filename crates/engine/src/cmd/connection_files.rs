//! Connections saved on this computer: Office Data Connection files (`.odc`) in
//! `Documents/My Data Sources`, as Excel keeps them. Export Connection File saves a query's
//! connection there, Existing Connections lists them and Open loads one as a new query, so a
//! server, database and table (or statement) set up once can be used from any workbook.
//!
//! A connection file never holds a password: opening one asks for the login as Get Data does.

use std::path::{Path, PathBuf};

use gridcraft_model::{QueryCommand, SqlAuth};
use gridcraft_xlsx::odc::{ConnectionFile, MAX_ODC_BYTES, odc_file_name, read_odc, write_odc};
use serde_json::{Value as Json, json};

use super::connections::{find_connection, from_sql_server, object_label};
use super::{CommandSpec, EngineError, Result, Session, always, bad, bool_param, cmd, has_doc, str_param};

/// At most this many files are listed from a folder.
const MAX_LISTED: usize = 1000;
/// Longer paths aren't paths.
const MAX_PATH: usize = 4096;

pub fn specs() -> Vec<CommandSpec> {
    vec![
        cmd!(
            noundo "data.exportConnectionFile",
            "Export Connection File",
            ["Data", "Queries & Connections", "Properties"],
            None,
            "{connection? | table? (default: the query of the table at the active cell), path? (default: Documents/My Data Sources/<server database object>.odc), overwrite?: false} saves the connection (never a password) → {path}",
            has_doc,
            export
        ),
        cmd!(
            query "data.existingConnections",
            "Existing Connections",
            ["Data", "Get & Transform Data", "Existing Connections"],
            None,
            "{folder? (default: Documents/My Data Sources) | path?: one file} → {folder, connections: [{name, description, path, server, database, auth, user, table | sql}]} (files that aren't SQL Server connections are left out)",
            always,
            existing
        ),
        cmd!(
            "data.fromConnectionFile",
            "Open Connection File",
            [],
            None,
            "{path: .odc file, user?, password? (kept for the session, never saved), name?: query name} loads it into a new sheet as data.fromSqlServer does → {loaded: [...]}",
            has_doc,
            open
        ),
    ]
}

/// `Documents/My Data Sources` in the user's home folder; `None` without a home (the web).
pub fn data_sources_dir() -> Option<PathBuf> {
    if cfg!(target_arch = "wasm32") {
        return None;
    }
    let home = std::env::var_os("USERPROFILE").filter(|_| cfg!(windows)).or_else(|| std::env::var_os("HOME"))?;
    Some(PathBuf::from(home).join("Documents").join("My Data Sources"))
}

fn path_param(cmd: &str, p: &Json, key: &str) -> Result<Option<PathBuf>> {
    match str_param(p, key).map(str::trim).filter(|v| !v.is_empty()) {
        Some(v) if v.len() > MAX_PATH || v.contains('\0') => Err(bad(cmd, format!("`{key}` is not a path"))),
        Some(v) => Ok(Some(PathBuf::from(v))),
        None => Ok(None),
    }
}

/// `dir/name`, or `dir/name (2).odc`… when that file exists: a saved connection is never
/// replaced unless asked.
fn free_path(dir: &Path, name: &str) -> PathBuf {
    let first = dir.join(name);
    if !first.exists() {
        return first;
    }
    let stem = name.strip_suffix(".odc").unwrap_or(name);
    (2..1000).map(|k| dir.join(format!("{stem} ({k}).odc"))).find(|p| !p.exists()).unwrap_or(first)
}

fn path_text(path: &Path) -> String {
    path.to_string_lossy().into_owned()
}

fn export(s: &mut Session, p: &Json) -> Result<Json> {
    const C: &str = "data.exportConnectionFile";
    let ci = find_connection(s, p, C)?;
    let file = s.doc()?.wb.connections.get(ci).map(ConnectionFile::of).ok_or_else(|| bad(C, "no such query"))?;
    let path = match path_param(C, p, "path")? {
        Some(path) => path,
        None => {
            let dir = data_sources_dir().ok_or_else(|| EngineError::Other("There is no Documents folder to save connections in here.".into()))?;
            crate::io::script_sandbox_check(&path_text(&dir))?;
            std::fs::create_dir_all(&dir).map_err(|e| EngineError::Other(format!("{}: {e}", dir.display())))?;
            free_path(&dir, &odc_file_name(&file))
        }
    };
    let text = path_text(&path);
    if !bool_param(p, "overwrite").unwrap_or(false) && path.exists() {
        return Err(EngineError::Other(format!("{text} already exists. Pass `overwrite: true` to replace it.")));
    }
    crate::io::write_file(&text, write_odc(&file).as_bytes())?;
    Ok(json!({"path": text}))
}

/// What a listing says about one file.
fn file_json(path: &Path, f: &ConnectionFile) -> Json {
    let auth = match f.source.auth() {
        SqlAuth::Sql => "sql",
        SqlAuth::Windows => "windows",
        SqlAuth::Integrated => "integrated",
    };
    let mut v = json!({
        "name": f.name,
        "description": f.description,
        "path": path_text(path),
        "server": f.source.server(),
        "database": f.source.database(),
        "auth": auth,
        "user": f.user,
    });
    match &f.command {
        QueryCommand::Table { schema, name } => v["table"] = json!(object_label(schema, name)),
        QueryCommand::Sql { text } => v["sql"] = json!(text),
    }
    v
}

/// The connection file at `path`, when it is one GridCraft can open.
fn read_connection_file(path: &Path) -> Option<ConnectionFile> {
    let meta = std::fs::metadata(path).ok()?;
    if !meta.is_file() || meta.len() > MAX_ODC_BYTES as u64 {
        return None;
    }
    read_odc(&crate::io::read_file(&path_text(path)).ok()?)
}

fn existing(_s: &mut Session, p: &Json) -> Result<Json> {
    const C: &str = "data.existingConnections";
    if let Some(path) = path_param(C, p, "path")? {
        crate::io::script_sandbox_check(&path_text(&path))?;
        let one: Vec<Json> = read_connection_file(&path).map(|f| file_json(&path, &f)).into_iter().collect();
        let folder = path.parent().map(path_text).unwrap_or_default();
        return Ok(json!({"folder": folder, "connections": one}));
    }
    let Some(folder) = path_param(C, p, "folder")?.or_else(data_sources_dir) else {
        return Ok(json!({"folder": "", "connections": []}));
    };
    let shown = path_text(&folder);
    crate::io::script_sandbox_check(&shown)?;
    // A folder that isn't there yet holds no connections.
    let Ok(entries) = std::fs::read_dir(&folder) else { return Ok(json!({"folder": shown, "connections": []})) };
    let mut paths: Vec<PathBuf> = entries
        .filter_map(|e| e.ok().map(|e| e.path()))
        .filter(|path| path.extension().is_some_and(|x| x.eq_ignore_ascii_case("odc")))
        .take(MAX_LISTED)
        .collect();
    paths.sort();
    let mut list: Vec<(String, Json)> =
        paths.iter().filter_map(|path| read_connection_file(path).map(|f| (f.name.to_lowercase(), file_json(path, &f)))).collect();
    list.sort_by(|a, b| a.0.cmp(&b.0));
    Ok(json!({"folder": shown, "connections": list.into_iter().map(|(_, v)| v).collect::<Vec<_>>()}))
}

fn open(s: &mut Session, p: &Json) -> Result<Json> {
    const C: &str = "data.fromConnectionFile";
    let path = path_param(C, p, "path")?.ok_or_else(|| bad(C, "missing `path` (see data.existingConnections)"))?;
    let text = path_text(&path);
    let bytes = crate::io::read_file(&text)?;
    let f = read_odc(&bytes).ok_or_else(|| EngineError::Other(format!("{text} isn't a SQL Server connection file GridCraft can open.")))?;
    let mut params = file_json(&path, &f);
    params["encrypt"] = json!(f.source.encrypted());
    params["trustServerCertificate"] = json!(f.source.trusts_server_certificate());
    params["name"] = json!(str_param(p, "name").map(str::trim).filter(|n| !n.is_empty()).unwrap_or(&f.name));
    if let Some(o) = params.as_object_mut() {
        o.remove("path");
        o.remove("description");
        o.remove("user");
    }
    // The login comes from the call, else the session (from_sql_server asks for it).
    for key in ["user", "password"] {
        if let Some(v) = str_param(p, key) {
            params[key] = json!(v);
        }
    }
    from_sql_server(s, &params)
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use super::*;
    use crate::datasource::{Credential, DataProvider, DbError, DbObject, DbValue, QueryResult};
    use gridcraft_model::DataSource;

    /// Answers every query with the same two rows.
    struct Fake;

    impl DataProvider for Fake {
        fn databases(&self, _: &DataSource, _: &Credential) -> std::result::Result<Vec<String>, DbError> {
            Ok(vec!["Shop".into()])
        }
        fn objects(&self, _: &DataSource, _: &Credential) -> std::result::Result<Vec<DbObject>, DbError> {
            Ok(vec![])
        }
        fn query(&self, _: &DataSource, c: &Credential, _: &QueryCommand, _: usize) -> std::result::Result<QueryResult, DbError> {
            if c.password != "secret" {
                return Err(DbError::Credentials("Login failed.".into()));
            }
            Ok(QueryResult {
                columns: vec!["Id".into(), "Name".into()],
                rows: vec![vec![DbValue::Number(1.0), DbValue::Text("Ana".into())], vec![DbValue::Number(2.0), DbValue::Text("Luis".into())]],
                truncated: false,
            })
        }
    }

    fn session() -> Session {
        let mut s = Session::new();
        s.execute("file.new", json!({})).unwrap();
        s.set_data_provider(Arc::new(Fake));
        s
    }

    fn temp_dir(tag: &str) -> PathBuf {
        let dir = std::env::temp_dir().join(format!("gridcraft-odc-{tag}-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();
        dir
    }

    #[test]
    fn a_saved_connection_opens_in_another_workbook() {
        let dir = temp_dir("roundtrip");
        let mut s = session();
        s.execute(
            "data.fromSqlServer",
            json!({"server": "db.example", "database": "Shop", "table": "dbo.Customers", "user": "reader", "password": "secret"}),
        )
        .unwrap();
        let path = dir.join("customers.odc");
        let r = s.execute("data.exportConnectionFile", json!({"path": path.to_string_lossy()})).unwrap();
        let saved = std::fs::read_to_string(&path).unwrap();
        assert!(r["path"].as_str().is_some_and(|p| p.ends_with("customers.odc")));
        assert!(!saved.contains("secret"), "a connection file never holds the password");
        // Saving again over it needs `overwrite`.
        assert!(s.execute("data.exportConnectionFile", json!({"path": path.to_string_lossy()})).is_err());
        s.execute("data.exportConnectionFile", json!({"path": path.to_string_lossy(), "overwrite": true})).unwrap();

        let list = s.execute("data.existingConnections", json!({"folder": dir.to_string_lossy()})).unwrap();
        let conns = list["connections"].as_array().unwrap();
        assert_eq!(conns.len(), 1);
        assert_eq!(
            (conns[0]["server"].as_str(), conns[0]["table"].as_str(), conns[0]["user"].as_str()),
            (Some("db.example"), Some("dbo.Customers"), Some("reader"))
        );

        // A new session (another day): the login is asked for, then the query loads.
        let mut other = session();
        let err = other.execute("data.fromConnectionFile", json!({"path": path.to_string_lossy()})).unwrap_err();
        assert!(err.to_string().contains(crate::datasource::CREDENTIALS_REQUIRED));
        let r = other.execute("data.fromConnectionFile", json!({"path": path.to_string_lossy(), "user": "reader", "password": "secret"})).unwrap();
        assert_eq!(r["loaded"][0]["rows"], json!(2));
        let wb = &other.doc().unwrap().wb;
        assert_eq!(wb.connections.len(), 1);
        assert_eq!(wb.connections[0].command, QueryCommand::Table { schema: "dbo".into(), name: "Customers".into() });
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn listings_skip_what_isnt_a_connection() {
        let dir = temp_dir("listing");
        std::fs::write(dir.join("notes.odc"), b"not a connection").unwrap();
        std::fs::write(dir.join("readme.txt"), b"hello").unwrap();
        std::fs::create_dir_all(dir.join("sub.odc")).unwrap();
        let mut s = session();
        let list = s.execute("data.existingConnections", json!({"folder": dir.to_string_lossy()})).unwrap();
        assert_eq!(list["connections"], json!([]));
        let missing = s.execute("data.existingConnections", json!({"folder": dir.join("nope").to_string_lossy()})).unwrap();
        assert_eq!(missing["connections"], json!([]));
        assert!(s.execute("data.fromConnectionFile", json!({"path": dir.join("notes.odc").to_string_lossy()})).is_err());
        assert!(s.execute("data.fromConnectionFile", json!({})).is_err());
        assert!(s.execute("data.existingConnections", json!({"folder": "x".repeat(MAX_PATH + 1)})).is_err());
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn default_names_never_replace_a_file() {
        let dir = temp_dir("names");
        assert_eq!(free_path(&dir, "a.odc"), dir.join("a.odc"));
        std::fs::write(dir.join("a.odc"), b"x").unwrap();
        std::fs::write(dir.join("a (2).odc"), b"x").unwrap();
        assert_eq!(free_path(&dir, "a.odc"), dir.join("a (3).odc"));
        let mut s = session();
        let one = s.execute("data.existingConnections", json!({"path": dir.join("a.odc").to_string_lossy()})).unwrap();
        assert_eq!(one["connections"], json!([]));
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn export_needs_a_query() {
        let mut s = session();
        assert!(s.execute("data.exportConnectionFile", json!({"path": "/nonexistent/x.odc"})).is_err());
    }
}

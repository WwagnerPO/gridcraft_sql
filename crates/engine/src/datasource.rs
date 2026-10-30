//! External data sources (Data › Get Data › From Database).
//!
//! The engine never opens a network connection itself: it compiles to wasm, where there are no
//! sockets. A host that can reach databases (the desktop app, the CLI) installs a
//! [`DataProvider`] with [`crate::Session::set_data_provider`]; without one, the database commands
//! say so instead of failing silently.
//!
//! Credentials live in the session ([`crate::Session::set_credential`]), never in the workbook or
//! the journal.

use std::sync::Arc;

use gridcraft_model::{DataSource, QueryCommand};

/// A SQL Server login. `Debug` never shows the password.
#[derive(Clone, PartialEq, Eq)]
pub struct Credential {
    pub user: String,
    pub password: String,
}

impl std::fmt::Debug for Credential {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Credential").field("user", &self.user).finish_non_exhaustive()
    }
}

/// What a session keeps for the database commands. Memory only: never saved in a workbook or
/// the journal.
#[derive(Default)]
pub struct SessionData {
    /// Reaches databases for Get Data › From Database; `None` where the host can't (the web).
    provider: Option<Arc<dyn DataProvider>>,
    /// Database logins entered this session, by [`credential_key`].
    credentials: Vec<(String, Credential)>,
    /// What each query last loaded, for Save to Database.
    pub(crate) snapshots: Vec<crate::cmd::dbwrite::Snapshot>,
    /// Documents (uid) whose VBA macros the user enabled this session (Enable Content); see
    /// [`crate::vba_host::macros_allowed`].
    pub(crate) macros_trusted: Vec<u64>,
    /// A macro is running: the changes it makes raise no events.
    pub(crate) in_macro: bool,
    /// The app runs events itself (see [`crate::vba_host::events`]).
    pub(crate) events_deferred: bool,
    pub(crate) pending_events: Vec<crate::vba_host::events::PendingEvent>,
}

impl crate::Session {
    /// Installs the provider the database commands reach servers through.
    pub fn set_data_provider(&mut self, provider: Arc<dyn DataProvider>) {
        self.data.provider = Some(provider);
    }
    pub fn data_provider(&self) -> Option<Arc<dyn DataProvider>> {
        self.data.provider.clone()
    }
    /// Remembers the login for `server` (Data Source Settings), replacing an earlier one.
    pub fn set_credential(&mut self, server: &str, credential: Credential) {
        let key = credential_key(server);
        self.data.credentials.retain(|(k, _)| *k != key);
        self.data.credentials.push((key, credential));
    }
    pub fn credential(&self, server: &str) -> Option<&Credential> {
        let key = credential_key(server);
        self.data.credentials.iter().find(|(k, _)| *k == key).map(|(_, c)| c)
    }
    /// Forgets the login for `server`; true if there was one.
    pub fn clear_credential(&mut self, server: &str) -> bool {
        let key = credential_key(server);
        let n = self.data.credentials.len();
        self.data.credentials.retain(|(k, _)| *k != key);
        n != self.data.credentials.len()
    }
    /// The servers with a remembered login, and its user name.
    pub fn credential_servers(&self) -> Vec<(String, String)> {
        self.data.credentials.iter().map(|(k, c)| (k.clone(), c.user.clone())).collect()
    }
}

/// A table or view a database offers (the Navigator's list).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DbObject {
    pub schema: String,
    pub name: String,
    /// A view rather than a table.
    pub view: bool,
}

/// One value of a result set. Dates and times are serial numbers in the 1900 date system.
#[derive(Clone, Debug, PartialEq)]
pub enum DbValue {
    Null,
    Bool(bool),
    Number(f64),
    Text(String),
    Date(f64),
    DateTime(f64),
    Time(f64),
}

/// The rows a query returned.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct QueryResult {
    pub columns: Vec<String>,
    pub rows: Vec<Vec<DbValue>>,
    /// More rows were available than `max_rows`.
    pub truncated: bool,
}

/// A statement to run with parameters: `@P1`, `@P2`… take `params` in order.
#[derive(Clone, Debug, PartialEq)]
pub struct Statement {
    pub sql: String,
    pub params: Vec<DbValue>,
    /// The rows it must affect, else the whole batch is rolled back (a row changed or deleted on
    /// the server since it was loaded).
    pub expect_rows: Option<u64>,
}

impl Statement {
    pub fn new(sql: impl Into<String>, params: Vec<DbValue>) -> Statement {
        Statement { sql: sql.into(), params, expect_rows: None }
    }
}

/// A column of a database table.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DbColumn {
    pub name: String,
    /// The type as SQL Server names it (`int`, `nvarchar`, `datetime2`…).
    pub sql_type: String,
    pub nullable: bool,
    pub primary_key: bool,
    /// Filled by the server (identity or computed): never written.
    pub generated: bool,
}

/// Errors a provider reports. [`DbError::Credentials`] makes the interface ask for a login.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DbError {
    /// The login was refused (or none was given).
    Credentials(String),
    /// The encrypted connection could not be set up (no TLS on the server, or an untrusted
    /// certificate).
    Encryption(String),
    /// Anything else: unreachable server, bad SQL, timeout…
    Other(String),
}

impl std::fmt::Display for DbError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            DbError::Credentials(m) => write!(f, "{CREDENTIALS_REQUIRED} {m}"),
            DbError::Encryption(m) => write!(f, "{ENCRYPTION_FAILED} {m}"),
            DbError::Other(m) => f.write_str(m),
        }
    }
}

/// Errors that need a login start with this; the interface then asks for one.
pub const CREDENTIALS_REQUIRED: &str = "Sign in to the database server to continue.";
/// Errors where encryption failed start with this; the interface offers to connect unencrypted.
pub const ENCRYPTION_FAILED: &str = "An encrypted connection to the server could not be made.";
/// The error when no provider is installed (the web version).
pub const NOT_AVAILABLE: &str = "Database connections aren't available in this version of GridCraft.";

/// Most rows a query loads: a worksheet holds 1,048,576 rows, one of them the header.
pub const MAX_ROWS: usize = 1_048_575;
/// Most cells a query loads, so a huge result can't exhaust memory.
pub const MAX_CELLS: usize = 50_000_000;
/// Longest text a cell holds.
pub const MAX_TEXT: usize = 32_767;

/// Reaches databases for the engine. Calls block until the server answers or times out.
pub trait DataProvider: Send + Sync {
    /// The databases the login can open.
    fn databases(&self, source: &DataSource, credential: &Credential) -> Result<Vec<String>, DbError>;
    /// The tables and views of the source's database.
    fn objects(&self, source: &DataSource, credential: &Credential) -> Result<Vec<DbObject>, DbError>;
    /// Runs `command`, returning at most `max_rows` rows (and at most [`MAX_CELLS`] cells).
    fn query(&self, source: &DataSource, credential: &Credential, command: &QueryCommand, max_rows: usize) -> Result<QueryResult, DbError>;
    /// The columns of a table, primary key marked.
    fn columns(&self, _source: &DataSource, _credential: &Credential, _schema: &str, _name: &str) -> Result<Vec<DbColumn>, DbError> {
        Err(DbError::Other(WRITE_NOT_SUPPORTED.into()))
    }
    /// Runs `statements` in one transaction: all of them or none. Returns the rows each affected.
    fn execute(&self, _source: &DataSource, _credential: &Credential, _statements: &[Statement]) -> Result<Vec<u64>, DbError> {
        Err(DbError::Other(WRITE_NOT_SUPPORTED.into()))
    }
}

/// The error of a provider that only reads.
pub const WRITE_NOT_SUPPORTED: &str = "This connection can only read data.";

/// A serial number as SQL Server reads it in a string: `2023-03-15T12:00:00.000` (date-times),
/// `2023-03-15` (dates), `12:00:00.000` (times). `None` for numbers that aren't dates.
pub fn serial_to_iso(serial: f64, value: &DbValue) -> Option<String> {
    if !serial.is_finite() || serial < 0.0 {
        return None;
    }
    let whole = serial.floor();
    let ms = ((serial - whole) * 86_400_000.0).round() as i64;
    let (h, m, s, f) = (ms / 3_600_000, ms / 60_000 % 60, ms / 1000 % 60, ms % 1000);
    let time = format!("{h:02}:{m:02}:{s:02}.{f:03}");
    if matches!(value, DbValue::Time(_)) {
        return Some(time);
    }
    // Serial 60 is the fictitious 29 February 1900; 61 is 1 March.
    let mut d = whole as i64;
    if d < 1 {
        return None;
    }
    if d >= 61 {
        d -= 1;
    } else if d == 60 {
        return None;
    }
    // Days since 1899-12-31 → civil date.
    let z = d - 25_568 + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = if mp < 10 { mp + 3 } else { mp - 9 };
    let year = yoe + era * 400 + i64::from(month <= 2);
    let date = format!("{year:04}-{month:02}-{day:02}");
    Some(if matches!(value, DbValue::Date(_)) { date } else { format!("{date}T{time}") })
}

/// The key credentials are kept under: the server name as typed, trimmed and lower-cased.
pub fn credential_key(server: &str) -> String {
    server.trim().to_lowercase()
}

/// `dbo.Customers`, `[dbo].[Order Details]`, `Customers` → (schema, name). Brackets and double
/// quotes delimit names that contain dots or spaces.
pub fn parse_object_name(text: &str) -> Option<(String, String)> {
    let mut parts: Vec<String> = Vec::new();
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
    match parts.as_slice() {
        [name] if !name.is_empty() => Some((String::new(), name.clone())),
        [schema, name] if !name.is_empty() => Some((schema.clone(), name.clone())),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn object_names() {
        assert_eq!(parse_object_name("dbo.Customers"), Some(("dbo".into(), "Customers".into())));
        assert_eq!(parse_object_name("Customers"), Some(("".into(), "Customers".into())));
        assert_eq!(parse_object_name("[sales].[Order Details]"), Some(("sales".into(), "Order Details".into())));
        assert_eq!(parse_object_name("\"a.b\".\"c\"\"d\""), Some(("a.b".into(), "c\"d".into())));
        assert_eq!(parse_object_name("[x]]y].z"), Some(("x]y".into(), "z".into())));
        assert_eq!(parse_object_name(""), None);
        assert_eq!(parse_object_name("a.b.c"), None);
        assert_eq!(parse_object_name("[unclosed"), None);
    }

    #[test]
    fn iso_dates() {
        assert_eq!(serial_to_iso(45_000.5, &DbValue::DateTime(0.0)).as_deref(), Some("2023-03-15T12:00:00.000"));
        assert_eq!(serial_to_iso(45_000.0, &DbValue::Date(0.0)).as_deref(), Some("2023-03-15"));
        assert_eq!(serial_to_iso(0.75, &DbValue::Time(0.0)).as_deref(), Some("18:00:00.000"));
        assert_eq!(serial_to_iso(1.0, &DbValue::Date(0.0)).as_deref(), Some("1900-01-01"));
        assert_eq!(serial_to_iso(61.0, &DbValue::Date(0.0)).as_deref(), Some("1900-03-01"));
        assert_eq!(serial_to_iso(60.0, &DbValue::Date(0.0)), None);
        assert_eq!(serial_to_iso(f64::NAN, &DbValue::Date(0.0)), None);
    }

    #[test]
    fn credential_debug_hides_the_password() {
        let c = Credential { user: "sa".into(), password: "hunter2".into() };
        assert!(!format!("{c:?}").contains("hunter2"));
    }
}

//! External data sources (Data › Get Data › From Database).
//!
//! The engine never opens a network connection itself: it compiles to wasm, where there are no
//! sockets. A host that can reach databases (the desktop app, the CLI) installs a
//! [`DataProvider`] with [`crate::Session::set_data_provider`]; without one, the database commands
//! say so instead of failing silently.
//!
//! Credentials live in the session ([`crate::Session::set_credential`]), never in the workbook or
//! the journal.

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
    fn credential_debug_hides_the_password() {
        let c = Credential { user: "sa".into(), password: "hunter2".into() };
        assert!(!format!("{c:?}").contains("hunter2"));
    }
}

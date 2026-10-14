//! External data connections (Data › Get Data › From Database): where the rows of a worksheet
//! table come from, so Refresh can fetch them again.
//!
//! A connection never holds a password. Credentials belong to the application (it asks for them,
//! or keeps them in the system's credential store), never to the file.

use serde::{Deserialize, Serialize};

/// The database server a connection reads from.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum DataSource {
    /// A Microsoft SQL Server (or Azure SQL) database, signed in with a SQL Server login.
    #[serde(rename_all = "camelCase")]
    SqlServer {
        /// `host`, `host,port` or `host\instance`, as typed.
        server: String,
        /// The database; empty means the login's default database.
        #[serde(default)]
        database: String,
        /// Encrypt the connection (TLS).
        #[serde(default = "yes")]
        encrypt: bool,
        /// Accept a server certificate that isn't signed by a trusted authority (a self-signed
        /// certificate on a local server).
        #[serde(default)]
        trust_server_certificate: bool,
        /// How to sign in.
        #[serde(default)]
        auth: SqlAuth,
    },
}

/// How a SQL Server connection signs in.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SqlAuth {
    /// A SQL Server login (user name and password).
    #[default]
    Sql,
    /// A Windows account given by name and password (`DOMAIN\user`).
    Windows,
    /// The Windows account the app runs as: no password asked.
    Integrated,
}

fn yes() -> bool {
    true
}

impl DataSource {
    /// The server this source signs in to: credentials are kept per server.
    pub fn server(&self) -> &str {
        match self {
            DataSource::SqlServer { server, .. } => server,
        }
    }
    pub fn auth(&self) -> SqlAuth {
        match self {
            DataSource::SqlServer { auth, .. } => *auth,
        }
    }
    pub fn database(&self) -> &str {
        match self {
            DataSource::SqlServer { database, .. } => database,
        }
    }
}

/// What a connection asks its source for.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum QueryCommand {
    /// Every row of a table or view.
    Table { schema: String, name: String },
    /// A SQL statement, run as written.
    Sql { text: String },
}

impl QueryCommand {
    /// The name a query loaded from this command gets: `dbo Customers` for a table, `Query1`
    /// style names for statements are chosen by the caller.
    pub fn object_name(&self) -> Option<String> {
        match self {
            QueryCommand::Table { schema, name } if schema.is_empty() => Some(name.clone()),
            QueryCommand::Table { schema, name } => Some(format!("{schema} {name}")),
            QueryCommand::Sql { .. } => None,
        }
    }
}

/// A workbook connection and the table it fills.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Connection {
    pub id: u32,
    /// The query's name (Queries & Connections), unique in the workbook.
    pub name: String,
    #[serde(default)]
    pub description: String,
    pub source: DataSource,
    /// The login name to sign in with. The password is never stored.
    #[serde(default)]
    pub user: String,
    pub command: QueryCommand,
    /// The worksheet table the rows are loaded into.
    pub table: String,
    /// Refresh when the workbook is opened.
    #[serde(default)]
    pub refresh_on_open: bool,
    /// Refresh every so many minutes while the workbook is open; 0 = never.
    #[serde(default)]
    pub refresh_minutes: u32,
    /// When the rows were last fetched, as a date-time serial number (1900 date system).
    #[serde(default)]
    pub refreshed: Option<f64>,
    /// How many rows the last refresh loaded.
    #[serde(default)]
    pub rows: u64,
    /// The Power Query (by name) this connection came from in an Excel file. Its M formula and
    /// connection stay in the file unchanged; GridCraft runs the same SQL itself.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub power_query: Option<String>,
    /// Its `<connection>` element is in the connections part kept from the file (set when the
    /// file has a Power Query): saving leaves that element as it is.
    #[serde(skip)]
    pub kept: bool,
}

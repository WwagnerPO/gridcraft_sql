//! The provider over tiberius and tokio.

use std::time::Duration;

use futures_util::TryStreamExt;
use gridcraft_engine::datasource::{self, Credential, DataProvider, DbColumn, DbError, DbObject, DbValue, QueryResult, Statement};
use gridcraft_engine::model::{DataSource, QueryCommand};
use tiberius::{AuthMethod, Client, ColumnData, Config, EncryptionLevel, QueryItem};
use tokio::net::TcpStream;
use tokio_util::compat::{Compat, TokioAsyncWriteCompatExt};

use crate::convert;

/// Reaches SQL Server for the engine.
#[derive(Clone, Debug)]
pub struct SqlServerProvider {
    /// How long to wait for the server to accept the connection and the login.
    pub connect_timeout: Duration,
    /// How long one statement may run.
    pub command_timeout: Duration,
}

impl Default for SqlServerProvider {
    fn default() -> Self {
        SqlServerProvider { connect_timeout: Duration::from_secs(15), command_timeout: Duration::from_secs(600) }
    }
}

type Conn = Client<Compat<TcpStream>>;

/// Login failures (18456 login failed, 18452 untrusted domain, 1848x password expired or must
/// change): the interface asks for a login again.
const LOGIN_ERRORS: &[u32] = &[18_456, 18_452, 18_486, 18_487, 18_488];

fn map_error(server: &str, e: tiberius::error::Error) -> DbError {
    use tiberius::error::Error as E;
    match e {
        E::Server(t) if LOGIN_ERRORS.contains(&t.code()) => DbError::Credentials(t.message().to_string()),
        E::Server(t) => DbError::Other(t.message().to_string()),
        E::Tls(m) => DbError::Encryption(convert::tls_message(&m)),
        E::Io { message, .. } => DbError::Other(format!("Couldn't reach the server `{server}`: {message}")),
        other => DbError::Other(other.to_string()),
    }
}

impl SqlServerProvider {
    fn config(&self, source: &DataSource, credential: &Credential, write: bool) -> Result<Config, DbError> {
        let DataSource::SqlServer { server, database, encrypt, trust_server_certificate } = source;
        let addr = convert::parse_server(server).map_err(DbError::Other)?;
        let mut c = Config::new();
        c.host(&addr.host);
        if let Some(p) = addr.port {
            c.port(p);
        }
        if let Some(i) = &addr.instance {
            c.instance_name(i);
        }
        if !database.is_empty() {
            c.database(database);
        }
        c.application_name("GridCraft");
        c.authentication(AuthMethod::sql_server(&credential.user, &credential.password));
        // Off still encrypts the login itself whenever the server offers TLS.
        c.encryption(if *encrypt { EncryptionLevel::Required } else { EncryptionLevel::Off });
        if *trust_server_certificate {
            c.trust_cert();
        }
        c.handshake_timeout(Some(self.connect_timeout));
        c.command_timeout(Some(self.command_timeout));
        // Read intent unless the call writes (Run SQL, Export, Save to Database).
        c.readonly(!write);
        Ok(c)
    }

    async fn connect(&self, source: &DataSource, credential: &Credential, write: bool) -> Result<Conn, DbError> {
        let server = source.server().to_string();
        let mut config = self.config(source, credential, write)?;
        // Azure SQL may redirect the login once to another address.
        for _ in 0..2 {
            let tcp = self.tcp(&server, &config).await?;
            match Client::connect(config.clone(), tcp.compat_write()).await {
                Ok(c) => return Ok(c),
                Err(tiberius::error::Error::Routing { host, port }) => {
                    config.host(&host);
                    config.port(port);
                }
                Err(e) => return Err(map_error(&server, e)),
            }
        }
        Err(DbError::Other(format!("The server `{server}` redirected the connection too many times.")))
    }

    async fn tcp(&self, server: &str, config: &Config) -> Result<TcpStream, DbError> {
        use tiberius::SqlBrowser;
        let named = convert::parse_server(server).map(|a| a.instance.is_some() && a.port.is_none()).unwrap_or(false);
        let connect = async {
            if named {
                TcpStream::connect_named(config).await.map_err(|e| map_error(server, e))
            } else {
                TcpStream::connect(config.get_addr()).await.map_err(|e| map_error(server, e.into()))
            }
        };
        let tcp = tokio::time::timeout(self.connect_timeout, connect)
            .await
            .map_err(|_| DbError::Other(format!("The server `{server}` didn't answer in {} seconds.", self.connect_timeout.as_secs())))??;
        let _ = tcp.set_nodelay(true);
        Ok(tcp)
    }

    /// Runs `work` on a connection, on a thread of its own with its own runtime, and closes it.
    fn run<T: Send + 'static>(
        &self,
        source: &DataSource,
        credential: &Credential,
        write: bool,
        work: impl for<'c> FnOnce(&'c mut Conn) -> futures_util::future::BoxFuture<'c, Result<T, DbError>> + Send + 'static,
    ) -> Result<T, DbError> {
        let this = self.clone();
        let source = source.clone();
        let credential = credential.clone();
        let handle = std::thread::Builder::new()
            .name("gridcraft-sqlserver".into())
            .spawn(move || {
                let rt = tokio::runtime::Builder::new_current_thread()
                    .enable_io()
                    .enable_time()
                    .build()
                    .map_err(|e| DbError::Other(format!("Couldn't start the connection: {e}")))?;
                rt.block_on(async move {
                    let mut client = this.connect(&source, &credential, write).await?;
                    let r = work(&mut client).await;
                    let _ = client.close().await;
                    r
                })
            })
            .map_err(|e| DbError::Other(format!("Couldn't start the connection: {e}")))?;
        handle.join().unwrap_or_else(|_| Err(DbError::Other("The database connection stopped unexpectedly.".into())))
    }
}

/// The first result set of `sql`: column names and at most `max_rows` rows.
async fn first_result(server: &str, client: &mut Conn, sql: String, max_rows: usize) -> Result<QueryResult, DbError> {
    let stream = client.simple_query(sql).await.map_err(|e| map_error(server, e))?;
    collect(server, stream, max_rows).await
}

async fn collect(server: &str, mut stream: tiberius::QueryStream<'_>, max_rows: usize) -> Result<QueryResult, DbError> {
    let mut out = QueryResult::default();
    let mut current: Option<usize> = None;
    let mut cells = 0usize;
    while let Some(item) = stream.try_next().await.map_err(|e| map_error(server, e))? {
        match item {
            QueryItem::Metadata(m) => {
                if current.is_some() {
                    break;
                }
                if !m.columns().is_empty() {
                    out.columns = m.columns().iter().map(|c| c.name().to_string()).collect();
                    current = Some(m.result_index());
                }
            }
            QueryItem::Row(row) => {
                if current != Some(row.result_index()) {
                    if current.is_some() {
                        break;
                    }
                    continue;
                }
                if out.rows.len() >= max_rows || cells.saturating_add(out.columns.len()) > datasource::MAX_CELLS {
                    out.truncated = true;
                    break;
                }
                cells += out.columns.len();
                out.rows.push(row.cells().map(|(_, v)| value(v)).collect());
            }
        }
    }
    Ok(out)
}

fn value(v: &ColumnData<'_>) -> DbValue {
    let num = |x: Option<f64>| x.map_or(DbValue::Null, DbValue::Number);
    let date_or_text = |serial: Option<f64>, days_0001: i64, make: fn(f64) -> DbValue| match serial {
        Some(s) => make(s),
        None => DbValue::Text(convert::iso_date(days_0001)),
    };
    match v {
        ColumnData::U8(x) => num(x.map(f64::from)),
        ColumnData::I16(x) => num(x.map(f64::from)),
        ColumnData::I32(x) => num(x.map(f64::from)),
        ColumnData::I64(x) => num(x.map(|n| n as f64)),
        ColumnData::F32(x) => num(x.map(f64::from)),
        ColumnData::F64(x) => num(*x),
        ColumnData::Bit(x) => x.map_or(DbValue::Null, DbValue::Bool),
        ColumnData::String(x) => x.as_ref().map_or(DbValue::Null, |s| DbValue::Text(s.to_string())),
        ColumnData::Guid(x) => x.map_or(DbValue::Null, |g| DbValue::Text(g.to_string().to_uppercase())),
        ColumnData::Binary(x) => x.as_ref().map_or(DbValue::Null, |b| DbValue::Text(convert::hex(b, datasource::MAX_TEXT))),
        ColumnData::Numeric(x) => num(x.map(f64::from)),
        ColumnData::Xml(x) => x.as_ref().map_or(DbValue::Null, |s| DbValue::Text(s.to_string())),
        ColumnData::DateTime(x) => x.map_or(DbValue::Null, |d| {
            date_or_text(convert::datetime(d.days(), d.seconds_fragments()), d.days() as i64 + 693_595, DbValue::DateTime)
        }),
        ColumnData::SmallDateTime(x) => x.map_or(DbValue::Null, |d| {
            date_or_text(convert::small_datetime(d.days(), d.seconds_fragments()), d.days() as i64 + 693_595, DbValue::DateTime)
        }),
        ColumnData::Time(x) => x.map_or(DbValue::Null, |t| DbValue::Time(convert::time(t.increments(), t.scale()))),
        ColumnData::Date(x) => x.map_or(DbValue::Null, |d| date_or_text(convert::date(d.days()), d.days() as i64, DbValue::Date)),
        ColumnData::DateTime2(x) => x.map_or(DbValue::Null, |d| {
            let (day, t) = (d.date().days(), d.time());
            date_or_text(convert::datetime2(day, t.increments(), t.scale(), 0), day as i64, DbValue::DateTime)
        }),
        ColumnData::DateTimeOffset(x) => x.map_or(DbValue::Null, |d| {
            let dt = d.datetime2();
            let (day, t) = (dt.date().days(), dt.time());
            date_or_text(convert::datetime2(day, t.increments(), t.scale(), d.offset()), day as i64, DbValue::DateTime)
        }),
    }
}

fn bind(q: &mut tiberius::Query<'_>, v: DbValue) {
    match v {
        DbValue::Null => q.bind(Option::<String>::None),
        DbValue::Bool(b) => q.bind(b),
        DbValue::Number(n) => q.bind(n),
        DbValue::Text(t) => q.bind(t),
        DbValue::Date(n) | DbValue::DateTime(n) | DbValue::Time(n) => match datasource::serial_to_iso(n, &v) {
            Some(iso) => q.bind(iso),
            None => q.bind(n),
        },
    }
}

/// Runs the statements in one transaction; any failure (or a statement that didn't affect the
/// rows it had to) rolls everything back.
async fn run_all(server: &str, client: &mut Conn, statements: Vec<Statement>) -> Result<Vec<u64>, DbError> {
    client.begin_transaction().await.map_err(|e| map_error(server, e))?;
    let mut out = Vec::with_capacity(statements.len());
    for (i, st) in statements.into_iter().enumerate() {
        let expect = st.expect_rows;
        let mut q = tiberius::Query::new(st.sql);
        for p in st.params {
            bind(&mut q, p);
        }
        let r = match q.execute(client).await {
            Ok(r) => r.total(),
            Err(e) => {
                let _ = client.rollback_transaction().await;
                return Err(map_error(server, e));
            }
        };
        if let Some(want) = expect
            && r != want
        {
            let _ = client.rollback_transaction().await;
            return Err(DbError::Other(format!(
                "Nothing was changed: statement {} affected {r} rows instead of {want} (the row may have been changed or deleted on the server since it was loaded). Refresh and try again.",
                i + 1
            )));
        }
        out.push(r);
    }
    client.commit_transaction().await.map_err(|e| map_error(server, e))?;
    Ok(out)
}

/// The statement for a command; `top` limits a table read (previews).
fn statement(command: &QueryCommand, top: Option<usize>) -> String {
    match command {
        QueryCommand::Sql { text } => text.clone(),
        QueryCommand::Table { schema, name } => {
            let object = if schema.is_empty() {
                convert::quote_identifier(name)
            } else {
                format!("{}.{}", convert::quote_identifier(schema), convert::quote_identifier(name))
            };
            match top {
                Some(n) => format!("SELECT TOP ({n}) * FROM {object}"),
                None => format!("SELECT * FROM {object}"),
            }
        }
    }
}

impl DataProvider for SqlServerProvider {
    fn columns(&self, source: &DataSource, credential: &Credential, schema: &str, name: &str) -> Result<Vec<DbColumn>, DbError> {
        let server = source.server().to_string();
        let object = if schema.is_empty() {
            convert::quote_identifier(name)
        } else {
            format!("{}.{}", convert::quote_identifier(schema), convert::quote_identifier(name))
        };
        let r = self.run(source, credential, false, move |c| {
            Box::pin(async move {
                let sql = "SELECT c.name, t.name, c.is_nullable, \
                    CAST(CASE WHEN c.is_identity = 1 OR c.is_computed = 1 OR t.name IN ('timestamp', 'rowversion') THEN 1 ELSE 0 END AS bit), \
                    CAST(CASE WHEN EXISTS (SELECT 1 FROM sys.index_columns ic JOIN sys.indexes i ON i.object_id = ic.object_id AND i.index_id = ic.index_id \
                        WHERE i.is_primary_key = 1 AND ic.object_id = c.object_id AND ic.column_id = c.column_id) THEN 1 ELSE 0 END AS bit) \
                    FROM sys.columns c JOIN sys.types t ON t.user_type_id = c.user_type_id \
                    WHERE c.object_id = OBJECT_ID(@P1) ORDER BY c.column_id";
                let mut q = tiberius::Query::new(sql);
                q.bind(object);
                let stream = q.query(c).await.map_err(|e| map_error(&server, e))?;
                collect(&server, stream, 10_000).await
            })
        })?;
        let text = |v: Option<&DbValue>| match v {
            Some(DbValue::Text(t)) => t.clone(),
            _ => String::new(),
        };
        let flag = |v: Option<&DbValue>| matches!(v, Some(DbValue::Bool(true)));
        let cols: Vec<DbColumn> = r
            .rows
            .iter()
            .map(|row| DbColumn {
                name: text(row.first()),
                sql_type: text(row.get(1)),
                nullable: flag(row.get(2)),
                generated: flag(row.get(3)),
                primary_key: flag(row.get(4)),
            })
            .collect();
        if cols.is_empty() {
            return Err(DbError::Other(format!(
                "The table `{}` doesn't exist, or this login can't see it.",
                if schema.is_empty() { name.to_string() } else { format!("{schema}.{name}") }
            )));
        }
        Ok(cols)
    }

    fn execute(&self, source: &DataSource, credential: &Credential, statements: &[Statement]) -> Result<Vec<u64>, DbError> {
        let server = source.server().to_string();
        let statements = statements.to_vec();
        self.run(source, credential, true, move |c| Box::pin(async move { run_all(&server, c, statements).await }))
    }

    fn databases(&self, source: &DataSource, credential: &Credential) -> Result<Vec<String>, DbError> {
        let server = source.server().to_string();
        let r = self.run(source, credential, false, move |c| {
            Box::pin(async move {
                let sql = "SELECT name FROM sys.databases WHERE HAS_DBACCESS(name) = 1 ORDER BY name".to_string();
                first_result(&server, c, sql, 100_000).await
            })
        })?;
        Ok(r.rows
            .into_iter()
            .filter_map(|row| match row.into_iter().next() {
                Some(DbValue::Text(t)) => Some(t),
                _ => None,
            })
            .collect())
    }

    fn objects(&self, source: &DataSource, credential: &Credential) -> Result<Vec<DbObject>, DbError> {
        let server = source.server().to_string();
        let r = self.run(source, credential, false, move |c| {
            Box::pin(async move {
                let sql = "SELECT TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE FROM INFORMATION_SCHEMA.TABLES ORDER BY TABLE_SCHEMA, TABLE_NAME".to_string();
                first_result(&server, c, sql, 1_000_000).await
            })
        })?;
        let text = |v: Option<&DbValue>| match v {
            Some(DbValue::Text(t)) => t.clone(),
            _ => String::new(),
        };
        Ok(r.rows
            .iter()
            .map(|row| DbObject { schema: text(row.first()), name: text(row.get(1)), view: text(row.get(2)).eq_ignore_ascii_case("VIEW") })
            .filter(|o| !o.name.is_empty())
            .collect())
    }

    fn query(&self, source: &DataSource, credential: &Credential, command: &QueryCommand, max_rows: usize) -> Result<QueryResult, DbError> {
        let server = source.server().to_string();
        // A preview of a table asks the server for no more than it shows.
        let top = (max_rows <= 10_000).then_some(max_rows.saturating_add(1));
        let sql = statement(command, top);
        self.run(source, credential, false, move |c| Box::pin(async move { first_result(&server, c, sql, max_rows).await }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn statements_quote_names() {
        let t = QueryCommand::Table { schema: "dbo".into(), name: "Order]Details".into() };
        assert_eq!(statement(&t, None), "SELECT * FROM [dbo].[Order]]Details]");
        assert_eq!(statement(&t, Some(101)), "SELECT TOP (101) * FROM [dbo].[Order]]Details]");
        let t = QueryCommand::Table { schema: String::new(), name: "x".into() };
        assert_eq!(statement(&t, None), "SELECT * FROM [x]");
        let q = QueryCommand::Sql { text: "SELECT 1".into() };
        assert_eq!(statement(&q, Some(5)), "SELECT 1");
    }

    #[test]
    fn values_convert() {
        assert_eq!(value(&ColumnData::I32(Some(7))), DbValue::Number(7.0));
        assert_eq!(value(&ColumnData::I32(None)), DbValue::Null);
        assert_eq!(value(&ColumnData::Bit(Some(true))), DbValue::Bool(true));
        assert_eq!(value(&ColumnData::String(Some("hi".into()))), DbValue::Text("hi".into()));
        assert_eq!(value(&ColumnData::DateTime(Some(tiberius::time::DateTime::new(44_998, 0)))), DbValue::DateTime(45_000.0));
        assert_eq!(value(&ColumnData::Date(Some(tiberius::time::Date::new(0)))), DbValue::Text("0001-01-01".into()));
    }

    /// A server that isn't there: a clear error, no hang, no panic.
    #[test]
    fn unreachable_server_is_an_error() {
        let p = SqlServerProvider { connect_timeout: Duration::from_secs(2), command_timeout: Duration::from_secs(2) };
        let source = DataSource::SqlServer { server: "127.0.0.1,9".into(), database: String::new(), encrypt: true, trust_server_certificate: false };
        let c = Credential { user: "sa".into(), password: "x".into() };
        let e = p.databases(&source, &c).unwrap_err();
        assert!(matches!(e, DbError::Other(_)), "{e:?}");
    }

    /// Against a real server when one is configured: `GRIDCRAFT_SQLSERVER_TEST=host;user;password`
    /// (for example a local `mssql/server` container).
    #[test]
    fn live_server() {
        let Ok(spec) = std::env::var("GRIDCRAFT_SQLSERVER_TEST") else { return };
        let mut parts = spec.splitn(3, ';');
        let (Some(server), Some(user), Some(password)) = (parts.next(), parts.next(), parts.next()) else { return };
        let p = SqlServerProvider::default();
        let c = Credential { user: user.into(), password: password.into() };
        let source = DataSource::SqlServer { server: server.into(), database: "master".into(), encrypt: true, trust_server_certificate: true };
        let dbs = p.databases(&source, &c).unwrap();
        assert!(dbs.iter().any(|d| d == "master"), "{dbs:?}");
        let sql = "SELECT CAST(1 AS int) AS n, N'ñ' AS t, CAST('2023-03-15T12:00:00' AS datetime2) AS d, CAST(NULL AS int) AS z, CAST(1.25 AS decimal(9,2)) AS m";
        let r = p.query(&source, &c, &QueryCommand::Sql { text: sql.into() }, 10).unwrap();
        assert_eq!(r.columns, ["n", "t", "d", "z", "m"]);
        assert_eq!(
            r.rows[0],
            vec![DbValue::Number(1.0), DbValue::Text("ñ".into()), DbValue::DateTime(45_000.5), DbValue::Null, DbValue::Number(1.25)]
        );
        let bad = Credential { user: user.into(), password: "wrong".into() };
        assert!(matches!(p.databases(&source, &bad), Err(DbError::Credentials(_))));
    }
}

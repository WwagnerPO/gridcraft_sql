//! Office Data Connection files (`.odc`, [MS-ODCFF]): a database connection saved on the
//! computer, apart from any workbook, as Excel keeps them in `Documents\My Data Sources`.
//!
//! The file is an HTML page whose `<xml id=docprops>` island names the connection and whose
//! `<xml id=msodc>` island holds the OLE DB connection string and the command. GridCraft writes
//! and reads SQL Server connections (the same ones [`crate::connections`] keeps in workbooks).
//! A connection file never holds a password.

use gridcraft_model::{Connection, DataSource, QueryCommand};

use crate::connections::{COMMAND_TABLE, command, connection_string, parse_connection_string, sql_server_source};
use crate::xml::{El, esc, parse};

/// Larger files aren't connection files.
pub const MAX_ODC_BYTES: usize = 1 << 20;

/// A connection saved in a file: where it reads from and what.
#[derive(Clone, Debug, PartialEq)]
pub struct ConnectionFile {
    pub name: String,
    pub description: String,
    pub source: DataSource,
    /// The login name; empty for the current Windows user or none.
    pub user: String,
    pub command: QueryCommand,
}

impl ConnectionFile {
    pub fn of(c: &Connection) -> ConnectionFile {
        ConnectionFile {
            name: c.name.clone(),
            description: c.description.clone(),
            source: c.source.clone(),
            user: c.user.clone(),
            command: c.command.clone(),
        }
    }

    /// A workbook connection with this file's source and command (id, table and refresh options
    /// left for the caller).
    fn connection(&self) -> Connection {
        Connection {
            id: 0,
            name: self.name.clone(),
            description: self.description.clone(),
            source: self.source.clone(),
            user: self.user.clone(),
            command: self.command.clone(),
            table: String::new(),
            refresh_on_open: false,
            refresh_minutes: 0,
            refreshed: None,
            rows: 0,
            power_query: None,
            kept: false,
        }
    }
}

/// The `.odc` file of a connection, UTF-8.
pub fn write_odc(f: &ConnectionFile) -> String {
    let c = f.connection();
    let (text, kind) = command(&c);
    let (prog_id, command_type) = if kind == COMMAND_TABLE { ("ODC.Table", "Table") } else { ("ODC.Database", "SQL") };
    let name = esc(&f.name);
    let description = esc(&f.description);
    let connection = esc(&connection_string(&c));
    let text = esc(&text);
    format!(
        r#"<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/x-ms-odc; charset=utf-8">
<meta name=ProgId content={prog_id}>
<meta name=SourceType content=OLEDB>
<title>{name}</title>
<xml id=docprops><o:DocumentProperties
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns="http://www.w3.org/TR/REC-html40">
  <o:Name>{name}</o:Name>
  <o:Description>{description}</o:Description>
 </o:DocumentProperties>
</xml><xml id=msodc><odc:OfficeDataConnection
  xmlns:odc="urn:schemas-microsoft-com:office:odc"
  xmlns="http://www.w3.org/TR/REC-html40">
  <odc:Connection odc:Type="OLEDB">
   <odc:ConnectionString>{connection}</odc:ConnectionString>
   <odc:CommandType>{command_type}</odc:CommandType>
   <odc:CommandText>{text}</odc:CommandText>
  </odc:Connection>
 </odc:OfficeDataConnection>
</xml>
</head>

<body>
<p>{name}</p>
</body>

</html>
"#
    )
}

/// The XML island whose root element is named `root` (any prefix), parsed.
fn island(text: &str, root: &str) -> Option<El> {
    let first = text.find(root)?;
    let start = text.get(..first)?.rfind('<')?;
    let close = text.rfind(root)?;
    let end = close + text.get(close..)?.find('>')? + 1;
    if end <= first {
        return None;
    }
    parse(text.get(start..end)?.as_bytes()).ok()
}

/// The text of `el`'s child `name`, trimmed.
fn child_text<'a>(el: &'a El, name: &str) -> Option<&'a str> {
    el.child(name).map(|c| c.text.trim())
}

/// Reads an `.odc` file. `None` when it isn't one, or its connection isn't a SQL Server one
/// GridCraft can run.
pub fn read_odc(bytes: &[u8]) -> Option<ConnectionFile> {
    if bytes.len() > MAX_ODC_BYTES {
        return None;
    }
    let text = crate::xml::to_utf8(bytes);
    let text = String::from_utf8_lossy(&text);
    let odc = island(&text, "OfficeDataConnection")?;
    let conn = odc.child("Connection")?;
    if !conn.attr("Type").is_some_and(|t| t.eq_ignore_ascii_case("OLEDB")) {
        return None;
    }
    let kv = parse_connection_string(child_text(conn, "ConnectionString")?);
    let table_command = match child_text(conn, "CommandType").unwrap_or("SQL").to_ascii_lowercase().as_str() {
        "table" => true,
        "sql" | "default" => false,
        _ => return None,
    };
    let (source, user, command) = sql_server_source(&kv, child_text(conn, "CommandText").unwrap_or(""), table_command)?;
    let props = island(&text, "DocumentProperties");
    let prop = |k: &str| props.as_ref().and_then(|p| child_text(p, k)).filter(|v| !v.is_empty()).map(str::to_string);
    let name = prop("Name").or_else(|| command.object_name()).unwrap_or_else(|| source.server().to_string());
    Some(ConnectionFile { name, description: prop("Description").unwrap_or_default(), source, user, command })
}

/// A file name for a connection, as Excel names them: `server database object.odc`, without the
/// characters file systems refuse.
pub fn odc_file_name(f: &ConnectionFile) -> String {
    let what = match &f.command {
        QueryCommand::Table { name, .. } => name.clone(),
        QueryCommand::Sql { .. } => f.name.clone(),
    };
    let base = [f.source.server(), f.source.database(), what.as_str()].iter().filter(|p| !p.trim().is_empty()).copied().collect::<Vec<_>>().join(" ");
    let clean: String = base
        .chars()
        .map(|c| if matches!(c, '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*') || c.is_control() { '_' } else { c })
        .take(120)
        .collect();
    let clean = clean.trim().trim_end_matches('.').to_string();
    format!("{}.odc", if clean.is_empty() { "Connection".to_string() } else { clean })
}

#[cfg(test)]
mod tests {
    use super::*;
    use gridcraft_model::SqlAuth;

    fn file(command: QueryCommand, auth: SqlAuth) -> ConnectionFile {
        ConnectionFile {
            name: "Customers & Co".into(),
            description: "Every <customer>".into(),
            source: DataSource::SqlServer {
                server: "db.example,1433".into(),
                database: "Shop".into(),
                encrypt: true,
                trust_server_certificate: true,
                auth,
            },
            user: if auth == SqlAuth::Integrated { String::new() } else { "reader".into() },
            command,
        }
    }

    #[test]
    fn table_connections_round_trip() {
        let f = file(QueryCommand::Table { schema: "dbo".into(), name: "Customers".into() }, SqlAuth::Sql);
        let text = write_odc(&f);
        assert!(
            text.contains("<odc:CommandText>&quot;Shop&quot;.&quot;dbo&quot;.&quot;Customers&quot;</odc:CommandText>")
                || text.contains(r#""Shop"."dbo"."Customers""#)
        );
        assert!(!text.to_ascii_lowercase().contains("password="));
        assert_eq!(read_odc(text.as_bytes()), Some(f));
    }

    #[test]
    fn sql_and_windows_connections_round_trip() {
        let sql = QueryCommand::Sql { text: "SELECT * FROM T WHERE a < 3 AND b = 'x'".into() };
        for auth in [SqlAuth::Integrated, SqlAuth::Windows] {
            let mut f = file(sql.clone(), auth);
            if auth == SqlAuth::Windows {
                f.user = "CORP\\ana".into();
            }
            assert_eq!(read_odc(write_odc(&f).as_bytes()), Some(f));
        }
    }

    #[test]
    fn a_file_written_elsewhere_reads() {
        // The layout MS-ODCFF describes, UTF-16 with a byte order mark, no document properties.
        let text = r#"<html><head><meta http-equiv=Content-Type content="text/x-ms-odc; charset=utf-16">
<xml id=msodc><odc:OfficeDataConnection xmlns:odc="urn:schemas-microsoft-com:office:odc" xmlns="http://www.w3.org/TR/REC-html40">
<odc:Connection odc:Type="OLEDB"><odc:ConnectionString>Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=Shop;Data Source=srv\SQL2022;Use Encryption for Data=False</odc:ConnectionString>
<odc:CommandType>Table</odc:CommandType><odc:CommandText>[dbo].[Orders]</odc:CommandText></odc:Connection></odc:OfficeDataConnection></xml></head></html>"#;
        let mut bytes = vec![0xFF, 0xFE];
        bytes.extend(text.encode_utf16().flat_map(u16::to_le_bytes));
        let f = read_odc(&bytes).expect("reads");
        assert_eq!(f.name, "dbo Orders");
        assert_eq!(
            f.source,
            DataSource::SqlServer {
                server: "srv\\SQL2022".into(),
                database: "Shop".into(),
                encrypt: false,
                trust_server_certificate: false,
                auth: SqlAuth::Integrated
            }
        );
        assert_eq!(f.command, QueryCommand::Table { schema: "dbo".into(), name: "Orders".into() });
    }

    #[test]
    fn other_files_are_refused() {
        assert_eq!(read_odc(b""), None);
        assert_eq!(read_odc(b"<html><body>hello</body></html>"), None);
        assert_eq!(read_odc(b"<xml><odc:OfficeDataConnection"), None);
        // Another provider (an Access database) isn't a SQL Server connection.
        let access = br#"<xml id=msodc><odc:OfficeDataConnection xmlns:odc="urn:schemas-microsoft-com:office:odc"><odc:Connection odc:Type="OLEDB"><odc:ConnectionString>Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\a.accdb</odc:ConnectionString><odc:CommandType>Table</odc:CommandType><odc:CommandText>T</odc:CommandText></odc:Connection></odc:OfficeDataConnection></xml>"#;
        assert_eq!(read_odc(access), None);
        assert_eq!(read_odc(&vec![b' '; MAX_ODC_BYTES + 1]), None);
    }

    #[test]
    fn file_names_are_safe() {
        let f = file(QueryCommand::Table { schema: "dbo".into(), name: "A/B:C".into() }, SqlAuth::Sql);
        assert_eq!(odc_file_name(&f), "db.example,1433 Shop A_B_C.odc");
        let mut f = file(QueryCommand::Sql { text: "SELECT 1".into() }, SqlAuth::Sql);
        f.name = "...".into();
        f.source =
            DataSource::SqlServer { server: "".into(), database: "".into(), encrypt: true, trust_server_certificate: false, auth: SqlAuth::Sql };
        assert_eq!(odc_file_name(&f), "Connection.odc");
    }
}

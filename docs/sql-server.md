# Loading data from SQL Server

> **Last reviewed:** 2026-10-13 · **Last updated:** 2026-10-13 · **Change:** major (created with the SQL Server connector) · **Target:** Microsoft Excel (Microsoft 365)

GridCraft loads tables, views and SQL statements from Microsoft SQL Server and Azure SQL into a
worksheet table, and refreshes them later. It speaks SQL Server's protocol (TDS) directly through
the pure-Rust [`tiberius`](https://crates.io/crates/tiberius) driver with TLS by rustls, so no ODBC
driver, OpenSSL or other system library has to be installed. The desktop app and the CLI can
connect; the web version can't (browsers have no sockets) and says so.

## In the app

1. **Data › Get Data › From SQL Server Database…**
2. Type the **Server** (`host`, `host,port` or `host\instance`) and, optionally, the **Database**.
   Under **Advanced options**: a **SQL statement** (needs a database), **Encrypt connection** (on by
   default) and **Trust the server certificate** (for a local server with a self-signed
   certificate).
3. Sign in with a **SQL Server login** (user name and password). If the server's certificate is
   refused (expired, self-signed, issued for another name), the dialog says why and offers
   **Trust the certificate and connect**. SQL Server checks the certificate even for an
   unencrypted connection, since the login itself is always encrypted.
4. In the **Navigator**, pick a database if you didn't name one, search the tables and views, click
   one to preview its first rows, tick **Select multiple items** to load several, then **Load**.
   With a SQL statement, the preview of its result shows instead.
5. Each query lands on a new sheet as a formatted table (`TableStyleMedium2`): numbers as numbers,
   dates and times with date formats, `bit` as TRUE/FALSE, `NULL` as an empty cell, text always as
   text (a value like `=1+1` or `0012` is never interpreted).

**Refresh** fetches the rows again: Data › Refresh All refreshes every query and then every
PivotTable; Refresh (in the Refresh All menu) refreshes the query of the table at the active cell. The table
grows or shrinks with the result; a refresh that would overwrite cells next to the table is refused
instead. A totals row stays at the bottom.

**Data › Queries & Connections** lists the workbook's queries with rows loaded, last refresh,
server and database, and Refresh, Properties and Delete for each. **Properties** sets the query
name, description, *Refresh data when opening the file*, *Refresh every N minutes* and, for a
statement, the SQL text. Deleting a query keeps its table with the rows last loaded; converting the
table to a range removes its query, as in Excel.

## Writing to the database

**Data › Database** writes back, always in **one transaction** (all of it or nothing), after showing
what it will send. Undo in GridCraft doesn't undo a write on the server, and these commands aren't
journaled, so replaying a recorded script never writes again.

- **Save Changes to Database…** sends what changed in a loaded table since its last load or
  refresh: new rows are inserted, edited cells updated, removed rows deleted, matched by the
  table's **primary key** (a table without one can't be saved back). Each update and delete must
  hit exactly one row, so a row someone changed or deleted on the server meanwhile stops the save
  and nothing is written. Identity and computed columns are left to the server; afterwards the
  table is refreshed. Only queries that load a whole table can be saved (not SQL statements), and
  after reopening a file, refresh first.
- **Export to SQL Server…** writes a range or table (column names in its first row) to a table:
  **New table** creates it with types read from the cells (text, whole and decimal numbers, dates,
  date-times, times, TRUE/FALSE), **Add rows** inserts into an existing table matching columns by
  name, **Replace all rows** deletes its rows first. **Preview** shows the rows and column types.
- **Run SQL…** runs a statement or stored procedure (`EXEC dbo.CloseMonth`), optionally with cell
  values as parameters `@P1`, `@P2`… (`UPDATE dbo.Prices SET Price = @P1 WHERE Code = @P2` with
  cells `B2, C2`), after asking.

## Credentials

Logins are kept **in memory for the session only**, per server. They never go into the workbook,
the journal, recorded scripts or logs (`password` is a secret parameter). When a refresh needs a
login the session doesn't have, or the server refuses the one it has, the app asks for it and then
refreshes. **Data › Get Data › Data Source Settings…** lists the kept logins and clears or replaces
them.

## In the file

Queries are saved in the XLSX as Excel does for an OLE DB SQL Server connection: a `<connection>`
with `<dbPr>` (provider `SQLOLEDB.1`, server, database and user name, never the password) in
`xl/connections.xml`, and a query table (`xl/queryTables/`) binding the worksheet table to it. A
file whose connections are all of this kind reopens with its queries refreshable; other connections
(Power Query, ODBC, Windows sign-in…) are kept unchanged on save, and queries added in GridCraft
are added next to them.

## Commands (CLI, MCP, control channel)

| Command | Does |
|---|---|
| `data.fromSqlServer` | `{server, database?, table? \| tables? \| sql?, user?, password?, encrypt?, trustServerCertificate?, name?}` loads into new sheets |
| `data.sqlServerNavigator` | `{server, database?, …}` lists databases, or a database's tables and views |
| `data.sqlServerPreview` | `{server, database?, table \| sql, rows?}` the first rows as text |
| `data.refresh` / `data.refreshAll` | refresh one query (`table` or `connection`) / all of them, then PivotTables |
| `data.dataSourceSettings` | list, set (`{server, user, password}`) or clear (`{server, clear: true}`) kept logins |
| `data.connectionProperties` | `{connection, name?, description?, refreshOnOpen?, refreshMinutes?, sql?, table?, user?}` |
| `data.deleteConnection` | `{connection}` |
| `data.queriesConnections` | `{tables, connections}` |
| `data.saveToDatabase` | `{table? \| connection?, dryRun?}` → `{inserted, updated, deleted}` |
| `data.exportToSqlServer` | `{table, range?, mode?: create \| append \| replace, connection? \| server…, dryRun?}` |
| `data.sqlExecute` | `{sql, params? \| cells?, connection? \| server…}` → `{rowsAffected}` |

```sh
gridcraft-cli run --cmd 'data.fromSqlServer={"server":"db.example,1433","database":"Shop","table":"dbo.Customers","user":"reader","password":"…"}' --out customers.xlsx
gridcraft-cli run --in customers.xlsx --cmd 'data.dataSourceSettings={"server":"db.example,1433","user":"reader","password":"…"}' --cmd 'data.refreshAll={}' --out customers.xlsx
```

## Limits

- **Sign-in:** SQL Server logins only. Windows (integrated) and Microsoft Entra sign-in are not
  supported yet.
- **Reading** opens the connection read-intent; only the Database menu's commands write.
- **Size:** at most 1,048,575 rows and 50 million cells per query; larger results load the first
  rows and say so. Text is cut at 32,767 characters, binary columns load as `0x…` hex text.
- **No Power Query editor:** queries can't be shaped (filter, merge, change types) before loading;
  a SQL statement is the way to shape them.
- Queries run while the window waits; the connection times out after 15 seconds and a statement
  after 10 minutes.
- Last-refreshed times are UTC and aren't saved in the file.

## Testing against a real server

The connector's live test runs when `GRIDCRAFT_SQLSERVER_TEST=host;user;password` is set, for
example against a local SQL Server container:

```sh
GRIDCRAFT_SQLSERVER_TEST='localhost,1433;sa;YourStrong!Passw0rd' cargo test -p gridcraft-sqlserver live_server
```

## Revision history

| Date | Change | Summary |
|---|---|---|
| 2026-10-13 | major | Created with the SQL Server connector |

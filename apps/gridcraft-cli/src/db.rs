//! Sessions that can reach databases (`data.fromSqlServer`, `data.refresh` and the other
//! database commands). `main.rs` calls these with one line each.

use std::sync::Arc;

use gridcraft_engine::Session;

/// Installs the SQL Server provider unless the session has one.
pub fn install(s: &mut Session) {
    if s.data_provider().is_none() {
        s.set_data_provider(Arc::new(gridcraft_sqlserver::SqlServerProvider::default()));
    }
}

pub fn with_databases(mut s: Session) -> Session {
    install(&mut s);
    s
}

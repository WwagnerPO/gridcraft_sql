//! The SQL Server connector: a [`DataProvider`](gridcraft_engine::datasource::DataProvider) that
//! speaks TDS to Microsoft SQL Server and Azure SQL through `tiberius`, with TLS by rustls. No
//! ODBC driver, OpenSSL or other system library is needed.
//!
//! Native only (wasm has no sockets): the apps install [`SqlServerProvider`] with
//! `Session::set_data_provider`. Each call opens a connection on its own thread and runtime, so
//! it works whatever runtime (if any) the caller runs on, and closes it when done.
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]
#![forbid(unsafe_code)]

pub mod convert;
#[cfg(not(target_arch = "wasm32"))]
mod native;

#[cfg(not(target_arch = "wasm32"))]
pub use native::SqlServerProvider;

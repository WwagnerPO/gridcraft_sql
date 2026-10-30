//! GridCraft's VBA engine: a clean-room interpreter for Visual Basic for Applications, built from
//! Microsoft's public [MS-VBAL] (the language) and [MS-OVBA] (the on-disk format). It runs the
//! source of a workbook's macros; it does not execute the compiled p-code.
//!
//! The crate knows the VBA *language* and its host-independent built-ins (text, maths, dates,
//! Collection, Dictionary). Excel's object model — `Range`, `Cells`, `MsgBox`, worksheets — is
//! supplied by the embedder through the [`Host`] trait, so this crate never depends on the engine
//! or any UI.
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]
#![forbid(unsafe_code)]

pub mod ast;
pub mod builtins;
pub mod error;
pub mod files;
pub mod interp;
pub mod lexer;
pub mod parser;
pub mod value;

pub use error::VbaError;
pub use interp::{Host, Interp, Program};
pub use value::{HostRef, Obj, Value};

/// Parses modules (name, source) and runs `entry` (`Proc` or `Module.Proc`) against `host`.
pub fn run(modules: &[(String, String)], entry: &str, args: Vec<Value>, host: &mut dyn Host) -> Result<Value, VbaError> {
    let program = Program::parse(modules)?;
    Interp::new(program, host).run(entry, args)
}

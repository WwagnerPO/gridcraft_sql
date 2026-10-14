//! VBA run-time and compile errors, numbered as VBA numbers them (`Err.Number`).

/// An error: its number, description and where it came from.
#[derive(Clone, Debug, PartialEq)]
pub struct VbaError {
    pub number: i32,
    pub description: String,
    /// The procedure (or module) running when it was raised.
    pub source: String,
    /// The source line, when known (1-based within its module).
    pub line: Option<usize>,
    /// The module.
    pub module: Option<String>,
}

impl std::fmt::Display for VbaError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Run-time error '{}': {}", self.number, self.description)?;
        if let (Some(m), Some(l)) = (&self.module, self.line) {
            write!(f, " ({m}, line {l})")?;
        }
        Ok(())
    }
}

pub fn err(number: i32, description: impl Into<String>) -> VbaError {
    VbaError { number, description: description.into(), source: String::new(), line: None, module: None }
}

/// A compile error (syntax), reported before anything runs.
pub fn syntax(module: &str, line: usize, description: impl Into<String>) -> VbaError {
    VbaError {
        number: 0,
        description: format!("Compile error: {}", description.into()),
        source: String::new(),
        line: Some(line),
        module: Some(module.to_string()),
    }
}

/// The description VBA gives an error number.
pub fn description(number: i32) -> &'static str {
    match number {
        3 => "Return without GoSub",
        5 => "Invalid procedure call or argument",
        6 => "Overflow",
        7 => "Out of memory",
        9 => "Subscript out of range",
        10 => "This array is fixed or temporarily locked",
        11 => "Division by zero",
        13 => "Type mismatch",
        28 => "Out of stack space",
        35 => "Sub or Function not defined",
        52 => "Bad file name or number",
        53 => "File not found",
        91 => "Object variable or With block variable not set",
        92 => "For loop not initialized",
        94 => "Invalid use of Null",
        424 => "Object required",
        438 => "Object doesn't support this property or method",
        440 => "Automation error",
        450 => "Wrong number of arguments or invalid property assignment",
        457 => "This key is already associated with an element of this collection",
        1004 => "Application-defined or object-defined error",
        _ => "Application-defined or object-defined error",
    }
}

use super::*;
use std::collections::HashMap;

/// A tiny host: a bag of named cells, a MsgBox log and a Debug.Print log.
#[derive(Default)]
struct TestHost {
    cells: HashMap<String, Value>,
    messages: Vec<String>,
    prints: Vec<String>,
    input: Option<String>,
}

impl Host for TestHost {
    fn global(&mut self, name: &str) -> Result<Option<Value>, VbaError> {
        Ok(match name.to_ascii_lowercase().as_str() {
            "vbcrlf" => Some(Value::str("\r\n")),
            "vbtab" => Some(Value::str("\t")),
            "vbok" | "vbokonly" => Some(Value::Long(if name.eq_ignore_ascii_case("vbok") { 1 } else { 0 })),
            "vbyes" => Some(Value::Long(6)),
            "vbyesno" => Some(Value::Long(4)),
            "true" => Some(Value::Bool(true)),
            "false" => Some(Value::Bool(false)),
            "cells" => Some(Value::Object(Obj::Host(HostRef { class: "Cells".into(), id: "sheet1".into() }))),
            _ => None,
        })
    }
    fn get_member(&mut self, obj: &HostRef, member: &str, args: &[Value]) -> Result<Value, VbaError> {
        match obj.class.as_str() {
            // Cells(r, c) -> a Range at "r,c".
            "Cells" => {
                let (r, c) = (args.first().map(|v| v.to_i64()).transpose()?.unwrap_or(0), args.get(1).map(|v| v.to_i64()).transpose()?.unwrap_or(0));
                Ok(Value::Object(Obj::Host(HostRef { class: "Range".into(), id: format!("{r},{c}") })))
            }
            "Range" => match member.to_ascii_lowercase().as_str() {
                "value" | "value2" | "_default" => Ok(self.cells.get(&obj.id).cloned().unwrap_or(Value::Empty)),
                _ => Err(err(438, "no member")),
            },
            _ => Err(err(438, "no member")),
        }
    }
    fn set_member(&mut self, obj: &HostRef, member: &str, _args: &[Value], value: Value, _by_ref: bool) -> Result<(), VbaError> {
        if obj.class == "Range" && matches!(member.to_ascii_lowercase().as_str(), "value" | "value2" | "_default") {
            self.cells.insert(obj.id.clone(), value);
            return Ok(());
        }
        Err(err(438, "no member"))
    }
    fn msgbox(&mut self, prompt: &str, _buttons: i64, _title: &str) -> Result<i64, VbaError> {
        self.messages.push(prompt.to_string());
        Ok(6)
    }
    fn inputbox(&mut self, _p: &str, _t: &str, _d: &str) -> Result<Option<String>, VbaError> {
        Ok(self.input.clone())
    }
    fn debug_print(&mut self, text: &str) {
        self.prints.push(text.to_string());
    }
}

fn run(src: &str, entry: &str) -> (Result<Value, VbaError>, TestHost) {
    let mut host = TestHost::default();
    let r = crate::run(&[("Module1".to_string(), src.to_string())], entry, vec![], &mut host);
    (r, host)
}

fn val(src: &str) -> Value {
    run(&format!("Function F()\r\n{src}\r\nEnd Function"), "F").0.unwrap()
}

#[test]
fn expressions_and_variables() {
    assert_eq!(val("F = 2 + 3 * 4").to_i64().unwrap(), 14);
    assert_eq!(val("Dim x As Long\r\nx = 5\r\nF = x * x").to_i64().unwrap(), 25);
    assert_eq!(val("F = \"a\" & 1 & True").to_str().unwrap(), "a1True");
    assert_eq!(val("F = 10 \\ 3").to_i64().unwrap(), 3);
    assert_eq!(val("F = 10 Mod 3").to_i64().unwrap(), 1);
    assert_eq!(val("F = 2 ^ 10").to_i64().unwrap(), 1024);
    assert_eq!(val("If 3 > 2 And 1 < 2 Then F = \"yes\" Else F = \"no\"").to_str().unwrap(), "yes");
}

#[test]
fn control_flow() {
    let fac = "Function F()\r\n  Dim i As Long, p As Long\r\n  p = 1\r\n  For i = 1 To 5\r\n    p = p * i\r\n  Next i\r\n  F = p\r\nEnd Function";
    assert_eq!(run(fac, "F").0.unwrap().to_i64().unwrap(), 120);
    let sum = "Function F()\r\n  Dim i As Long, s As Long\r\n  i = 1\r\n  Do While i <= 100\r\n    s = s + i\r\n    i = i + 1\r\n  Loop\r\n  F = s\r\nEnd Function";
    assert_eq!(run(sum, "F").0.unwrap().to_i64().unwrap(), 5050);
    let sel = "Function F(ByVal n As Long)\r\n  Select Case n\r\n    Case 1, 2: F = \"low\"\r\n    Case 3 To 5: F = \"mid\"\r\n    Case Is > 5: F = \"high\"\r\n    Case Else: F = \"?\"\r\n  End Select\r\nEnd Function";
    let mut host = TestHost::default();
    assert_eq!(crate::run(&[("M".into(), sel.into())], "F", vec![Value::Integer(4)], &mut host).unwrap().to_str().unwrap(), "mid");
}

#[test]
fn recursion_and_calls() {
    let src = "Function Fib(ByVal n As Long) As Long\r\n  If n < 2 Then\r\n    Fib = n\r\n  Else\r\n    Fib = Fib(n - 1) + Fib(n - 2)\r\n  End If\r\nEnd Function\r\nFunction F()\r\n  F = Fib(10)\r\nEnd Function";
    assert_eq!(run(src, "F").0.unwrap().to_i64().unwrap(), 55);
}

#[test]
fn byref_arguments() {
    let src = "Sub Inc(ByRef x As Long)\r\n  x = x + 1\r\nEnd Sub\r\nFunction F()\r\n  Dim a As Long\r\n  a = 10\r\n  Inc a\r\n  Inc a\r\n  F = a\r\nEnd Function";
    assert_eq!(run(src, "F").0.unwrap().to_i64().unwrap(), 12);
}

#[test]
fn host_cells_and_with() {
    let src = r#"Sub Fill()
    Dim r As Long
    For r = 1 To 3
        Cells(r, 1).Value = r * 10
    Next r
End Sub"#;
    let (r, host) = run(src, "Fill");
    assert!(r.is_ok(), "{r:?}");
    assert_eq!(host.cells.get("1,1").unwrap().to_i64().unwrap(), 10);
    assert_eq!(host.cells.get("3,1").unwrap().to_i64().unwrap(), 30);
}

#[test]
fn strings_and_arrays() {
    assert_eq!(val("F = UCase(Left(\"hello\", 3))").to_str().unwrap(), "HEL");
    let arr = "Function F()\r\n  Dim a(1 To 3) As Long\r\n  a(1) = 10 : a(2) = 20 : a(3) = 30\r\n  F = a(1) + a(2) + a(3)\r\nEnd Function";
    assert_eq!(run(arr, "F").0.unwrap().to_i64().unwrap(), 60);
    let redim = "Function F()\r\n  Dim a() As Long\r\n  ReDim a(2)\r\n  a(0) = 1 : a(1) = 2\r\n  ReDim Preserve a(3)\r\n  a(3) = 9\r\n  F = a(0) + a(1) + a(3)\r\nEnd Function";
    assert_eq!(run(redim, "F").0.unwrap().to_i64().unwrap(), 12);
}

#[test]
fn error_handling() {
    let src = "Function F()\r\n  On Error GoTo Oops\r\n  Dim x As Long\r\n  x = 1 / 0\r\n  F = \"unreached\"\r\n  Exit Function\r\nOops:\r\n  F = \"caught \" & Err.Number\r\nEnd Function";
    assert_eq!(run(src, "F").0.unwrap().to_str().unwrap(), "caught 11");
    let resume = "Function F()\r\n  Dim s As Long\r\n  On Error Resume Next\r\n  s = 1 / 0\r\n  s = 5\r\n  F = s\r\nEnd Function";
    assert_eq!(run(resume, "F").0.unwrap().to_i64().unwrap(), 5);
}

#[test]
fn collection_and_dictionary_objects() {
    let col =
        "Function F()\r\n  Dim c As New Collection\r\n  c.Add \"x\"\r\n  c.Add \"y\", \"key\"\r\n  F = c.Count & c(1) & c(\"key\")\r\nEnd Function";
    assert_eq!(run(col, "F").0.unwrap().to_str().unwrap(), "2xy");
    let dict = "Function F()\r\n  Dim d As Object\r\n  Set d = CreateObject(\"Scripting.Dictionary\")\r\n  d(\"a\") = 1\r\n  d(\"a\") = d(\"a\") + 4\r\n  F = d(\"a\") & d.Count\r\nEnd Function";
    assert_eq!(run(dict, "F").0.unwrap().to_str().unwrap(), "51");
}

#[test]
fn msgbox_and_debug() {
    let src = "Sub T()\r\n  MsgBox \"Hola \" & (2 + 2)\r\n  Debug.Print \"x\"; 42\r\nEnd Sub";
    let (r, host) = run(src, "T");
    assert!(r.is_ok());
    assert_eq!(host.messages, vec!["Hola 4"]);
    assert_eq!(host.prints, vec!["x42"]);
}

#[test]
fn class_modules() {
    let class = ("Account".to_string(), "Private mBalance As Double\r\nPublic Sub Deposit(ByVal amt As Double)\r\n  mBalance = mBalance + amt\r\nEnd Sub\r\nPublic Property Get Balance() As Double\r\n  Balance = mBalance\r\nEnd Property".to_string());
    let main = (
        "Module1".to_string(),
        "Function F()\r\n  Dim a As New Account\r\n  a.Deposit 100\r\n  a.Deposit 50\r\n  F = a.Balance\r\nEnd Function".to_string(),
    );
    let mut host = TestHost::default();
    let r = crate::run(&[class, main], "F", vec![], &mut host).unwrap();
    assert_eq!(r.to_f64().unwrap(), 150.0);
}

#[test]
fn macros_are_listed() {
    let src = "Public Sub Visible()\r\nEnd Sub\r\nPrivate Sub Hidden()\r\nEnd Sub\r\nPublic Sub NeedsArg(x As Long)\r\nEnd Sub\r\nPublic Function NotAMacro()\r\nEnd Function";
    let program = Program::parse(&[("Module1".to_string(), src.to_string())]).unwrap();
    let mut host = TestHost::default();
    let names = Interp::new(program, &mut host).macros();
    assert_eq!(names, vec!["Module1.Visible"]);
}

#[test]
fn runaway_loops_are_stopped() {
    let src = "Sub T()\r\n  Do\r\n  Loop\r\nEnd Sub";
    let (r, _) = run(src, "T");
    assert_eq!(r.unwrap_err().number, 28);
}

#[test]
fn hostile_sources_never_panic() {
    for src in [
        "",
        "Sub",
        "Function F()\r\n F = ((((\r\nEnd Function",
        "Sub T()\r\n GoTo Nowhere\r\nEnd Sub",
        "Sub T()\r\n Dim a(1)\r\n a(99) = 1\r\nEnd Sub",
        "Sub T()\r\n x = 1 / 0\r\nEnd Sub",
        "Sub T()\r\n Set x = Nothing\r\n x.Foo\r\nEnd Sub",
    ] {
        let mut host = TestHost::default();
        let _ = crate::run(&[("M".into(), src.into())], "T", vec![], &mut host);
    }
}

/// Runs `src` on a thread with a 2 MiB stack, as small as a thread usually gets.
fn run_on_small_stack(src: String) -> Result<String, VbaError> {
    std::thread::Builder::new()
        .stack_size(2 * 1024 * 1024)
        .spawn(move || {
            let mut host = TestHost::default();
            crate::run(&[("M".into(), src)], "T", vec![], &mut host).map(|v| v.to_str().unwrap_or_default())
        })
        .unwrap()
        .join()
        .map_err(|_| err(51, "the thread died"))
        .and_then(|r| r)
}

#[test]
fn deep_nesting_is_refused_not_a_crash() {
    // These overflowed the stack and killed the app.
    let parens = format!("Function T()\r\n T = {}1{}\r\nEnd Function", "(".repeat(100_000), ")".repeat(100_000));
    let ifs = format!("Function T()\r\n{}T = 1\r\n{}End Function", "If 1 Then\r\n".repeat(20_000), "End If\r\n".repeat(20_000));
    let nots = format!("Function T()\r\n T = {}True\r\nEnd Function", "Not ".repeat(100_000));
    for src in [parens, ifs, nots] {
        let e = run_on_small_stack(src).unwrap_err();
        assert!(e.description.contains("too complex"), "{e:?}");
    }
    // Recursion through nested blocks stops with "Out of stack space".
    let body = format!("{}F = F(n - 1) + 1\r\n{}", "If n > 0 Then\r\n".repeat(120), "End If\r\n".repeat(120));
    let e = run_on_small_stack(format!("Function F(n)\r\n{body}End Function\r\nFunction T()\r\n T = F(250)\r\nEnd Function")).unwrap_err();
    assert_eq!(e.number, 28);
    // Shallow recursion still runs.
    let r = run_on_small_stack(
        "Function F(n)\r\nIf n = 0 Then\r\nF = 0\r\nElse\r\nF = F(n - 1) + 1\r\nEnd If\r\nEnd Function\r\nFunction T()\r\n T = F(20)\r\nEnd Function"
            .into(),
    );
    assert_eq!(r.unwrap(), "20");
}

#[test]
fn huge_strings_and_arrays_are_out_of_memory() {
    for src in [
        "Function T()\r\n T = Space(1000000000000)\r\nEnd Function",
        "Function T()\r\n T = String(1000000000000, \"x\")\r\nEnd Function",
        "Function T()\r\n s = \"ab\"\r\n For i = 1 To 60\r\n  s = s & s\r\n Next\r\n T = s\r\nEnd Function",
        "Function T()\r\n s = String(100000, \"a\")\r\n T = Replace(s, \"a\", String(100000, \"b\"))\r\nEnd Function",
        "Function T()\r\n Dim a()\r\n ReDim a(1 To 100000000)\r\nEnd Function",
    ] {
        let mut host = TestHost::default();
        let e = crate::run(&[("M".into(), src.into())], "T", vec![], &mut host).unwrap_err();
        assert_eq!(e.number, 7, "{src}: {e:?}");
    }
    // Bounds that overflow are an error, not a panic.
    let mut host = TestHost::default();
    let r = crate::run(
        &[("M".into(), "Function T()\r\n Dim a()\r\n ReDim a(-9223372036854775807 To 9223372036854775807)\r\nEnd Function".into())],
        "T",
        vec![],
        &mut host,
    );
    assert!(r.is_err());
    assert_eq!(val("F = Len(Space(3) & String(2, \"x\"))").to_i64().unwrap(), 5);
}

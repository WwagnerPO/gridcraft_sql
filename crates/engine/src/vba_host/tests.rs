use super::*;
use gridcraft_core::Value as CV;

/// A session with one workbook holding `code` in Module1.
fn session(code: &str) -> Session {
    let mut s = Session::new();
    s.execute("file.new", json!({})).unwrap();
    let bin = gridcraft_xlsx::ovba::tests_support::sample_project(&[("Module1", true, code)]);
    std::sync::Arc::make_mut(&mut s.doc_mut().unwrap().wb).vba_project = Some(bin);
    s
}

fn cell(s: &Session, a1: &str) -> CV {
    s.doc().unwrap().wb.active().unwrap().value(CellRef::parse(a1).unwrap())
}

fn run(s: &mut Session, name: &str) -> Json {
    s.execute("vba.run", json!({"macro": name})).unwrap()
}

#[test]
fn writes_and_reads_cells() {
    let mut s = session(
        r#"Attribute VB_Name = "Module1"
Sub Fill()
    Dim i As Long
    For i = 1 To 5
        Cells(i, 1).Value = i * 10
    Next i
    Range("B1").Value = "Total"
    Range("B2").Formula = "=SUM(A1:A5)"
    Range("C1").Value = Range("A3").Value + 1
End Sub"#,
    );
    run(&mut s, "Fill");
    assert_eq!(cell(&s, "A5"), CV::Number(50.0));
    assert_eq!(cell(&s, "B1"), CV::text("Total"));
    assert_eq!(cell(&s, "B2"), CV::Number(150.0));
    assert_eq!(cell(&s, "C1"), CV::Number(31.0));
    // One undo step undoes the whole macro.
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(cell(&s, "A1"), CV::Empty);
    assert_eq!(cell(&s, "B2"), CV::Empty);
}

#[test]
fn ranges_navigate_like_excel() {
    let mut s = session(
        r#"Sub Nav()
    Range("A1:A4").Value = 1
    Dim last As Long
    last = Range("A1").End(xlDown).Row
    Range("B1").Value = last
    Range("B2").Value = Range("A1").CurrentRegion.Address
    Range("A1").Offset(5, 1).Value = "off"
    Range("B3").Value = Range("A1:C3").Count
    Range("B4").Value = Range("A1").Resize(2, 3).Address(False, False)
    Range("A10").Value = Cells(Rows.Count, 1).End(xlUp).Row
End Sub"#,
    );
    run(&mut s, "Nav");
    assert_eq!(cell(&s, "B1"), CV::Number(4.0));
    assert_eq!(cell(&s, "B2"), CV::text("$A$1:$B$4"));
    assert_eq!(cell(&s, "B6"), CV::text("off"));
    assert_eq!(cell(&s, "B3"), CV::Number(9.0));
    assert_eq!(cell(&s, "B4"), CV::text("A1:C2"));
    assert_eq!(cell(&s, "A10"), CV::Number(4.0));
}

#[test]
fn formatting() {
    let mut s = session(
        r#"Sub Fmt()
    With Range("A1:B2")
        .Value = 5
        .Font.Bold = True
        .Font.Color = RGB(255, 0, 0)
        .Interior.Color = vbYellow
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
    End With
End Sub"#,
    );
    run(&mut s, "Fmt");
    let d = s.doc().unwrap();
    let sh = d.wb.active().unwrap();
    let st = d.wb.styles.get(sh.style_id(CellRef::parse("B2").unwrap()));
    assert!(st.font.bold);
    assert_eq!(st.font.color, Color::Rgb(0xFF0000));
    assert_eq!(st.fill.fg, Color::Rgb(0xFFFF00));
    assert_eq!(st.num_fmt.0.as_ref(), "0.00");
    assert_eq!(st.align.h, HAlign::Center);
}

#[test]
fn worksheets() {
    let mut s = session(
        r#"Sub Sheets_()
    Dim ws As Worksheet
    Set ws = Worksheets.Add
    ws.Name = "Datos"
    ws.Range("A1").Value = "hola"
    Worksheets("Datos").Range("A2").Value = Worksheets.Count
    Dim n As String
    For Each ws In Worksheets
        n = n & ws.Name & ";"
    Next ws
    Sheets("Sheet1").Range("A1").Value = n
End Sub"#,
    );
    run(&mut s, "Sheets_");
    let wb = &s.doc().unwrap().wb;
    let datos = wb.sheet(wb.sheet_index("Datos").unwrap()).unwrap();
    assert_eq!(datos.value(CellRef::parse("A1").unwrap()), CV::text("hola"));
    assert_eq!(datos.value(CellRef::parse("A2").unwrap()), CV::Number(2.0));
    let s1 = wb.sheet(wb.sheet_index("Sheet1").unwrap()).unwrap();
    assert!(s1.value(CellRef::parse("A1").unwrap()).display().contains("Datos;"));
}

#[test]
fn arrays_in_and_out() {
    let mut s = session(
        r#"Sub Arr()
    Dim a As Variant
    Range("A1:B2").Value = Array(1, 2)
    a = Range("A1:B2").Value
    Range("D1").Value = a(1, 2) + a(2, 1)
    Range("D2").Value = UBound(a, 1) & "x" & UBound(a, 2)
End Sub"#,
    );
    run(&mut s, "Arr");
    assert_eq!(cell(&s, "D1"), CV::Number(3.0));
    assert_eq!(cell(&s, "D2"), CV::text("2x2"));
}

#[test]
fn worksheet_function_and_messages() {
    let mut s = session(
        r#"Sub WF()
    Range("A1:A3").Value = 4
    Range("B1").Value = Application.WorksheetFunction.Sum(Range("A1:A3"))
    Range("B2").Value = WorksheetFunction.Max(3, 9, 2)
    If MsgBox("¿Seguir?", vbYesNo) = vbYes Then Range("B3").Value = "sí"
    Debug.Print "listo"
End Sub"#,
    );
    let r = run(&mut s, "WF");
    assert_eq!(cell(&s, "B1"), CV::Number(12.0));
    assert_eq!(cell(&s, "B2"), CV::Number(9.0));
    assert_eq!(cell(&s, "B3"), CV::text("sí"));
    assert_eq!(r["messages"][0], "¿Seguir?");
    assert_eq!(r["output"][0], "listo");
}

#[test]
fn formula_r1c1() {
    assert_eq!(r1c1_formula_to_a1("=R[-1]C+RC[1]", CellRef::parse("B3").unwrap()), "=B2+C3");
    assert_eq!(r1c1_formula_to_a1("=SUM(R1C1:R2C2)", CellRef::parse("D4").unwrap()), "=SUM($A$1:$B$2)");
    assert_eq!(r1c1_formula_to_a1("=ROUND(R[-1]C,2)&\"RC\"", CellRef::parse("A2").unwrap()), "=ROUND(A1,2)&\"RC\"");
    let mut s = session(
        r#"Sub R1()
    Range("A1:A3").Value = 2
    Range("B1:B3").FormulaR1C1 = "=RC[-1]*10"
End Sub"#,
    );
    run(&mut s, "R1");
    assert_eq!(cell(&s, "B3"), CV::Number(20.0));
}

#[test]
fn macros_list_and_errors() {
    let mut s = session("Public Sub A()\r\nEnd Sub\r\nPrivate Sub B()\r\nEnd Sub\r\nSub Bad()\r\n Range(\"ZZZ\").Value = 1\r\nEnd Sub");
    let list = s.execute("vba.macros", json!({})).unwrap();
    assert!(list.as_array().unwrap().iter().any(|m| m == "Module1.A"));
    let e = s.execute("vba.run", json!({"macro": "Bad"})).unwrap_err().to_string();
    assert!(e.contains("1004"), "{e}");
    assert!(s.execute("vba.run", json!({"macro": "Missing"})).is_err());
    let mut bare = Session::new();
    bare.execute("file.new", json!({})).unwrap();
    assert!(bare.execute("vba.run", json!({"macro": "A"})).unwrap_err().to_string().contains("no macros"));
}

#[test]
fn hostile_macros_never_panic() {
    for code in [
        "Sub T()\r\n Cells(0, 0).Value = 1\r\nEnd Sub",
        "Sub T()\r\n Range(\"A1\").Offset(-5, 0).Value = 1\r\nEnd Sub",
        "Sub T()\r\n Range(\"A:A\").Value = 1\r\nEnd Sub",
        "Sub T()\r\n Worksheets(99).Activate\r\nEnd Sub",
        "Sub T()\r\n Range(\"A1\").Resize(0, 0).Value = 1\r\nEnd Sub",
        "Sub T()\r\n ActiveSheet.Name = \"\"\r\nEnd Sub",
    ] {
        let mut s = session(code);
        let r = s.execute("vba.run", json!({"macro": "T"}));
        assert!(!matches!(r, Err(crate::EngineError::Internal(..))), "{code}: {r:?}");
    }
}

#[test]
fn a_session_can_run_on_a_worker_thread() {
    fn assert_send<T: Send>() {}
    assert_send::<Session>();
}

/// A sheet with form buttons that run `Stamp`, like an approval sheet: the macro finds its own
/// button, comments the button's row, and relabels the button.
fn button_session(protect: bool) -> Session {
    let code = r#"Attribute VB_Name = "Module1"
Sub Stamp()
    Dim ws As Worksheet
    Dim btn As Shape
    Dim cell As Range
    Set ws = ActiveSheet
    Set btn = ws.Shapes(Application.Caller)
    Set cell = btn.TopLeftCell
    On Error Resume Next
    cell.CommentThreaded.Delete
    On Error GoTo 0
    If ws.Cells(cell.Row, 6).Text = "" Then
        btn.TextFrame.Characters.Text = "Fill F first"
        Exit Sub
    End If
    cell.AddCommentThreaded "Approved by " & Application.UserName
    btn.TextFrame.Characters.Text = "Approved"
End Sub
"#;
    let mut s = session(code);
    let p = gridcraft_xlsx::form_controls::sample_preserved(&[
        (1025, "Button 1", "Click to approve", "Stamp", [1, 5, 2, 3, 3, 8, 2, 18]),
        (1026, "Button 2", "Click to approve", "Stamp", [1, 5, 3, 3, 3, 8, 3, 18]),
    ]);
    {
        let wb = std::sync::Arc::make_mut(&mut s.doc_mut().unwrap().wb);
        let sh = wb.sheet_mut(0).unwrap();
        sh.preserved = Some(std::sync::Arc::new(p));
    }
    s.execute("cell.set", json!({"cell": "F4", "input": "INV-1"})).unwrap();
    if protect {
        s.execute("review.protectSheet", json!({"password": "x"})).unwrap();
    }
    s
}

#[test]
fn buttons_run_their_macro_with_application_caller() {
    let mut s = button_session(false);
    let list = s.execute("vba.buttons", json!({})).unwrap();
    assert_eq!(list[0]["name"], "Button 1");
    assert_eq!(list[0]["macro"], "Stamp");
    assert_eq!(list[1]["range"], "B4:D4");
    // Row 3 (Button 1) has no F value: the button says so, no comment.
    s.execute("vba.click", json!({"button": "Button 1"})).unwrap();
    let buttons = s.execute("vba.buttons", json!({})).unwrap();
    assert_eq!(buttons[0]["caption"], "Fill F first");
    // Row 4 (Button 2) has one: a threaded comment on B4 and the button relabelled.
    s.execute("vba.click", json!({"button": "Button 2"})).unwrap();
    let d = s.doc().unwrap();
    let c = d.wb.active().unwrap().comments.get(&CellRef::parse("B4").unwrap()).cloned().unwrap();
    assert!(c.threaded);
    assert!(c.text.starts_with("Approved by "), "{}", c.text);
    let buttons = s.execute("vba.buttons", json!({})).unwrap();
    assert_eq!(buttons[1]["caption"], "Approved");
}

#[test]
fn buttons_on_a_protected_sheet() {
    let mut s = button_session(true);
    s.execute("vba.click", json!({"button": "Button 2"})).unwrap();
    let buttons = s.execute("vba.buttons", json!({})).unwrap();
    assert_eq!(buttons[1]["caption"], "Approved");
}

#[test]
fn a_button_with_no_macro_and_a_missing_button() {
    let mut s = button_session(false);
    assert!(s.execute("vba.click", json!({"button": "Button 99"})).is_err());
    // Run without a button: Application.Caller is an error value, Shapes() fails cleanly.
    let e = s.execute("vba.run", json!({"macro": "Stamp"})).unwrap_err().to_string();
    assert!(e.contains("13") || e.contains("1004"), "{e}");
}

#[test]
fn files_through_adodb_stream_dir_and_kill() {
    let dir = std::env::temp_dir().join(format!("gridcraft-vba-{}", std::process::id()));
    let _ = std::fs::create_dir_all(&dir);
    let path = dir.join("out.csv");
    let path_s = path.to_string_lossy().replace('"', "\"\"");
    let code = format!(
        r#"Sub Csv()
    Dim f As String
    f = "{path_s}"
    Range("A1:A3").Value = "x"
    Range("A2").Value = "año"
    If Len(Dir(f)) > 0 Then Kill f
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    Dim r As Long
    For r = 1 To 3
        If Len(ActiveSheet.Cells(r, 1)) > 0 Then st.WriteText ActiveSheet.Cells(r, 1).Value & vbCrLf
    Next r
    st.SaveToFile f, 2
    Range("B1").Value = Dir(f)
End Sub"#
    );
    let mut s = session(&code);
    std::fs::write(&path, b"old").unwrap();
    run(&mut s, "Csv");
    let bytes = std::fs::read(&path).unwrap();
    assert_eq!(bytes, "\u{feff}x\r\naño\r\nx\r\n".as_bytes());
    assert_eq!(cell(&s, "B1"), CV::text("out.csv"));
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn default_member_values() {
    let mut s = session(
        r#"Sub D()
    Dim folder As String
    Range("P3").Value = "C:\data"
    folder = Range("P3")
    Range("A1").Value = folder & "\" & Len(Range("P3"))
    Range("A2") = Range("P3")
    If Range("P3") = "C:\data" Then Range("A3").Value = "same"
End Sub"#,
    );
    run(&mut s, "D");
    assert_eq!(cell(&s, "A1"), CV::text("C:\\data\\7"));
    assert_eq!(cell(&s, "A2"), CV::text("C:\\data"));
    assert_eq!(cell(&s, "A3"), CV::text("same"));
}

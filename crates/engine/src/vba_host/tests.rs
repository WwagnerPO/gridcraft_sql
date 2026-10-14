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

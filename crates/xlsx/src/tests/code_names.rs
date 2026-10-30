//! Sheet code names (`<sheetPr codeName>`) tie a sheet to its VBA module: they survive saving,
//! and a copied sheet's duplicate is left for Excel to assign anew.

use gridcraft_model::{Sheet, Workbook};

use crate::{read_xlsx, write_xlsx};

#[test]
fn code_names_round_trip() {
    let mut wb = Workbook::new();
    let mut first = Sheet::new("Ventas");
    first.code_name = Some("Hoja1".into());
    let mut copy = Sheet::new("Ventas (2)");
    copy.code_name = Some("Hoja1".into());
    let mut third = Sheet::new("Notas");
    third.code_name = Some("Hoja3 & <co>".into());
    wb.sheets = vec![first.into(), copy.into(), third.into()];
    let bytes = write_xlsx(&wb).unwrap();
    let (back, _) = read_xlsx(&bytes).unwrap();
    let names: Vec<Option<String>> = back.sheets.iter().map(|s| s.code_name.clone()).collect();
    assert_eq!(names, vec![Some("Hoja1".into()), None, Some("Hoja3 & <co>".into())]);
}

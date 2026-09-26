//! Structured references in formulas: Excel table references such as
//! `Sales[Amount]`, `Sales[@Amount]`, `Sales[[#All],[Amount]]` (ECMA-376
//! Part 1 §18.17.2.5). A table name resolves to its `xl/tables/tableN.xml`
//! range, and item specifiers pick the row band (data / headers / totals /
//! all) and the column slice.
//!
//! Resolution strategy: IronCalc 0.8.3 already resolves structured
//! references natively at parse time — its importer loads every workbook
//! table and its parser rewrites `Table[...]` tokens into plain range nodes
//! (including cross-sheet qualification, `#Totals` semantics and escaped
//! column names), so no second resolver is maintained here. The one gap in
//! 0.8.3 is the `@` this-row shorthand Excel writes by default
//! (`Sales[@Amount]`, `Sales[@[Unit Price]]`, `Sales[@[Jan]:[Dec]]`): its
//! lexer only accepts the spelled-out `[#This Row]` form. `normalize_at_shorthand`
//! rewrites those formulas into the spelled-out form before they reach
//! IronCalc — on the recalc side only. Stored formulas, the renderer model
//! and saved files always keep the user-facing `@` text, so files round-trip
//! exactly as Excel wrote them.

use std::borrow::Cow;

use ironcalc::base::Model;
use ironcalc::base::types::Cell;

/// Rewrites Excel's `@` this-row shorthand inside structured references into
/// the spelled-out `[#This Row]` form IronCalc 0.8.3 lexes:
///
/// - `Sales[@Amount]`            → `Sales[[#This Row],[Amount]]`
/// - `Sales[@[Unit Price]]`      → `Sales[[#This Row],[Unit Price]]`
/// - `Sales[@[Jan]:[Dec]]`       → `Sales[[#This Row],[Jan]:[Dec]]`
/// - `Sales[@]`                  → `Sales[#This Row]`
///
/// Everything else passes through untouched: string literals (with `""`
/// escapes), `@` outside a `[...]` item (implicit intersection, plain text),
/// already-spelled `[#This Row]` forms, and bracketed items not preceded by a
/// table-name token (external-workbook references, error literals).
pub(crate) fn normalize_at_shorthand(formula: &str) -> Cow<'_, str> {
    if !formula.contains('@') {
        return Cow::Borrowed(formula);
    }
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len());
    let mut position = 0usize;
    let mut in_string = false;
    let mut rewritten = false;
    while position < chars.len() {
        let ch = chars[position];
        if in_string {
            out.push(ch);
            if ch == '"' {
                if chars.get(position + 1) == Some(&'"') {
                    // "" is an escaped quote inside the literal.
                    out.push('"');
                    position += 1;
                } else {
                    in_string = false;
                }
            }
            position += 1;
            continue;
        }
        match ch {
            '"' => {
                in_string = true;
                out.push(ch);
                position += 1;
            }
            '[' => {
                let Some(after_item) = scan_bracketed_item(&chars, position) else {
                    // Unterminated item: copy the rest verbatim.
                    out.push(ch);
                    position += 1;
                    continue;
                };
                let item: String = chars[position + 1..after_item - 1].iter().collect();
                if let (true, Some(replacement)) = (
                    ends_with_table_name(&chars, position),
                    at_item_replacement(&item),
                ) {
                    out.push_str(&replacement);
                    rewritten = true;
                    position = after_item;
                    continue;
                }
                // Not an `@` shorthand (or no table name in front): copy the
                // whole item verbatim. Its inner text cannot start a new
                // structured reference, so rescanning it is pointless.
                out.push_str(&item_copy(&chars, position, after_item));
                position = after_item;
            }
            _ => {
                out.push(ch);
                position += 1;
            }
        }
    }
    if rewritten {
        Cow::Owned(out)
    } else {
        Cow::Borrowed(formula)
    }
}

/// Whether the formula text right before `index` looks like a table name:
/// a run of name characters (letters, digits, `_`, `\`, `.` — table display
/// names cannot contain spaces or quotes) or the closing quote of a quoted
/// name. `[` with no token in front is not a structured reference.
fn ends_with_table_name(chars: &[char], index: usize) -> bool {
    if index == 0 {
        return false;
    }
    if chars[index - 1] == '\'' {
        return true;
    }
    let mut cursor = index;
    while cursor > 0 {
        let ch = chars[cursor - 1];
        if ch.is_alphanumeric() || matches!(ch, '_' | '\\' | '.') {
            cursor -= 1;
        } else {
            break;
        }
    }
    cursor < index
}

/// Scans the `]`-terminated item whose `[` sits at `open`: returns the index
/// just past the closing `]`, honoring the `']` / `''` escapes of bracketed
/// column names. `None` when the formula ends inside the item.
fn scan_bracketed_item(chars: &[char], open: usize) -> Option<usize> {
    let mut cursor = open + 1;
    while cursor < chars.len() {
        match chars[cursor] {
            ']' => return Some(cursor + 1),
            '\'' => cursor += 1,
            _ => {}
        }
        cursor += 1;
    }
    None
}

/// The replacement content for an `@`-shorthand item, brackets included;
/// `None` when the item does not start with the shorthand. IronCalc's lexer
/// requires a `[` right after the specifier comma, so a bare column gets
/// bracketed (`[@Amount]` → `[[#This Row],[Amount]]`); already-bracketed
/// items (`[@[Unit Price]]`, `[@[Jan]:[Dec]]`) pass through as written.
fn at_item_replacement(item: &str) -> Option<String> {
    let rest = item.strip_prefix('@')?;
    if rest.is_empty() {
        return Some("[#This Row]".to_string());
    }
    if rest.starts_with('[') {
        Some(format!("[[#This Row],{rest}]"))
    } else {
        Some(format!("[[#This Row],[{rest}]]"))
    }
}

fn item_copy(chars: &[char], open: usize, after_item: usize) -> String {
    chars[open..after_item].iter().collect()
}

/// Rewrites every plain formula cell of an imported IronCalc model whose
/// stored text uses the `@` shorthand. Runs before `pin_unparsable_formulas`
/// on the cold-import path, so formulas Excel wrote as `Sales[@Amount]`
/// resolve and evaluate instead of erroring (or being pinned to their cached
/// values). The rewrite reads the stored import text rather than the parsed
/// node: a broken structured reference re-parses into a name-like node whose
/// stringify no longer carries the original `[@...]` text. Array/CSE formulas
/// are left alone: `set_user_input` would drop their CSE/spill semantics.
/// Only the resident compute model changes — never the file.
pub(crate) fn normalize_model_structured_references(model: &mut Model) {
    let mut rewrites: Vec<(u32, i32, i32, String)> = Vec::new();
    for (sheet_index, worksheet) in model.workbook.worksheets.iter().enumerate() {
        for (row, columns) in &worksheet.sheet_data {
            for (column, cell) in columns {
                let Cell::CellFormula { f, .. } = cell else {
                    continue;
                };
                // Healthy structured references were resolved to plain ranges
                // at import and their stored text no longer contains `[`;
                // R1C1-relative texts (`R[1]C`) never contain an `@` item.
                let Some(text) = worksheet.shared_formulas.get(*f as usize) else {
                    continue;
                };
                if !text.contains('[') {
                    continue;
                }
                let Cow::Owned(normalized) = normalize_at_shorthand(text) else {
                    continue;
                };
                rewrites.push((sheet_index as u32, *row, *column, format!("={normalized}")));
            }
        }
    }
    for (sheet, row, column, input) in rewrites {
        // A rejected rewrite leaves the cell as imported; the pin step then
        // falls back to the file's cached value.
        let _ = model.set_user_input(sheet, row, column, input);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn normalize(formula: &str) -> String {
        normalize_at_shorthand(formula).into_owned()
    }

    #[test]
    fn plain_structured_reference_is_untouched() {
        assert_eq!(
            normalize("=SUM(Sales[Amount])"),
            "=SUM(Sales[Amount])".to_string()
        );
    }

    #[test]
    fn at_column_becomes_this_row() {
        assert_eq!(
            normalize("=Sales[@Amount]"),
            "=Sales[[#This Row],[Amount]]".to_string()
        );
    }

    #[test]
    fn at_bracketed_column_with_spaces_becomes_this_row() {
        assert_eq!(
            normalize("=SUM(Sales[@[Unit Price]])"),
            "=SUM(Sales[[#This Row],[Unit Price]])".to_string()
        );
    }

    #[test]
    fn at_column_range_becomes_this_row() {
        assert_eq!(
            normalize("=SUM(Sales[@[Jan]:[Dec]])"),
            "=SUM(Sales[[#This Row],[Jan]:[Dec]])".to_string()
        );
    }

    #[test]
    fn bare_at_becomes_this_row_specifier() {
        assert_eq!(normalize("=Sales[@]"), "=Sales[#This Row]".to_string());
    }

    #[test]
    fn at_shorthand_inside_aggregate() {
        assert_eq!(
            normalize("=SUM(Sales[@Amount]*2)+1"),
            "=SUM(Sales[[#This Row],[Amount]]*2)+1".to_string()
        );
    }

    #[test]
    fn multiple_tables_are_rewritten_independently() {
        assert_eq!(
            normalize("=SUM(Table1[@Qty],Table2[@[Unit Cost]])"),
            "=SUM(Table1[[#This Row],[Qty]],Table2[[#This Row],[Unit Cost]])".to_string()
        );
    }

    #[test]
    fn spelled_out_this_row_is_untouched() {
        assert_eq!(
            normalize("=SUM(Sales[[#This Row],[Amount]])"),
            "=SUM(Sales[[#This Row],[Amount]])".to_string()
        );
    }

    #[test]
    fn special_items_are_untouched() {
        assert_eq!(
            normalize("=COUNTA(Sales[#All])+COUNTA(Sales[[#Headers],[Amount]])+SUM(Sales[#Data])+SUM(Sales[[#Totals],[Amount]])"),
            "=COUNTA(Sales[#All])+COUNTA(Sales[[#Headers],[Amount]])+SUM(Sales[#Data])+SUM(Sales[[#Totals],[Amount]])".to_string()
        );
    }

    #[test]
    fn string_literals_keep_their_text() {
        assert_eq!(
            normalize("=IF(A1=\"[@not a ref]\",\"[@keep]\",0)"),
            "=IF(A1=\"[@not a ref]\",\"[@keep]\",0)".to_string()
        );
    }

    #[test]
    fn external_workbook_reference_is_untouched() {
        assert_eq!(
            normalize("=[1]Sheet1!A1+[@x]"),
            "=[1]Sheet1!A1+[@x]".to_string()
        );
    }

    #[test]
    fn at_without_table_name_is_untouched() {
        // Implicit intersection / bare item: no table-name token before `[`.
        assert_eq!(normalize("=SUM([@x])"), "=SUM([@x])".to_string());
        assert_eq!(normalize("=@A1:A5"), "=@A1:A5".to_string());
    }

    #[test]
    fn escaped_column_characters_survive_the_rewrite() {
        assert_eq!(
            normalize("=SUM(Sales[@[Column With ']Bracket]])"),
            "=SUM(Sales[[#This Row],[Column With ']Bracket]])".to_string()
        );
    }

    #[test]
    fn unicode_table_names_are_recognized() {
        assert_eq!(
            normalize("=Таблица1[@Сумма]"),
            "=Таблица1[[#This Row],[Сумма]]".to_string()
        );
    }

    #[test]
    fn formulas_without_at_are_rejected_fast() {
        assert!(matches!(
            normalize_at_shorthand("=SUM(Sales[Amount])"),
            Cow::Borrowed(_)
        ));
    }

    // -- Resolution through the recalc channel -------------------------------
    //
    // These tests exercise the production path end to end: a fixture workbook
    // with real table parts is imported by IronCalc (which loads the tables),
    // formulas enter through the same code the sidecar serves (cold file
    // import and user edits), and the evaluated numbers must match the
    // hand-computed sums.
    //
    // Reference data: Sales (A1:B6, totals row 6) Product/Amount with amounts
    // 10/20/30/40 and a totals row carrying 100; Prices (C1:C5) with the
    // escaped column name "Unit Price" and values 1.5/2.5/3.5/4.5. Both
    // tables live on Sheet1.

    use crate::CellRange;
    use crate::recalc::{RecalcCache, RecalcEdit, RecalcRead, recalc_cells};
    use std::fs::File;
    use std::io::Write;
    use std::path::Path;

    const WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;

    const WORKBOOK_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;

    const STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#;

    const SALES_TABLE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B6" totalsRowCount="1" headerRowCount="1"><autoFilter ref="A1:B5"/><tableColumns count="2"><tableColumn id="1" name="Product"/><tableColumn id="2" name="Amount"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>"#;

    const PRICES_TABLE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="Prices" displayName="Prices" ref="C1:C5" headerRowCount="1"><autoFilter ref="C1:C5"/><tableColumns count="1"><tableColumn id="1" name="Unit Price"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>"#;

    const SHEET_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/></Relationships>"#;

    /// `row2_cell`/`row3_cell` inject optional `<c>` elements at the end of
    /// rows 2/3 (the cold import test carries file-side formulas there).
    fn write_table_workbook(path: &Path, row2_cell: &str, row3_cell: &str) {
        let content_types = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>"#
        );
        let sheet = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C6"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Product</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Unit Price</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>alpha</t></is></c><c r="B2"><v>10</v></c><c r="C2"><v>1.5</v></c>{row2_cell}</row><row r="3"><c r="A3" t="inlineStr"><is><t>beta</t></is></c><c r="B3"><v>20</v></c><c r="C3"><v>2.5</v></c>{row3_cell}</row><row r="4"><c r="A4" t="inlineStr"><is><t>gamma</t></is></c><c r="B4"><v>30</v></c><c r="C4"><v>3.5</v></c></row><row r="5"><c r="A5" t="inlineStr"><is><t>delta</t></is></c><c r="B5"><v>40</v></c><c r="C5"><v>4.5</v></c></row><row r="6"><c r="A6" t="inlineStr"><is><t>Total</t></is></c><c r="B6"><v>100</v></c></row></sheetData><tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts></worksheet>"#
        );
        let dir = path.parent().expect("fixture parent");
        std::fs::create_dir_all(dir.join("xl/worksheets/_rels")).unwrap();
        std::fs::create_dir_all(dir.join("xl/tables")).unwrap();
        let mut writer = zip::ZipWriter::new(File::create(path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, content) in [
            ("[Content_Types].xml", content_types.as_str()),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            ("xl/workbook.xml", WORKBOOK_XML),
            ("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
            ("xl/styles.xml", STYLES_XML),
            ("xl/worksheets/sheet1.xml", sheet.as_str()),
            ("xl/worksheets/_rels/sheet1.xml.rels", SHEET_RELS),
            ("xl/tables/table1.xml", SALES_TABLE_XML),
            ("xl/tables/table2.xml", PRICES_TABLE_XML),
        ] {
            writer.start_file(name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
    }

    fn read_cell(sheet: &str, row: usize, column: usize) -> RecalcRead {
        RecalcRead {
            sheet: sheet.to_string(),
            range: CellRange {
                start_row: row,
                start_column: column,
                end_row: row,
                end_column: column,
            },
        }
    }

    fn recalc_number(path: &Path, edits: &[RecalcEdit], read: &RecalcRead) -> Option<f64> {
        let mut cache = RecalcCache::new();
        let result = recalc_cells(&mut cache, path, edits, std::slice::from_ref(read)).unwrap();
        assert_eq!(result.cells.len(), 1, "one read cell expected");
        result.cells[0].number
    }

    fn edit(row: u32, column: u32, input: &str) -> RecalcEdit {
        RecalcEdit {
            sheet: "Sheet1".to_string(),
            row,
            column,
            input: input.to_string(),
        }
    }

    #[test]
    fn recalc_resolves_the_reference_table_sums() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("tables.xlsx");
        write_table_workbook(&path, "", "");
        // Sales[Amount] = 10+20+30+40 (totals row excluded); the hand sum
        // matches the evaluated value.
        assert_eq!(
            recalc_number(
                &path,
                &[edit(7, 4, "=SUM(Sales[Amount])")],
                &read_cell("Sheet1", 7, 4)
            ),
            Some(100.0)
        );
    }

    #[test]
    fn recalc_resolves_this_row_and_special_items() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("tables.xlsx");
        write_table_workbook(&path, "", "");
        let mut cache = RecalcCache::new();
        let edits = [
            // E8: row 7 is not a table row, so plain band sums:
            edit(7, 4, "=SUM(Sales[Amount])+SUM(Sales[#Data])"),
            // E9: #All spans headers + data + totals; SUM skips the texts.
            edit(8, 4, "=SUM(Sales[[#All],[Amount]])"),
            // E10: header band has the two text headers.
            edit(9, 4, "=COUNTA(Sales[#Headers])+COUNTA(Sales[#Data])"),
            // E11: totals band only.
            edit(10, 4, "=SUM(Sales[[#Totals],[Amount]])"),
        ];
        let reads = [
            read_cell("Sheet1", 7, 4),
            read_cell("Sheet1", 8, 4),
            read_cell("Sheet1", 9, 4),
            read_cell("Sheet1", 10, 4),
        ];
        let result = recalc_cells(&mut cache, &path, &edits, &reads).unwrap();
        let number = |row: u32| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row)
                .and_then(|cell| cell.number)
        };
        // 100 + 100, 100 + 100 (totals included in #All), 2 + 8, 100.
        assert_eq!(number(7), Some(200.0));
        assert_eq!(number(8), Some(200.0));
        assert_eq!(number(9), Some(10.0));
        assert_eq!(number(10), Some(100.0));
    }

    #[test]
    fn recalc_resolves_at_shorthand_edits_and_escaped_columns() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("tables.xlsx");
        write_table_workbook(&path, "", "");
        let mut cache = RecalcCache::new();
        let edits = [
            // E2 sits in the first table row: @ picks B2/C2 of that row.
            edit(1, 4, "=Sales[@Amount]+Prices[@[Unit Price]]"),
            // E3: same row, spelled-out twin must agree.
            edit(2, 4, "=Sales[[#This Row],[Amount]]"),
            // E4: escaped column with spaces, whole-column sum.
            edit(3, 4, "=SUM(Prices[[Unit Price]])"),
        ];
        let reads = [
            read_cell("Sheet1", 1, 4),
            read_cell("Sheet1", 2, 4),
            read_cell("Sheet1", 3, 4),
        ];
        let result = recalc_cells(&mut cache, &path, &edits, &reads).unwrap();
        let number = |row: u32| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row)
                .and_then(|cell| cell.number)
        };
        assert_eq!(number(1), Some(11.5)); // 10 + 1.5
        assert_eq!(number(2), Some(20.0)); // E3's row: B3 = 20
        assert_eq!(number(3), Some(12.0)); // 1.5+2.5+3.5+4.5
    }

    #[test]
    fn cold_import_computes_file_formulas_with_at_shorthand() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("tables.xlsx");
        // E2/E3 come from the "Excel" file itself, cached at a stale 99:
        // without the rewrite E2 would stay pinned to that cache.
        write_table_workbook(
            &path,
            r#"<c r="E2"><f>Sales[@Amount]</f><v>99</v></c>"#,
            r#"<c r="E3"><f>SUM(Sales[Amount])</f><v>99</v></c>"#,
        );
        let mut cache = RecalcCache::new();
        let result = recalc_cells(
            &mut cache,
            &path,
            &[],
            &[read_cell("Sheet1", 1, 4), read_cell("Sheet1", 2, 4)],
        )
        .unwrap();
        let number = |row: u32| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row)
                .and_then(|cell| cell.number)
        };
        eprintln!("cells: {:?}", result.cells);
        assert_eq!(number(1), Some(10.0));
        assert_eq!(number(2), Some(100.0));
    }

    #[test]
    fn second_recalc_serves_consistent_values_from_the_resident_model() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("tables.xlsx");
        write_table_workbook(&path, "", "");
        let mut cache = RecalcCache::new();
        let edits = [edit(1, 4, "=Sales[@Amount]")];
        let reads = [read_cell("Sheet1", 1, 4)];
        let first = recalc_cells(&mut cache, &path, &edits, &reads).unwrap();
        assert_eq!(first.cells[0].number, Some(10.0));
        // The applied-edit cache keys on the ORIGINAL user input; the second
        // request must hit the resident model without changing the value.
        let second = recalc_cells(&mut cache, &path, &edits, &reads).unwrap();
        assert!(second.cached);
        assert_eq!(second.cells[0].number, Some(10.0));
    }
}

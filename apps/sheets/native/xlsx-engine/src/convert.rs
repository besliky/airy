//! Legacy-format import: reads an .xls (or anything calamine understands)
//! and writes a fresh .xlsx with values, formulas, date number formats and —
//! for BIFF8 books — the layout that makes a form look like a form: merged
//! ranges, column widths, row heights and hidden rows, and basic cell
//! styles (bold/size, fills, borders, alignments, number formats) via
//! `xls_layout` (BUG-1602, BUG-1607). The converted file is a fresh
//! workbook, not a byte-preserving edit.

use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;

use calamine::{Data, Reader, open_workbook_auto};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::SidecarError;
use crate::legacy_xls::{self, SheetStrings};
use crate::ole2::validate_ole2_header;
use crate::xls_layout::{self, StyleInterner};

#[derive(Debug)]
pub struct ConvertResult {
    pub sheets: usize,
    pub cells: usize,
    /// Raw on-disk byte length of the conversion source, measured when the
    /// conversion starts (BUG-1305): the session re-runs its open-size fence
    /// against the size actually served even when the file grew past the cap
    /// between the host's stat and this read.
    pub source_bytes: u64,
}

/// SEC-1301: calamine reads .ods — and every other zip-based workbook its
/// auto-detection accepts — through the same ZIP container as .xlsx, so the
/// SEC-1103 central-directory fence must cover the convert path too: entry
/// count, canonical entry paths, and the declared-uncompressed total, all
/// read from metadata before calamine decompresses a single entry. The
/// check is content-based, not extension-based, because the host dispatches
/// by extension while calamine itself falls back to content sniffing. A
/// source without the zip magic (legacy .xls is an OLE2 compound document,
/// no central directory to walk) or one the zip crate cannot even open
/// skips the fence and fails inside calamine with its own message.
fn validate_zip_based_source(source: &Path) -> Result<(), SidecarError> {
    let mut magic = [0u8; 2];
    let starts_with_zip_magic = File::open(source)?
        .read_exact(&mut magic)
        .is_ok_and(|()| &magic == b"PK");
    if !starts_with_zip_magic {
        return Ok(());
    }
    let mut archive = match ZipArchive::new(File::open(source)?) {
        Ok(archive) => archive,
        // Not a readable zip: calamine reports the failure itself, keeping
        // the pre-fence error text for corrupt inputs.
        Err(_) => return Ok(()),
    };
    crate::archive::validate_entries(&mut archive)
}

pub fn convert_to_xlsx(source: &Path, target: &Path) -> Result<ConvertResult, SidecarError> {
    let source_bytes = std::fs::metadata(source)
        .map(|metadata| metadata.len())
        .unwrap_or(0);
    validate_zip_based_source(source)?;
    // BUG-1606: refuse truncated compound-file headers before calamine
    // touches them — its CFB reader panics (index out of bounds) on that
    // class instead of returning an error, and the panic kills the sidecar.
    validate_ole2_header(source)?;
    let mut workbook =
        contain_calamine_panic(|| open_workbook_auto(source)).and_then(|opened| {
            opened.map_err(|error| {
                SidecarError::Workbook(format!("Unable to read the workbook: {error}"))
            })
        })?;
    // BUG-1600: calamine decodes legacy .xls strings through the workbook
    // codepage even when a string's own fHighByte flag selects UTF-16, so
    // single-byte-codepage books come out as mojibake. The overlay is the
    // spec-correct re-decode of the SST; empty for every other format.
    let repaired = legacy_xls::LegacyXlsStrings::extract(source);
    // BUG-1602: merges, column widths and basic cell styles, walked from the
    // BIFF stream directly (calamine exposes none of them). Empty for every
    // non-.xls source, which keeps those conversions unchanged.
    let layout = xls_layout::WorkbookLayout::extract(source);
    let mut styler = StyleInterner::new(&layout);
    let names: Vec<String> = workbook.sheet_names().to_vec();
    if names.is_empty() {
        return Err(SidecarError::Workbook("The workbook has no sheets.".into()));
    }

    let mut sheet_xmls: Vec<String> = Vec::new();
    let mut cells = 0usize;
    for (index, name) in names.iter().enumerate() {
        let range =
            contain_calamine_panic(|| workbook.worksheet_range(name)).and_then(|range| {
                range.map_err(|error| SidecarError::Workbook(format!("Sheet {name}: {error}")))
            })?;
        let formulas = contain_calamine_panic(|| workbook.worksheet_formula(name))
            .ok()
            .and_then(|formulas| formulas.ok());
        let mut formula_map: HashMap<(u32, u32), String> = HashMap::new();
        if let Some(formula_range) = formulas {
            let (start_row, start_col) = formula_range.start().unwrap_or((0, 0));
            for (row, column, formula) in formula_range.used_cells() {
                if !formula.is_empty() {
                    formula_map.insert(
                        (start_row + row as u32, start_col + column as u32),
                        formula.clone(),
                    );
                }
            }
        }
        let (xml, sheet_cells) = worksheet_xml(
            &range,
            &formula_map,
            repaired.sheet(index),
            layout.sheet(index),
            &mut styler,
        );
        cells += sheet_cells;
        sheet_xmls.push(xml);
    }

    let file = File::create(target)?;
    let mut writer = ZipWriter::new(file);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    let mut add = |name: &str, content: &str| -> Result<(), SidecarError> {
        writer.start_file(name, options)?;
        writer.write_all(content.as_bytes())?;
        Ok(())
    };

    add("[Content_Types].xml", &content_types_xml(names.len()))?;
    add("_rels/.rels", ROOT_RELS)?;
    add("xl/workbook.xml", &workbook_xml(&names, &repaired))?;
    add(
        "xl/_rels/workbook.xml.rels",
        &workbook_rels_xml(names.len()),
    )?;
    // Styles last of the parts that reference them: cell styles are
    // interned while the sheets serialize, so styles.xml can only be built
    // once every sheet has been walked.
    add("xl/styles.xml", &styler.styles_xml())?;
    for (index, xml) in sheet_xmls.iter().enumerate() {
        add(&format!("xl/worksheets/sheet{}.xml", index + 1), xml)?;
    }
    writer.finish()?.sync_all()?;
    Ok(ConvertResult {
        sheets: names.len(),
        cells,
        source_bytes,
    })
}

/// BUG-1606: calamine panics instead of returning an error on some
/// malformed compound files (slice indexing in its CFB reader, and any
/// future panic on hostile input), and a panic takes the whole sidecar
/// down with it. The header fence in `ole2` refuses the truncated-header
/// class before calamine opens the file; this wrapper is the safety net
/// for anything that slips past it — every calamine call on the convert
/// path degrades to a normal workbook error. Mirrors the recalc path's
/// containment of IronCalc panics.
fn contain_calamine_panic<T>(read: impl FnOnce() -> T) -> Result<T, SidecarError> {
    std::panic::catch_unwind(std::panic::AssertUnwindSafe(read)).map_err(|payload| {
        let detail = payload
            .downcast_ref::<&str>()
            .map(|text| (*text).to_owned())
            .or_else(|| payload.downcast_ref::<String>().cloned())
            .unwrap_or_else(|| "unexpected internal failure".to_owned());
        SidecarError::Workbook(format!(
            "Unable to read the workbook: malformed file stopped the reader ({detail})."
        ))
    })
}

fn worksheet_xml(
    range: &calamine::Range<Data>,
    formulas: &HashMap<(u32, u32), String>,
    repaired: Option<&SheetStrings>,
    layout: Option<&xls_layout::SheetLayout>,
    styler: &mut StyleInterner<'_>,
) -> (String, usize) {
    // Rows carrying values plus rows carrying only formulas.
    let mut rows: HashMap<u32, Vec<(u32, String)>> = HashMap::new();
    let mut cells = 0usize;
    let (start_row, start_col) = range.start().unwrap_or((0, 0));
    let position_is_written = |rows: &HashMap<u32, Vec<(u32, String)>>, position: (u32, u32)| {
        rows.get(&position.0)
            .is_some_and(|line| line.iter().any(|(column, _)| *column == position.1))
    };
    for (row, column, value) in range.used_cells() {
        let absolute = (start_row + row as u32, start_col + column as u32);
        // BUG-1600: the overlay carries the spec-correct SST decode for
        // cells whose string calamine mangled through the book codepage.
        let formula = formulas.get(&absolute).map(String::as_str);
        // BUG-1602: cells with a walked style reference their interned xf;
        // plain cells keep the fallback (dates stay date-formatted).
        let style = layout
            .and_then(|layout| layout.cells.get(&absolute))
            .and_then(|&xf| styler.cell_style(xf));
        let outcome = match repaired.and_then(|sheet| sheet.cells.get(&absolute)) {
            Some(text) => cell_xml(absolute, &Data::String(text.clone()), formula, style),
            None => cell_xml(absolute, value, formula, style),
        };
        if let Some(cell) = outcome {
            cells += 1;
            rows.entry(absolute.0).or_default().push((absolute.1, cell));
        }
    }
    for (position, formula) in formulas {
        let covered = range
            .get_value((position.0, position.1))
            .is_some_and(|v| *v != Data::Empty);
        if !covered {
            cells += 1;
            rows.entry(position.0).or_default().push((
                position.1,
                format!(
                    r#"<c r="{}"><f>{}</f></c>"#,
                    cell_reference(position.0, position.1),
                    escape_xml(formula),
                ),
            ));
        }
    }
    // BUG-1602: styled empty cells — merge continuations and blank painted/
    // bordered boxes the value walk above never sees. Without them the
    // borders of a merged header die with the anchor row.
    if let Some(layout) = layout {
        for (&position, &xf) in layout.cells.iter() {
            if position_is_written(&rows, position) {
                continue;
            }
            let Some(style) = styler.cell_style(xf) else {
                continue;
            };
            cells += 1;
            rows.entry(position.0).or_default().push((
                position.1,
                format!(
                    r#"<c r="{}" s="{}"/>"#,
                    cell_reference(position.0, position.1),
                    style,
                ),
            ));
        }
    }

    // BUG-1607: rows the walk found interesting but that carry no cells —
    // a hidden row, a sized spacer — still need their <row> element.
    let mut row_numbers: Vec<u32> = rows.keys().copied().collect();
    if let Some(layout) = layout {
        for &(row, _) in &layout.rows {
            if !rows.contains_key(&row) {
                row_numbers.push(row);
            }
        }
    }
    row_numbers.sort_unstable();
    let mut body = String::new();
    for row in row_numbers {
        // BUG-1607: carried heights and hidden flags become row attributes.
        let attributes = layout
            .and_then(|layout| layout.row_spec(row))
            .map(row_attributes)
            .unwrap_or_default();
        let mut line = rows.remove(&row).unwrap_or_default();
        line.sort_unstable_by_key(|(column, _)| *column);
        if line.is_empty() {
            body.push_str(&format!(r#"<row r="{}"{attributes}/>"#, row + 1));
            continue;
        }
        body.push_str(&format!(r#"<row r="{}"{attributes}>"#, row + 1));
        for (_, cell) in line {
            body.push_str(&cell);
        }
        body.push_str("</row>");
    }

    let dimension = match (range.start(), range.end()) {
        (Some(start), Some(end)) => format!(
            "{}:{}",
            cell_reference(start.0, start.1),
            cell_reference(end.0, end.1),
        ),
        _ => "A1:A1".into(),
    };
    // Schema order: dimension, sheetFormatPr, cols, sheetData, mergeCells.
    let format_pr = layout.map(sheet_format_pr_xml).unwrap_or_default();
    let cols = layout.map(columns_xml).unwrap_or_default();
    let merges = layout.map(merges_xml).unwrap_or_default();
    (
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="{dimension}"/>{format_pr}{cols}<sheetData>{body}</sheetData>{merges}</worksheet>"#,
        ),
        cells,
    )
}

/// `<cols>` runs for the user-set columns of a walked sheet.
fn columns_xml(layout: &xls_layout::SheetLayout) -> String {
    if layout.cols.is_empty() {
        return String::new();
    }
    let mut xml = String::from("<cols>");
    for span in &layout.cols {
        xml.push_str(&format!(
            r#"<col min="{}" max="{}" width="{}" customWidth="1"{}/>"#,
            span.first + 1,
            span.last + 1,
            format_width(span.width),
            if span.hidden { r#" hidden="1""# } else { "" },
        ));
    }
    xml.push_str("</cols>");
    xml
}

/// Column width without a trailing decimal point (8, not 8.00).
fn format_width(width: f64) -> String {
    let rounded = (width * 100.0).round() / 100.0;
    if (rounded - rounded.trunc()).abs() < f64::EPSILON {
        format!("{}", rounded as i64)
    } else {
        format!("{rounded}")
    }
}

/// `<sheetFormatPr>` for a book whose default row height differs from the
/// converter's implicit 15pt (Excel 97-2003 books say 12.75): rows without
/// a height of their own keep the source's auto height, not ours.
fn sheet_format_pr_xml(layout: &xls_layout::SheetLayout) -> String {
    layout
        .default_row_height
        .map(|height| {
            format!(
                r#"<sheetFormatPr defaultRowHeight="{}"/>"#,
                format_width(height)
            )
        })
        .unwrap_or_default()
}

/// `ht`/`customHeight`/`hidden` attributes for a carried-over row. A kept
/// height is by definition manually set (the walk drops auto-fit values),
/// so it always brings customHeight.
fn row_attributes(spec: &xls_layout::RowSpec) -> String {
    let mut xml = String::new();
    if let Some(height) = spec.height {
        xml.push_str(&format!(
            r#" ht="{}" customHeight="1""#,
            format_width(height)
        ));
    }
    if spec.hidden {
        xml.push_str(r#" hidden="1""#);
    }
    xml
}

/// `<mergeCells>` for the walked ranges, sorted, deduplicated, clamped to
/// the sheet geometry.
fn merges_xml(layout: &xls_layout::SheetLayout) -> String {
    if layout.merges.is_empty() {
        return String::new();
    }
    let mut xml = format!(r#"<mergeCells count="{}">"#, layout.merges.len());
    for (rows, columns) in &layout.merges {
        xml.push_str(&format!(
            r#"<mergeCell ref="{}:{}"/>"#,
            cell_reference(u32::from(rows[0]), u32::from(columns[0])),
            cell_reference(u32::from(rows[1]), u32::from(columns[1])),
        ));
    }
    xml.push_str("</mergeCells>");
    xml
}

fn cell_xml(
    position: (u32, u32),
    value: &Data,
    formula: Option<&str>,
    style: Option<usize>,
) -> Option<String> {
    let reference = cell_reference(position.0, position.1);
    let style_attr = style
        .map(|style| format!(r#" s="{style}""#))
        .unwrap_or_default();
    let formula_xml = formula
        .map(|text| format!("<f>{}</f>", escape_xml(text)))
        .unwrap_or_default();
    let cell = match value {
        Data::Empty => {
            if formula.is_none() {
                return None;
            }
            format!(r#"<c r="{reference}"{style_attr}>{formula_xml}</c>"#)
        }
        Data::String(text) => format!(
            r#"<c r="{reference}"{style_attr} t="inlineStr">{formula_xml}<is><t xml:space="preserve">{}</t></is></c>"#,
            escape_xml(text),
        ),
        Data::Float(number) => {
            format!(r#"<c r="{reference}"{style_attr}>{formula_xml}<v>{number}</v></c>"#)
        }
        Data::Int(number) => {
            format!(r#"<c r="{reference}"{style_attr}>{formula_xml}<v>{number}</v></c>"#)
        }
        Data::Bool(flag) => format!(
            r#"<c r="{reference}"{style_attr} t="b">{formula_xml}<v>{}</v></c>"#,
            if *flag { 1 } else { 0 },
        ),
        Data::DateTime(datetime) => {
            let serial = datetime.as_f64();
            // Fallback styling for books without a walked layout: xf 1 =
            // short date, xf 2 = date+time. Walked books carry the real
            // number format in the cell's own style.
            let fallback = if serial.fract() == 0.0 { 1 } else { 2 };
            let style_attr = style
                .map(|style| format!(r#" s="{style}""#))
                .unwrap_or_else(|| format!(r#" s="{fallback}""#));
            format!(r#"<c r="{reference}"{style_attr}>{formula_xml}<v>{serial}</v></c>"#)
        }
        Data::Error(error) => format!(
            r#"<c r="{reference}"{style_attr} t="e">{formula_xml}<v>{}</v></c>"#,
            escape_xml(&error.to_string()),
        ),
        Data::DateTimeIso(text) | Data::DurationIso(text) => format!(
            r#"<c r="{reference}"{style_attr} t="inlineStr">{formula_xml}<is><t xml:space="preserve">{}</t></is></c>"#,
            escape_xml(text),
        ),
    };
    Some(cell)
}

fn cell_reference(row: u32, column: u32) -> String {
    let mut letters = String::new();
    let mut remaining = column + 1;
    while remaining > 0 {
        remaining -= 1;
        letters.insert(0, char::from(b'A' + (remaining % 26) as u8));
        remaining /= 26;
    }
    format!("{letters}{}", row + 1)
}

fn escape_xml(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn content_types_xml(sheet_count: usize) -> String {
    let overrides: String = (1..=sheet_count)
        .map(|index| format!(
            r#"<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
        ))
        .collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>{overrides}</Types>"#,
    )
}

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;

fn workbook_xml(names: &[String], repaired: &legacy_xls::LegacyXlsStrings) -> String {
    let sheets: String = names
        .iter()
        .enumerate()
        .map(|(index, name)| {
            // BUG-1600: BOUNDSHEET names mangle like the strings they sit
            // next to; the overlay carries the spec-correct decode.
            let display = repaired
                .sheet(index)
                .and_then(|sheet| sheet.name.as_deref())
                .unwrap_or(name);
            format!(
                r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
                escape_xml(display),
                index + 1,
                index + 1,
            )
        })
        .collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{sheets}</sheets></workbook>"#,
    )
}

fn workbook_rels_xml(sheet_count: usize) -> String {
    let mut relationships: String = (1..=sheet_count)
        .map(|index| format!(
            r#"<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>"#,
        ))
        .collect();
    relationships.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
        sheet_count + 1,
    ));
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}</Relationships>"#,
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ole2::test_fixtures;
    use std::io::Read;

    fn read_entry(path: &Path, name: &str) -> String {
        let mut archive = zip::ZipArchive::new(File::open(path).unwrap()).unwrap();
        let mut entry = archive.by_name(name).unwrap();
        let mut content = String::new();
        entry.read_to_string(&mut content).unwrap();
        content
    }

    #[test]
    fn converts_a_calamine_readable_workbook_into_minimal_xlsx() {
        // calamine reads xlsx through the same auto-detect path as xls, so a
        // generated xlsx fixture exercises the converter end to end.
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.xlsx");
        convert_fixture(&source);
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(&source, &target).unwrap();
        assert_eq!(result.sheets, 1);
        assert!(result.cells >= 4);

        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        assert!(sheet.contains(
            r#"<c r="A1" t="inlineStr"><is><t xml:space="preserve">Name &amp; Co</t></is></c>"#
        ));
        assert!(sheet.contains(r#"<c r="B1"><v>42</v></c>"#));
        assert!(sheet.contains(r#"<c r="B2"><f>B1*2</f>"#));
        let workbook = read_entry(&target, "xl/workbook.xml");
        assert!(workbook.contains(r#"<sheet name="Data" sheetId="1" r:id="rId1"/>"#));
    }

    /// Minimal .ods readable by calamine — the converter's real .ods shape.
    /// The mimetype must be the 46-byte ODF spreadsheet media type calamine
    /// compares byte for byte; `junk/pad.bin` mirrors the xlsx bomb fixture:
    /// an entry conversion never reads, so an inflated (but under-budget)
    /// declaration on it alone must not block conversion.
    fn write_ods_fixture(path: &Path, extra_entries: &[(&str, &str)]) {
        let file = File::create(path).unwrap();
        let mut writer = ZipWriter::new(file);
        let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
        let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        writer.start_file("mimetype", stored).unwrap();
        writer
            .write_all(b"application/vnd.oasis.opendocument.spreadsheet")
            .unwrap();
        let mut add = |name: &str, content: &str| {
            writer.start_file(name, deflated).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        };
        add(
            "META-INF/manifest.xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#,
        );
        add("content.xml", ODS_CONTENT_XML);
        for (name, content) in extra_entries {
            add(name, content);
        }
        writer.finish().unwrap();
    }

    const ODS_CONTENT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.2"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="string"><text:p>Alpha</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;

    #[test]
    fn converts_an_ods_spreadsheet_into_minimal_xlsx() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("book.ods");
        write_ods_fixture(&source, &[("junk/pad.bin", "x")]);
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(&source, &target).unwrap();
        assert_eq!(result.sheets, 1);
        assert_eq!(result.cells, 2);

        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        assert!(
            sheet.contains(
                r#"<c r="A1" t="inlineStr"><is><t xml:space="preserve">Alpha</t></is></c>"#
            )
        );
        assert!(sheet.contains(r#"<c r="B1"><v>2</v></c>"#));
        let workbook = read_entry(&target, "xl/workbook.xml");
        assert!(workbook.contains(r#"<sheet name="Sheet1" sheetId="1" r:id="rId1"/>"#));
    }

    /// SEC-1301: an .ods of a few hundred bytes whose central directory
    /// declares gigabytes must be refused by the convert path — the same
    /// SEC-1103 budget as the xlsx open, before calamine decompresses a
    /// single entry (mirror of the xlsx refuses_a_single_entry_declaring_
    /// gigabytes).
    #[test]
    fn refuses_an_ods_single_entry_declaring_gigabytes() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("bomb.ods");
        write_ods_fixture(&source, &[("junk/pad.bin", "x")]);
        assert_eq!(
            crate::archive::forge_declared_size(&source, "content.xml", 2 * 1024 * 1024 * 1024),
            1
        );
        let target = dir.path().join("converted.xlsx");

        let error = convert_to_xlsx(&source, &target).unwrap_err();
        let message = error.to_string();
        assert!(
            message.contains("uncompressed bytes") && message.contains("open budget"),
            "unexpected refusal: {message}"
        );
        assert!(!target.exists());
    }

    /// SEC-1301: the fence reads the SUM of the declared sizes — four modest
    /// 450 MiB entries total 1.8 GiB and must be refused just like the xlsx
    /// open path (mirror of refuses_the_combined_declared_total_across_
    /// modest_entries).
    #[test]
    fn refuses_an_ods_whose_modest_entries_total_over_the_budget() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("bomb.ods");
        // mimetype + manifest + content.xml + pad = 4 forged records
        write_ods_fixture(&source, &[("junk/pad.bin", "x")]);
        assert_eq!(
            crate::archive::forge_declared_size(&source, "", 450 * 1024 * 1024),
            4
        );
        let target = dir.path().join("converted.xlsx");

        let error = convert_to_xlsx(&source, &target).unwrap_err();
        assert!(error.to_string().contains("uncompressed"));
        assert!(!target.exists());
    }

    /// SEC-1301: a declared total under the 1.5 GiB budget still converts —
    /// only the never-read junk entry carries an inflated 1.4 GiB
    /// declaration (mirror of opens_when_the_declared_total_stays_under_
    /// the_budget). The fence refuses bombs, not big-looking metadata.
    #[test]
    fn converts_an_ods_when_the_declared_total_stays_under_the_budget() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("padded.ods");
        write_ods_fixture(&source, &[("junk/pad.bin", "x")]);
        assert_eq!(
            crate::archive::forge_declared_size(&source, "pad.bin", 1400 * 1024 * 1024),
            1
        );
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(&source, &target).unwrap();
        assert_eq!(result.sheets, 1);
    }

    /// SEC-1301: the parts-count budget covers .ods too — thousands of
    /// junk parts refuse conversion exactly like an oversized xlsx.
    #[test]
    fn refuses_an_ods_with_too_many_zip_entries() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("many.ods");
        let names: Vec<String> = (0..10_000)
            .map(|index| format!("junk/{index}.bin"))
            .collect();
        let junk: Vec<(&str, &str)> = names.iter().map(|name| (name.as_str(), "x")).collect();
        write_ods_fixture(&source, &junk);
        let target = dir.path().join("converted.xlsx");

        let error = convert_to_xlsx(&source, &target).unwrap_err();
        assert!(
            error
                .to_string()
                .contains("ZIP entries, above the 10000 entry open budget")
        );
        assert!(!target.exists());
    }

    /// SEC-1301: the central-directory structure check covers .ods too —
    /// an entry escaping the package root refuses conversion (mirror of the
    /// xlsx rejects_entries_escaping_the_package).
    #[test]
    fn refuses_an_ods_with_an_unsafe_entry_path() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("evil.ods");
        write_ods_fixture(&source, &[("../evil.xml", "<e/>")]);
        let target = dir.path().join("converted.xlsx");

        let error = convert_to_xlsx(&source, &target).unwrap_err();
        assert!(error.to_string().contains("unsafe ZIP path"));
        assert!(!target.exists());
    }

    /// SEC-1301: the fence is content-based and skips non-zip sources — a
    /// legacy .xls (OLE2 compound document, no central directory) fails
    /// inside calamine with its own message, exactly as before the fence.
    #[test]
    fn skips_the_fence_for_a_non_zip_source() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("book.xls");
        let mut bytes = vec![0xD0u8, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        bytes.extend_from_slice(&[0u8; 64]);
        std::fs::write(&source, bytes).unwrap();
        let target = dir.path().join("converted.xlsx");

        let error = convert_to_xlsx(&source, &target).unwrap_err();
        let message = error.to_string();
        assert!(
            message.contains("Unable to read the workbook"),
            "expected the calamine refusal, got: {message}"
        );
        assert!(!target.exists());
    }

    /// BUG-1600: a BIFF8 book with a single-byte codepage (1252) but
    /// UTF-16-flagged SST strings used to convert every Cyrillic string
    /// into mojibake plus U+0004 control characters, compressed strings
    /// into NUL-padded ones, and the sheet name into garbage. The committed
    /// fixture is a LibreOffice-generated workbook whose CODEPAGE record
    /// was patched 1200 -> 1252 (the exact report shape); xlrd reads all of
    /// its text cleanly, and so must the conversion.
    #[test]
    fn repairs_utf16_sst_strings_from_a_single_byte_codepage_book() {
        let source = std::path::Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/bug-1600-utf16-sst-codepage-1252.xls"
        ));
        let dir = tempfile::tempdir().unwrap();
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(source, &target).unwrap();
        assert_eq!(result.sheets, 1);

        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        assert!(sheet.contains(r#"xml:space="preserve">Наименование</t>"#));
        // Embedded newline of the source cell survives conversion.
        assert!(sheet.contains("Первая строка\nВторая строка"));
        // An 8-bit compressed SST string (fHighByte=0) must come out clean,
        // not NUL-padded the way calamine's codepage decode leaves it.
        assert!(sheet.contains("Plain ASCII item"));
        assert!(sheet.contains("Кириллица и ASCII mix"));
        assert!(!sheet.contains('\u{0}'), "NUL bytes in converted sheet");

        let workbook = read_entry(&target, "xl/workbook.xml");
        assert!(workbook.contains(r#"<sheet name="Прайс""#));
    }

    /// BUG-1602: a BIFF8 form — merged two-tier header, custom column
    /// widths, bold title on a solid fill, bordered boxes, date columns —
    /// must keep that shape through conversion. The committed fixture is a
    /// LibreOffice-generated workbook whose merges, widths, fonts, fills,
    /// borders and custom number formats were verified against xlrd and
    /// against LibreOffice's own xlsx conversion of the same file.
    #[test]
    fn carries_a_form_layout_from_a_legacy_book() {
        let source = std::path::Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/bug-1602-merges-widths-styles.xls"
        ));
        let dir = tempfile::tempdir().unwrap();
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(source, &target).unwrap();
        assert_eq!(result.sheets, 1);

        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        // All three merged ranges, sorted (title, horizontal, vertical).
        assert!(sheet.contains(r#"<mergeCells count="3">"#));
        assert!(sheet.contains(r#"<mergeCell ref="A1:D1"/>"#));
        assert!(sheet.contains(r#"<mergeCell ref="B3:D3"/>"#));
        assert!(sheet.contains(r#"<mergeCell ref="A3:A4"/>"#));
        // The three custom widths; the producer's catch-all record for the
        // untouched tail (8.68 chars) must not become 256 custom columns.
        assert!(sheet.contains(r#"<col min="1" max="1" width="8" customWidth="1"/>"#));
        assert!(sheet.contains(r#"<col min="2" max="2" width="32" customWidth="1"/>"#));
        assert!(sheet.contains(r#"<col min="3" max="3" width="15" customWidth="1"/>"#));
        assert!(!sheet.contains("8.68"), "catch-all width leaked into cols");
        // Merge continuations keep the header's fill and borders via styled
        // empty cells, the way LibreOffice writes them.
        assert!(sheet.contains(r#"<c r="B1" s="3"/>"#));
        assert!(sheet.contains(r#"<c r="C3" s="4"/>"#));
        // Title cell: bold 14pt on the amber fill, centered — cellXfs 3.
        assert!(sheet.contains(r#"<c r="A1" s="3" t="inlineStr">"#));
        // A date keeps its real number format (custom 165 = yyyy-mm-dd),
        // not the fallback short-date style.
        assert!(sheet.contains(r#"<c r="D5" s="6"><v>46223</v></c>"#));

        let styles = read_entry(&target, "xl/styles.xml");
        assert!(styles.contains(r#"<font><b/><sz val="14"/><name val="Cambria"/></font>"#));
        assert!(styles.contains(r#"<fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill>"#));
        assert!(styles.contains(r#"<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill>"#));
        assert!(styles.contains(r#"<left style="thin"><color rgb="FF000000"/></left>"#));
        assert!(styles.contains(r#"<numFmt numFmtId="165" formatCode="yyyy\-mm\-dd"/>"#));
        assert!(styles.contains(r#"<alignment horizontal="center" vertical="center"/>"#));
        // cellXfs 0-2 are the converter's fallback styles, unchanged.
        assert!(styles.contains(r#"<cellXfs count="7">"#));
        // BUG-1659: the mandatory named style, after cellXfs in schema
        // order — IronCalc's importer panics on a book without it.
        assert!(styles.contains(
            r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#
        ));
        assert!(styles.find("</cellXfs>").unwrap() < styles.find("<cellStyles").unwrap());
    }

    /// BUG-1607: a book with custom row heights and hidden rows must keep
    /// both through conversion — forms with non-standard header heights
    /// used to arrive auto-height, and hidden rows reappeared. The
    /// committed fixture is a LibreOffice-generated workbook whose row
    /// table was verified against xlrd (heights, hidden flags) and against
    /// LibreOffice's own xlsx conversion of the same file.
    #[test]
    fn carries_row_heights_and_hidden_rows_from_a_legacy_book() {
        let source = std::path::Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/bug-1607-row-heights-hidden.xls"
        ));
        let dir = tempfile::tempdir().unwrap();
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(source, &target).unwrap();
        assert_eq!(result.sheets, 1);
        assert_eq!(result.cells, 4);

        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        // Manually-set heights become ht + customHeight.
        assert!(sheet.contains(r#"<row r="1" ht="30" customHeight="1">"#));
        assert!(sheet.contains(r#"<row r="4" ht="7.5" customHeight="1">"#));
        // A hidden row keeps the flag and gains no customHeight (its height
        // is the producer's default); default rows stay untouched.
        assert!(sheet.contains(r#"<row r="2" hidden="1">"#));
        assert!(sheet.contains(r#"<row r="3">"#));
        // Heights and hidden flags on rows without any cells still emit
        // their <row> stub — a sized spacer or a filtered empty row.
        assert!(sheet.contains(r#"<row r="5" ht="45" customHeight="1"/>"#));
        assert!(sheet.contains(r#"<row r="6" hidden="1"/>"#));
        // The book's default row height (15pt) is our implicit default:
        // nothing to carry into sheetFormatPr.
        assert!(!sheet.contains("sheetFormatPr"));
    }

    /// BUG-1607: a book whose default row height differs from the
    /// converter's implicit 15pt carries it as `<sheetFormatPr>` — placed
    /// in schema order, before cols/sheetData — so its auto-fit rows keep
    /// the source's height instead of inheriting ours.
    #[test]
    fn a_foreign_default_row_height_is_carried_into_sheet_format_pr() {
        let layout = xls_layout::WorkbookLayout::for_test(vec![xls_layout::SheetLayout {
            rows: vec![
                (
                    0,
                    xls_layout::RowSpec {
                        height: Some(30.0),
                        hidden: false,
                    },
                ),
                (
                    3,
                    xls_layout::RowSpec {
                        height: None,
                        hidden: true,
                    },
                ),
            ],
            default_row_height: Some(12.75),
            ..Default::default()
        }]);
        let range =
            calamine::Range::<Data>::from_sparse(vec![calamine::Cell::new((0, 0), Data::Float(1.0))]);
        let mut styler = StyleInterner::new(&layout);
        let (xml, cells) = worksheet_xml(&range, &HashMap::new(), None, layout.sheet(0), &mut styler);
        assert_eq!(cells, 1);
        assert!(xml.contains(r#"<sheetFormatPr defaultRowHeight="12.75"/>"#));
        // A carried height lands on rows with cells too, attributes first.
        assert!(xml.contains(r#"<row r="1" ht="30" customHeight="1"><c r="A1"><v>1</v></c></row>"#));
        assert!(xml.contains(r#"<row r="4" hidden="1"/>"#));
        // Schema order: sheetFormatPr between dimension and sheetData.
        let format_pr = xml.find("sheetFormatPr").unwrap();
        let sheet_data = xml.find("<sheetData>").unwrap();
        let dimension = xml.find("<dimension").unwrap();
        assert!(dimension < format_pr && format_pr < sheet_data);
    }

    /// BUG-1602: zip-based sources have no BIFF walk; their conversion must
    /// stay byte-for-byte the minimal output — no merges, no cols, no extra
    /// styles.
    #[test]
    fn keeps_zip_sources_free_of_legacy_layout() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.xlsx");
        convert_fixture(&source);
        let target = dir.path().join("converted.xlsx");

        convert_to_xlsx(&source, &target).unwrap();
        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        assert!(!sheet.contains("mergeCell"));
        assert!(!sheet.contains("<cols>"));
        let styles = read_entry(&target, "xl/styles.xml");
        assert!(styles.contains(r#"<cellXfs count="3">"#));
    }

    /// BUG-1659: a real .ods book converts with a complete styles.xml —
    /// the BIFF walk finds nothing in a zip container, so this exercises
    /// the BASE_STYLES_XML path the .ods import always takes.
    #[test]
    fn carries_cell_styles_from_an_ods_book() {
        let source = std::path::Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/bug-1659-legacy-formulas.ods"
        ));
        let dir = tempfile::tempdir().unwrap();
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(source, &target).unwrap();
        assert_eq!(result.sheets, 1);

        let styles = read_entry(&target, "xl/styles.xml");
        assert!(styles.contains(
            r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#
        ));
        assert!(styles.ends_with("</cellStyles></styleSheet>"));
        // Values and formulas came through the conversion (the verbatim
        // ODF `of:=` syntax is a separate known gap).
        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        assert!(sheet.contains(r#"<c r="A1"><v>10</v></c>"#));
        assert!(sheet.contains("<f>of:=SUM([.A1:.A2])</f>"));
    }

    /// BUG-1659: an empty book (no cells, no styles) still converts with a
    /// complete styles.xml — the minimal output must not skip the section
    /// IronCalc's importer requires.
    #[test]
    fn an_empty_book_still_carries_cell_styles() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("empty.xlsx");
        convert_empty_book_fixture(&source);
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(&source, &target).unwrap();
        assert_eq!(result.cells, 0);
        let styles = read_entry(&target, "xl/styles.xml");
        assert!(styles.contains(r#"<cellXfs count="3">"#));
        assert!(styles.contains(
            r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#
        ));
        assert!(styles.ends_with("</cellStyles></styleSheet>"));
    }

    /// BUG-1602: the layout walk must never panic, on truncated books just
    /// like the string overlay it mirrors (same truncation sweep). The walk
    /// runs on its own; the full convert path cannot be swept this way
    /// because calamine itself panics on some truncated OLE2 headers —
    /// pre-existing upstream behavior, unchanged by this module.
    #[test]
    fn truncated_layout_fixture_never_panics_the_walk() {
        let books = [
            "bug-1602-merges-widths-styles.xls",
            // BUG-1607: exercises the ROW/DEFAULTROWHEIGHT record paths.
            "bug-1607-row-heights-hidden.xls",
        ];
        let manifest = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
        let dir = tempfile::tempdir().unwrap();
        for name in books {
            let full = std::fs::read(manifest.join(name)).unwrap();
            for cut in (0..full.len()).step_by(64) {
                let path = dir.path().join("cut.xls");
                std::fs::write(&path, &full[..cut]).unwrap();
                let _ = xls_layout::WorkbookLayout::extract(&path);
            }
        }
    }

    /// BUG-1606: the exact truncated-header shape behind the calamine panic
    /// (slice index failure at cfb.rs:306): a well-formed magic and sector
    /// shift, but the only FAT sector named by the header DIFAT lies beyond
    /// the end of the file, so calamine slices past the bytes it read while
    /// loading the FAT. Must refuse with a normal conversion error, not
    /// crash the sidecar.
    #[test]
    fn refuses_a_header_whose_fat_sector_lies_beyond_eof() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("cut-fat.xls");
        let mut bytes = test_fixtures::zeros(2);
        bytes[76..80].copy_from_slice(&9u32.to_le_bytes()); // DIFAT[0] = 9
        std::fs::write(&source, bytes).unwrap();
        let target = dir.path().join("converted.xlsx");

        let error = convert_to_xlsx(&source, &target).unwrap_err();
        let message = error.to_string();
        assert!(
            message.contains("truncated or corrupt"),
            "expected the OLE2 fence refusal, got: {message}"
        );
        assert!(!target.exists());
    }

    /// BUG-1606: a header the fence accepts must still fail gracefully when
    /// calamine finds nothing readable inside — this structurally valid
    /// compound file has no Workbook stream, and that is an error, not a
    /// crash. (The fixture is content-valid on purpose: an all-zero FAT
    /// would trip the fence's empty-FAT refusal instead of exercising the
    /// calamine path, and makes calamine 0.36 allocate without bound.)
    #[test]
    fn reports_a_streamless_valid_compound_file_as_an_error() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("empty.xls");
        std::fs::write(&source, test_fixtures::valid_cfb(0)).unwrap();
        let target = dir.path().join("converted.xlsx");

        let error = convert_to_xlsx(&source, &target).unwrap_err();
        assert!(
            error.to_string().contains("Unable to read the workbook"),
            "{error}"
        );
        assert!(!target.exists());
    }

    /// BUG-1606: the smallest compound file calamine can actually convert —
    /// built programmatically (header, FAT, directory, one 4096-byte
    /// Workbook stream with a BIFF8 globals+sheet stub) — passes the fence
    /// and converts end to end.
    #[test]
    fn converts_a_minimal_programmatic_legacy_book() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("minimal.xls");
        std::fs::write(&source, test_fixtures::minimal_biff8_cfb()).unwrap();
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(&source, &target).unwrap();
        assert_eq!(result.sheets, 1);
        let workbook = read_entry(&target, "xl/workbook.xml");
        assert!(workbook.contains(r#"<sheet name="S" sheetId="1" r:id="rId1"/>"#));
    }

    /// BUG-1606: cutting a real legacy book at every 64-byte boundary —
    /// including every header truncation that used to panic inside
    /// calamine's CFB reader — must end in a converted file or a normal
    /// conversion error, never in a panic that kills the sidecar. Cuts are
    /// generated in the test; no truncated fixture is committed.
    #[test]
    fn truncated_legacy_books_never_panic_the_convert_path() {
        let full = std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/bug-1602-merges-widths-styles.xls"
        ))
        .unwrap();
        let dir = tempfile::tempdir().unwrap();
        for cut in (0..full.len()).step_by(64) {
            let path = dir.path().join("cut.xls");
            std::fs::write(&path, &full[..cut]).unwrap();
            let target = dir.path().join("converted.xlsx");
            let _ = std::fs::remove_file(&target);
            match convert_to_xlsx(&path, &target) {
                Ok(_) => assert!(target.exists()),
                Err(_) => assert!(!target.exists()),
            }
        }
    }

    /// Same minimal book as `convert_fixture`, but with an empty sheet —
    /// the source shape behind the empty-book regression test.
    fn convert_empty_book_fixture(path: &Path) {
        let file = File::create(path).unwrap();
        let mut writer = ZipWriter::new(file);
        let options = SimpleFileOptions::default();
        let mut add = |name: &str, content: &str| {
            writer.start_file(name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        };
        add(
            "[Content_Types].xml",
            r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
        );
        add(
            "_rels/.rels",
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        );
        add(
            "xl/workbook.xml",
            r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        );
        add(
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        );
        add(
            "xl/worksheets/sheet1.xml",
            r#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#,
        );
        writer.finish().unwrap();
    }

    fn convert_fixture(path: &Path) {
        let file = File::create(path).unwrap();
        let mut writer = ZipWriter::new(file);
        let options = SimpleFileOptions::default();
        let mut add = |name: &str, content: &str| {
            writer.start_file(name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        };
        add(
            "[Content_Types].xml",
            r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
        );
        add(
            "_rels/.rels",
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        );
        add(
            "xl/workbook.xml",
            r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        );
        add(
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        );
        add(
            "xl/worksheets/sheet1.xml",
            r#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name &amp; Co</t></is></c><c r="B1"><v>42</v></c></row><row r="2"><c r="A2" t="b"><v>1</v></c><c r="B2"><f>B1*2</f><v>84</v></c></row></sheetData></worksheet>"#,
        );
        writer.finish().unwrap();
    }
}

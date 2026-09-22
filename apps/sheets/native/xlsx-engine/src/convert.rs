//! Legacy-format import: reads an .xls (or anything calamine understands)
//! and writes a minimal .xlsx with values, formulas, and date number formats.
//! Styles beyond date formats are not carried over — the converted file is a
//! fresh workbook, not a byte-preserving edit.

use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;

use calamine::{open_workbook_auto, Data, Reader};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::SidecarError;

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
    let mut workbook = open_workbook_auto(source)
        .map_err(|error| SidecarError::Workbook(format!("Unable to read the workbook: {error}")))?;
    let names: Vec<String> = workbook.sheet_names().to_vec();
    if names.is_empty() {
        return Err(SidecarError::Workbook("The workbook has no sheets.".into()));
    }

    let mut sheet_xmls: Vec<String> = Vec::new();
    let mut cells = 0usize;
    for name in &names {
        let range = workbook
            .worksheet_range(name)
            .map_err(|error| SidecarError::Workbook(format!("Sheet {name}: {error}")))?;
        let formulas = workbook.worksheet_formula(name).ok();
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
        let (xml, sheet_cells) = worksheet_xml(&range, &formula_map);
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
    add("xl/workbook.xml", &workbook_xml(&names))?;
    add("xl/_rels/workbook.xml.rels", &workbook_rels_xml(names.len()))?;
    add("xl/styles.xml", STYLES_XML)?;
    for (index, xml) in sheet_xmls.iter().enumerate() {
        add(&format!("xl/worksheets/sheet{}.xml", index + 1), xml)?;
    }
    writer.finish()?.sync_all()?;
    Ok(ConvertResult { sheets: names.len(), cells, source_bytes })
}

fn worksheet_xml(
    range: &calamine::Range<Data>,
    formulas: &HashMap<(u32, u32), String>,
) -> (String, usize) {
    // Rows carrying values plus rows carrying only formulas.
    let mut rows: HashMap<u32, Vec<(u32, String)>> = HashMap::new();
    let mut cells = 0usize;
    let (start_row, start_col) = range.start().unwrap_or((0, 0));
    for (row, column, value) in range.used_cells() {
        let absolute = (start_row + row as u32, start_col + column as u32);
        let formula = formulas.get(&absolute).map(String::as_str);
        if let Some(cell) = cell_xml(absolute, value, formula) {
            cells += 1;
            rows.entry(absolute.0).or_default().push((absolute.1, cell));
        }
    }
    for (position, formula) in formulas {
        let covered = range.get_value((position.0, position.1)).is_some_and(|v| *v != Data::Empty);
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

    let mut row_numbers: Vec<u32> = rows.keys().copied().collect();
    row_numbers.sort_unstable();
    let mut body = String::new();
    for row in row_numbers {
        let mut line = rows.remove(&row).unwrap_or_default();
        line.sort_unstable_by_key(|(column, _)| *column);
        body.push_str(&format!(r#"<row r="{}">"#, row + 1));
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
    (
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="{dimension}"/><sheetData>{body}</sheetData></worksheet>"#,
        ),
        cells,
    )
}

fn cell_xml(position: (u32, u32), value: &Data, formula: Option<&str>) -> Option<String> {
    let reference = cell_reference(position.0, position.1);
    let formula_xml = formula
        .map(|text| format!("<f>{}</f>", escape_xml(text)))
        .unwrap_or_default();
    let cell = match value {
        Data::Empty => {
            if formula.is_none() {
                return None;
            }
            format!(r#"<c r="{reference}">{formula_xml}</c>"#)
        }
        Data::String(text) => format!(
            r#"<c r="{reference}" t="inlineStr">{formula_xml}<is><t xml:space="preserve">{}</t></is></c>"#,
            escape_xml(text),
        ),
        Data::Float(number) => format!(r#"<c r="{reference}">{formula_xml}<v>{number}</v></c>"#),
        Data::Int(number) => format!(r#"<c r="{reference}">{formula_xml}<v>{number}</v></c>"#),
        Data::Bool(flag) => format!(
            r#"<c r="{reference}" t="b">{formula_xml}<v>{}</v></c>"#,
            if *flag { 1 } else { 0 },
        ),
        Data::DateTime(datetime) => {
            let serial = datetime.as_f64();
            let style = if serial.fract() == 0.0 { 1 } else { 2 };
            format!(r#"<c r="{reference}" s="{style}">{formula_xml}<v>{serial}</v></c>"#)
        }
        Data::Error(error) => format!(
            r#"<c r="{reference}" t="e">{formula_xml}<v>{}</v></c>"#,
            escape_xml(&error.to_string()),
        ),
        Data::DateTimeIso(text) | Data::DurationIso(text) => format!(
            r#"<c r="{reference}" t="inlineStr">{formula_xml}<is><t xml:space="preserve">{}</t></is></c>"#,
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

fn workbook_xml(names: &[String]) -> String {
    let sheets: String = names
        .iter()
        .enumerate()
        .map(|(index, name)| format!(
            r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
            escape_xml(name),
            index + 1,
            index + 1,
        ))
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

/// xf 1 = short date (numFmt 14), xf 2 = date+time (numFmt 22).
const STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#;

#[cfg(test)]
mod tests {
    use super::*;
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
        assert!(sheet.contains(r#"<c r="A1" t="inlineStr"><is><t xml:space="preserve">Name &amp; Co</t></is></c>"#));
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
        assert!(sheet.contains(r#"<c r="A1" t="inlineStr"><is><t xml:space="preserve">Alpha</t></is></c>"#));
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
        let names: Vec<String> = (0..10_000).map(|index| format!("junk/{index}.bin")).collect();
        let junk: Vec<(&str, &str)> = names.iter().map(|name| (name.as_str(), "x")).collect();
        write_ods_fixture(&source, &junk);
        let target = dir.path().join("converted.xlsx");

        let error = convert_to_xlsx(&source, &target).unwrap_err();
        assert!(error
            .to_string()
            .contains("ZIP entries, above the 10000 entry open budget"));
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

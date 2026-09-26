//! Slicer parts (`xl/slicers/slicerN.xml` + `xl/slicerCaches/slicerCacheN.xml`).
//! Only TABLE slicers are surfaced: an x15 `tableSlicerCache` extension binds
//! the slicer to a table's column, which the host can restore as a filter
//! panel over the table's auto-filter range. Pivot slicers (a `<pivotTables>`
//! binding instead) stay host-side session state and are skipped here.

use super::*;

/// One table slicer bound to a sheet's table column.
#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerInfo {
    /// slicer/@name — the workbook-unique slicer name.
    pub name: String,
    /// slicer/@cache — the slicerCacheDefinition this slicer binds to.
    pub cache_name: String,
    /// slicer/@caption when present (Excel shows it as the panel title).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub caption: Option<String>,
    /// The bound table's name (displayName falling back to name).
    pub table_name: String,
    /// x15:tableSlicerCache/@column — the 1-based position of the column in
    /// the table's tableColumns list.
    pub column: usize,
}

/// The workbook's slicerCache parts keyed by their slicerCacheDefinition
////@name (the token slicer/@cache references).
struct SlicerCacheBinding {
    /// (tableId, column) of the x15 tableSlicerCache extension, when the
    /// cache is a table slicer.
    table_slicer: Option<(usize, usize)>,
}

/// Reads every table slicer reachable from a worksheet's slicerList extension.
/// Malformed or unbindable entries (cache missing, tableId unknown) are
/// skipped — a slicer the host cannot bind must not fail the whole workbook.
pub(crate) fn read_sheet_slicers(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
) -> Result<Vec<SlicerInfo>, SidecarError> {
    // slicer/@cache → binding, gathered once for the whole workbook.
    let caches = read_slicer_caches(archive)?;
    if caches.is_empty() {
        return Ok(Vec::new());
    }
    let table_ids = read_table_ids(archive, worksheet_path)?;
    if table_ids.is_empty() {
        return Ok(Vec::new());
    }
    let mut slicers = Vec::new();
    for slicer_path in worksheet_slicer_paths(archive, worksheet_path)? {
        let xml = read_zip_string(archive, &slicer_path)?;
        for (name, cache_name, caption) in parse_slicer_entries(&xml) {
            let Some(binding) = caches.get(&cache_name) else {
                continue;
            };
            // Pivot slicer caches carry a pivotTables binding instead.
            let Some((table_id, column)) = binding.table_slicer else {
                continue;
            };
            let Some(table_name) = table_ids.get(&table_id) else {
                continue;
            };
            slicers.push(SlicerInfo {
                name,
                cache_name,
                caption,
                table_name: table_name.clone(),
                column,
            });
        }
    }
    Ok(slicers)
}

/// Package paths of the slicer parts referenced from a worksheet's
/// relationships (the x14/x15 slicerList ext hangs off the same rels).
fn worksheet_slicer_paths(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
) -> Result<Vec<String>, SidecarError> {
    Ok(visuals::read_relationships(archive, worksheet_path)?
        .into_values()
        .filter(|relationship| relationship.relationship_type.ends_with("/slicer"))
        .filter_map(|relationship| {
            visuals::resolve_part_target(worksheet_path, &relationship.target).ok()
        })
        .collect())
}

/// Parses `<slicer name cache caption/>` entries; the caption is optional.
fn parse_slicer_entries(xml: &str) -> Vec<(String, String, Option<String>)> {
    let mut reader = Reader::from_str(xml);
    let mut entries = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Empty(element)) | Ok(Event::Start(element))
                if element.local_name().as_ref() == b"slicer" =>
            {
                let read = |key: &[u8]| {
                    crate::xml_util::attribute_value(&reader, &element, key)
                        .ok()
                        .flatten()
                };
                if let (Some(name), Some(cache)) = (read(b"name"), read(b"cache")) {
                    entries.push((name, cache, read(b"caption")));
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    entries
}

/// Every slicerCache part of the workbook, keyed by definition name. The
/// cache paths come from the workbook rels; a cache without a table binding
/// (pivot slicers) is kept but flagged through `table_slicer: None`.
fn read_slicer_caches(
    archive: &mut ZipArchive<File>,
) -> Result<HashMap<String, SlicerCacheBinding>, SidecarError> {
    let mut caches = HashMap::new();
    let workbook_rels = visuals::read_relationships(archive, "xl/workbook.xml")?;
    for relationship in workbook_rels.values() {
        if !relationship.relationship_type.ends_with("/slicerCache") {
            continue;
        }
        let Ok(path) = visuals::resolve_part_target("xl/workbook.xml", &relationship.target) else {
            continue;
        };
        let Ok(xml) = read_zip_string(archive, &path) else {
            continue;
        };
        if let Some((name, binding)) = parse_slicer_cache(&xml) {
            caches.insert(name, binding);
        }
    }
    Ok(caches)
}

/// Parses one slicerCacheDefinition: its name, sourceName, and the x15
/// tableSlicerCache extension when present.
fn parse_slicer_cache(xml: &str) -> Option<(String, SlicerCacheBinding)> {
    let mut reader = Reader::from_str(xml);
    let mut name = None;
    let mut in_ext = false;
    let mut table_slicer = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element))
                if element.local_name().as_ref() == b"slicerCacheDefinition" =>
            {
                // The x15 twin <x15:slicerCacheDefinition> (inside the extLst)
                // shares the local name and carries no attributes — only the
                // outer definition's name may stick.
                if name.is_none() {
                    name = crate::xml_util::attribute_value(&reader, &element, b"name")
                        .ok()
                        .flatten();
                }
            }
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"extLst" => {
                in_ext = true;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"extLst" => {
                in_ext = false;
            }
            Ok(Event::Empty(element)) | Ok(Event::Start(element))
                if in_ext && element.local_name().as_ref() == b"tableSlicerCache" =>
            {
                let read = |key: &[u8]| {
                    crate::xml_util::attribute_value(&reader, &element, key)
                        .ok()
                        .flatten()
                };
                table_slicer = match (read(b"tableId"), read(b"column")) {
                    (Some(table_id), Some(column)) => Some((
                        table_id.parse::<usize>().ok()?,
                        column.parse::<usize>().ok()?,
                    )),
                    _ => None,
                };
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    Some((name?, SlicerCacheBinding { table_slicer }))
}

/// table/@id → table name for the worksheet's table parts: the token
/// x15:tableSlicerCache/@tableId references.
fn read_table_ids(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
) -> Result<HashMap<usize, String>, SidecarError> {
    let mut ids = HashMap::new();
    for table_path in visuals::table_part_paths(archive, worksheet_path)? {
        let Ok(xml) = read_zip_string(archive, &table_path) else {
            continue;
        };
        let mut reader = Reader::from_str(&xml);
        loop {
            match reader.read_event() {
                Ok(Event::Empty(element)) | Ok(Event::Start(element))
                    if element.local_name().as_ref() == b"table" =>
                {
                    let read = |key: &[u8]| {
                        crate::xml_util::attribute_value(&reader, &element, key)
                            .ok()
                            .flatten()
                    };
                    if let Some(id) = read(b"id").and_then(|value| value.parse::<usize>().ok()) {
                        let name = read(b"displayName").or_else(|| read(b"name"));
                        if let Some(name) = name {
                            ids.insert(id, name);
                        }
                    }
                    break;
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
    }
    Ok(ids)
}

/// Live per-column criteria of a worksheet's FIRST table autoFilter, for
/// sheets whose filter belongs to a table rather than the worksheet (Excel
/// stores those criteria in the table part, not the sheet). Only the
/// value/blank and custom shapes are mapped; color criteria (dxf-bound) and
/// icon/dynamic/top10 filters yield no criteria here.
pub(crate) fn read_table_filter_columns(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
) -> Result<Vec<FilterColumnCriteria>, SidecarError> {
    let Some(table_path) = visuals::table_part_paths(archive, worksheet_path)?
        .first()
        .cloned()
    else {
        return Ok(Vec::new());
    };
    let xml = read_zip_string(archive, &table_path)?;
    let mut reader = Reader::from_str(&xml);
    let mut in_auto_filter = false;
    let mut column: Option<FilterColumnCriteria> = None;
    let mut columns: Vec<FilterColumnCriteria> = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if element.local_name().as_ref() == b"autoFilter" =>
            {
                in_auto_filter = true;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"autoFilter" => {
                in_auto_filter = false;
            }
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if in_auto_filter && element.local_name().as_ref() == b"filterColumn" =>
            {
                column = crate::xml_util::attribute_value(&reader, &element, b"colId")?
                    .and_then(|value| value.parse::<usize>().ok())
                    .map(|col_id| FilterColumnCriteria {
                        col_id,
                        values: None,
                        blank: false,
                        customs: None,
                        color_filter: None,
                    });
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"filterColumn" => {
                if let Some(column) = column.take() {
                    if column.values.is_some() || column.blank || column.customs.is_some() {
                        columns.push(column);
                    }
                }
            }
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if in_auto_filter && element.local_name().as_ref() == b"filters" =>
            {
                if let Some(column) = column.as_mut() {
                    column.blank = crate::xml_util::attribute_value(&reader, &element, b"blank")?
                        .is_some_and(|value| value == "1" || value == "true");
                    column.values.get_or_insert_with(Vec::new);
                }
            }
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if in_auto_filter && element.local_name().as_ref() == b"filter" =>
            {
                if let Some(values) = column.as_mut().and_then(|column| column.values.as_mut()) {
                    if let Some(value) =
                        crate::xml_util::attribute_value(&reader, &element, b"val")?
                    {
                        values.push(value);
                    }
                }
            }
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if in_auto_filter && element.local_name().as_ref() == b"customFilters" =>
            {
                if let Some(column) = column.as_mut() {
                    column.customs = Some(CustomFilterCriteria {
                        and: crate::xml_util::attribute_value(&reader, &element, b"and")?
                            .is_some_and(|value| value == "1" || value == "true"),
                        filters: Vec::new(),
                    });
                }
            }
            Ok(Event::Empty(element))
                if in_auto_filter && element.local_name().as_ref() == b"customFilter" =>
            {
                let read = |key: &[u8]| {
                    crate::xml_util::attribute_value(&reader, &element, key)
                        .ok()
                        .flatten()
                };
                if let (Some(value), Some(customs)) = (
                    read(b"val"),
                    column.as_mut().and_then(|column| column.customs.as_mut()),
                ) {
                    customs.filters.push(CustomFilterItem {
                        val: value,
                        operator: read(b"operator"),
                    });
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    Ok(columns)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write as IoWrite;

    const WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;

    const WORKBOOK_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="/xl/slicerCaches/slicerCache1.xml"/><Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="/xl/slicerCaches/slicerCache2.xml"/></Relationships>"#;

    const TABLE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="7" name="Sales" displayName="Sales" ref="A1:B6" headerRowCount="1"><autoFilter ref="A1:B6"><filterColumn colId="0"><filters><filter val="alpha"/><filter val="beta"/></filters></filterColumn></autoFilter><tableColumns count="2"><tableColumn id="1" name="Product"/><tableColumn id="2" name="Amount"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>"#;

    const SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Product</t></is></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>"#;

    const SHEET_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/></Relationships>"#;

    const SLICER_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><slicer name="Product" cache="Slicer_Product" caption="Product" rowHeight="241300"/></slicers>"#;

    const SLICER_CACHE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Slicer_Product" sourceName="Product"><extLst><ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:slicerCacheDefinition><x15:tableSlicerCache tableId="7" column="1"/></x15:slicerCacheDefinition></ext></extLst></slicerCacheDefinition>"#;

    /// A second cache that is NOT a table slicer (pivotTables binding) —
    /// its slicer must be skipped.
    const PIVOT_CACHE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_Region" sourceName="Region"><pivotTables><pivotTable tabId="1" name="Pivot1"/></pivotTables></slicerCacheDefinition>"#;

    fn write_slicer_workbook(path: &Path) {
        let dir = path.parent().expect("fixture parent");
        std::fs::create_dir_all(dir.join("xl/worksheets/_rels")).unwrap();
        std::fs::create_dir_all(dir.join("xl/tables")).unwrap();
        std::fs::create_dir_all(dir.join("xl/slicers")).unwrap();
        std::fs::create_dir_all(dir.join("xl/slicerCaches")).unwrap();
        let mut writer = zip::ZipWriter::new(File::create(path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, content) in [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            ),
            ("xl/workbook.xml", WORKBOOK_XML),
            ("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
            ("xl/worksheets/sheet1.xml", SHEET_XML),
            ("xl/worksheets/_rels/sheet1.xml.rels", SHEET_RELS),
            ("xl/tables/table1.xml", TABLE_XML),
            ("xl/slicers/slicer1.xml", SLICER_XML),
            ("xl/slicerCaches/slicerCache1.xml", SLICER_CACHE_XML),
            ("xl/slicerCaches/slicerCache2.xml", PIVOT_CACHE_XML),
        ] {
            writer.start_file(name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
    }

    #[test]
    fn reads_table_slicers_and_skips_pivot_caches() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("slicers.xlsx");
        write_slicer_workbook(&path);
        let file = File::open(&path).unwrap();
        let mut archive = ZipArchive::new(file).unwrap();
        let slicers = read_sheet_slicers(&mut archive, "xl/worksheets/sheet1.xml").unwrap();
        assert_eq!(slicers.len(), 1);
        assert_eq!(slicers[0].name, "Product");
        assert_eq!(slicers[0].cache_name, "Slicer_Product");
        assert_eq!(slicers[0].caption.as_deref(), Some("Product"));
        assert_eq!(slicers[0].table_name, "Sales");
        assert_eq!(slicers[0].column, 1);
    }

    #[test]
    fn reads_table_filter_criteria_when_the_sheet_has_no_auto_filter() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("slicers.xlsx");
        write_slicer_workbook(&path);
        let file = File::open(&path).unwrap();
        let mut archive = ZipArchive::new(file).unwrap();
        let columns = read_table_filter_columns(&mut archive, "xl/worksheets/sheet1.xml").unwrap();
        assert_eq!(columns.len(), 1);
        assert_eq!(columns[0].col_id, 0);
        assert_eq!(
            columns[0].values.as_deref(),
            Some(["alpha".to_string(), "beta".to_string()].as_slice())
        );
    }
}

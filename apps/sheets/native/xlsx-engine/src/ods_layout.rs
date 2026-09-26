//! .ods layout carry-over for the convert path (PAR-214).
//!
//! calamine reads the cell values of an .ods book but drops its shape:
//! merged ranges live in `table:number-columns-spanned` /
//! `table:number-rows-spanned` anchors and their `table:covered-table-cell`
//! continuations, and none of that survives the value walk — a form's
//! merged header converts into a grid of loose cells while LibreOffice
//! keeps every merge. This module reads the source package directly
//! (content.xml, the same part calamine consumes) and rebuilds what the
//! converter can carry, reusing the `xls_layout` emission types so the
//! convert path treats both walks alike.
//!
//! Everything is best-effort, exactly like the BIFF walk: only an ODF
//! spreadsheet package (mimetype entry, content-based) is walked, any
//! structural surprise yields an empty layout, and every count is
//! bounds-checked so hostile input cannot grow the walk unboundedly.

use std::fs::File;
use std::io::Read;
use std::path::Path;

use quick_xml::Reader;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use zip::ZipArchive;

use crate::xls_layout::{SheetLayout, WorkbookLayout};

/// Caps: a legitimate form carries tens of merges; the caps only bind for
/// hostile or pathological input, where the walk degrades to a partial
/// layout instead of buffering unbounded state.
const MAX_MERGES_PER_SHEET: usize = 65_536;
/// ODF/Calc sheets carry at most this many rows/columns (Excel's own grid
/// limits); repeat counts are clamped so a hostile
/// `table:number-*-repeated` cannot spin the walk or the coordinates.
const MAX_ROWS: u32 = 1_048_576;
const MAX_COLUMNS: u32 = 16_384;
const MAX_ROW_REPEAT: u32 = 1_048_576;
const MAX_COLUMN_REPEAT: u32 = 16_384;
/// ODF spans are counts (>= 1); anything past the grid is hostile input.
const MAX_SPAN: u32 = 1_048_576;

/// Best-effort layout of an .ods workbook. Empty unless the source is an
/// ODF spreadsheet package the walk can vouch for.
pub(crate) fn extract(source: &Path) -> WorkbookLayout {
    let empty = || WorkbookLayout::from_parts(Vec::new(), Default::default(), 0);
    let mut file = match File::open(source) {
        Ok(file) => file,
        Err(_) => return empty(),
    };
    let Ok(mut archive) = ZipArchive::new(&mut file) else {
        return empty();
    };
    // Content-based gate, like the BIFF walk's CFB magic: only a package
    // carrying the ODF spreadsheet mimetype is walked. The mimetype is the
    // first, uncompressed 46-byte entry of a conforming package; anything
    // else simply never qualifies.
    let is_spreadsheet = archive
        .by_name("mimetype")
        .ok()
        .and_then(|mut entry| {
            let mut magic = [0u8; 46];
            entry.read_exact(&mut magic).ok()?;
            Some(magic == *b"application/vnd.oasis.opendocument.spreadsheet")
        })
        .unwrap_or(false);
    if !is_spreadsheet {
        return empty();
    }
    let Ok(mut content) = archive.by_name("content.xml") else {
        return empty();
    };
    let mut xml = String::new();
    if content.read_to_string(&mut xml).is_err() {
        return empty();
    }
    WorkbookLayout::from_parts(walk(&xml), Default::default(), 0)
}

/// Walks content.xml and returns one layout per `table:table`, in document
/// order — the same order calamine numbers sheets by.
fn walk(xml: &str) -> Vec<SheetLayout> {
    let mut reader = Reader::from_str(xml);
    let mut sheets: Vec<SheetLayout> = Vec::new();
    // Current table state; None outside any table (automatic-styles,
    // named expressions, ...).
    let mut table: Option<TableWalk> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) => match local_name(element.name()) {
                b"table" => {
                    // A table without a name is invisible to calamine too
                    // (it only collects named tables): skip it whole so its
                    // rows never pollute the neighboring sheet's walk.
                    if read_attr(&reader, &element, b"name").is_some() {
                        if let Some(finished) = table.take() {
                            sheets.push(finished.finish());
                        }
                        table = Some(TableWalk::default());
                    }
                }
                b"table-row" => {
                    if let Some(walk) = table.as_mut() {
                        walk.begin_row(&reader, &element);
                    }
                }
                b"table-cell" | b"covered-table-cell" => {
                    if let Some(walk) = table.as_mut() {
                        walk.cell(&reader, &element);
                    }
                }
                _ => {}
            },
            Ok(Event::End(element)) => match local_name(element.name()) {
                b"table-row" => {
                    if let Some(walk) = table.as_mut() {
                        walk.end_row();
                    }
                }
                b"table" => {
                    if let Some(finished) = table.take() {
                        sheets.push(finished.finish());
                    }
                }
                _ => {}
            },
            // Empty elements (`<table:covered-table-cell/>`,
            // `<table:table-row table:number-rows-repeated="2"/>`) carry
            // the same attributes; a repeated empty row is one event.
            Ok(Event::Empty(element)) => match local_name(element.name()) {
                b"table-row" => {
                    if let Some(walk) = table.as_mut() {
                        walk.begin_row(&reader, &element);
                        walk.end_row();
                    }
                }
                b"table-cell" | b"covered-table-cell" => {
                    if let Some(walk) = table.as_mut() {
                        walk.cell(&reader, &element);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            // Malformed XML: whatever was walked so far stands (best-effort).
            Err(_) => break,
            _ => {}
        }
    }
    if let Some(finished) = table.take() {
        sheets.push(finished.finish());
    }
    sheets
}

fn local_name(name: quick_xml::name::QName<'_>) -> &[u8] {
    name.local_name().into_inner()
}

/// The decoded, unescaped value of the attribute whose LOCAL name matches
/// (the ODF namespaces are fixed but their prefixes are not).
fn read_attr<'a>(
    reader: &Reader<&'a [u8]>,
    element: &BytesStart<'a>,
    name: &[u8],
) -> Option<String> {
    element.attributes().flatten().find_map(|attribute| {
        if attribute.key.local_name().as_ref() == name {
            attribute
                .decode_and_unescape_value(reader.decoder())
                .ok()
                .map(|value| value.into_owned())
        } else {
            None
        }
    })
}

/// Attribute as a bounded count, defaulting to 1 (the ODF default).
fn read_count(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
    max: u32,
) -> u32 {
    read_attr(reader, element, name)
        .and_then(|value| value.parse().ok())
        .map_or(1, |count: u32| count.clamp(1, max))
}

/// Per-table accumulator: the merge ranges of one sheet.
struct TableWalk {
    merges: Vec<([u32; 2], [u32; 2])>,
    row: u32,
    column: u32,
    /// The repeat count of the row in flight (1 outside a row).
    row_repeat: u32,
}

impl Default for TableWalk {
    fn default() -> Self {
        Self {
            merges: Vec::new(),
            row: 0,
            column: 0,
            row_repeat: 1,
        }
    }
}

impl TableWalk {
    fn begin_row<'a>(&mut self, reader: &Reader<&'a [u8]>, element: &BytesStart<'a>) {
        self.row_repeat = read_count(reader, element, b"number-rows-repeated", MAX_ROW_REPEAT);
        // A new row restarts the column cursor.
        self.column = 0;
    }

    fn end_row(&mut self) {
        // Merge anchors are recorded only in the first repeat of a row
        // block: a producer repeating a merged row would multiply the same
        // range, and repeated rows are by definition free of new anchors.
        self.row = self.row.saturating_add(self.row_repeat).min(MAX_ROWS);
        self.row_repeat = 1;
    }

    fn cell<'a>(&mut self, reader: &Reader<&'a [u8]>, element: &BytesStart<'a>) {
        let repeat = read_count(reader, element, b"number-columns-repeated", MAX_COLUMN_REPEAT);
        let columns_spanned = read_count(reader, element, b"number-columns-spanned", MAX_SPAN);
        let rows_spanned = read_count(reader, element, b"number-rows-spanned", MAX_SPAN);
        // Anchors are plain `table:table-cell` elements with a span past
        // 1x1; their `covered-table-cell` continuations (and spanless
        // cells) only advance the column cursor.
        if columns_spanned > 1 || rows_spanned > 1 {
            if self.merges.len() < MAX_MERGES_PER_SHEET {
                let last_row = self
                    .row
                    .saturating_add(rows_spanned - 1)
                    .min(MAX_ROWS - 1);
                let last_column = self
                    .column
                    .saturating_add(columns_spanned - 1)
                    .min(MAX_COLUMNS - 1);
                self.merges
                    .push(([self.row, last_row], [self.column, last_column]));
            }
        }
        self.column = self.column.saturating_add(repeat).min(MAX_COLUMNS);
    }

    fn finish(mut self) -> SheetLayout {
        self.merges.sort_unstable();
        self.merges.dedup();
        SheetLayout {
            merges: self.merges,
            ..Default::default()
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn merges_of(content: &str) -> Vec<([u32; 2], [u32; 2])> {
        let mut sheets = walk(content);
        assert_eq!(sheets.len(), 1, "expected exactly one sheet");
        sheets.remove(0).merges
    }

    /// The three fixture shapes: a horizontal title merge, a 2x2 box whose
    /// continuations carry the `covered-table-cell` repeats, and a vertical
    /// merge reached past a repeated empty row block.
    #[test]
    fn reads_the_fixture_merge_shapes() {
        let content = r#"<office:document-content xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:number-columns-spanned="3" table:number-rows-spanned="1" office:value-type="string"><text:p>title</text:p></table:table-cell><table:covered-table-cell table:number-columns-repeated="2"/><table:table-cell table:number-columns-repeated="3"/></table:table-row><table:table-row table:number-rows-repeated="2"><table:table-cell table:number-columns-repeated="6"/></table:table-row><table:table-row><table:table-cell/><table:table-cell table:number-columns-spanned="3" table:number-rows-spanned="2" office:value-type="string"><text:p>box</text:p></table:table-cell><table:covered-table-cell table:number-columns-repeated="2"/><table:table-cell table:number-columns-repeated="2"/></table:table-row><table:table-row><table:table-cell/><table:covered-table-cell table:number-columns-repeated="3"/><table:table-cell table:number-columns-repeated="2"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert_eq!(
            merges_of(content),
            vec![
                ([0, 0], [0, 2]),
                ([3, 4], [1, 3]),
            ]
        );
    }

    /// Self-closing covered cells and self-closing rows keep positions in
    /// step with the element shapes LibreOffice actually writes.
    #[test]
    fn handles_empty_element_shapes() {
        let content = r#"<office:document-content xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:spreadsheet><table:table table:name="S"><table:table-row><table:table-cell/><table:covered-table-cell/><table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="2"/></table:table-row><table:table-row><table:covered-table-cell table:number-columns-repeated="2"/><table:covered-table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert_eq!(merges_of(content), vec![([0, 1], [2, 3])]);
    }

    /// Multiple tables come out in document order, nameless tables are
    /// skipped (calamine never sees them either), and ranges are sorted
    /// and deduplicated.
    #[test]
    fn walks_tables_in_order_and_skips_nameless_ones() {
        let content = r#"<office:document-content xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:spreadsheet><table:table table:name="A"><table:table-row><table:table-cell table:number-columns-spanned="2"/></table:table-row></table:table><table:table><table:table-row><table:table-cell table:number-columns-spanned="2"/></table:table-row></table:table><table:table table:name="B"><table:table-row><table:table-cell table:number-columns-spanned="4" table:number-rows-spanned="2"/><table:covered-table-cell table:number-columns-repeated="3"/></table:table-row><table:table-row><table:covered-table-cell table:number-columns-repeated="4"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
        let sheets = walk(content);
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].merges, vec![([0, 0], [0, 1])]);
        assert_eq!(sheets[1].merges, vec![([0, 1], [0, 3])]);
    }

    /// A merge row repeated by `number-rows-repeated` records the range
    /// once, at the block's first row.
    #[test]
    fn a_repeated_merge_row_records_one_range() {
        let content = r#"<office:document-content xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:spreadsheet><table:table table:name="S"><table:table-row table:number-rows-repeated="3"><table:table-cell table:number-columns-spanned="2"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert_eq!(merges_of(content), vec![([0, 0], [0, 1])]);
    }

    /// Hostile numbers clamp: an oversized repeat cannot wrap the cursor
    /// coordinates back into the sheet, and garbage counts are ignored.
    #[test]
    fn hostile_counts_cannot_move_anchors_out_of_the_grid() {
        let content = r#"<office:document-content xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:spreadsheet><table:table table:name="S"><table:table-row table:number-rows-repeated="999999999999"><table:table-cell table:number-columns-repeated="999999999999" table:number-columns-spanned="999" table:number-rows-spanned="999"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
        let merges = merges_of(content);
        assert_eq!(merges.len(), 1);
        // The clamps hold: both endpoints stay inside the grid.
        assert!(merges[0].0[1] < MAX_ROWS);
        assert!(merges[0].1[1] < MAX_COLUMNS);
        // Garbage parses as the ODF default of 1.
        let content = r#"<office:document-content xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:spreadsheet><table:table table:name="S"><table:table-row><table:table-cell table:number-columns-spanned="many"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
        assert!(merges_of(content).is_empty());
    }
}

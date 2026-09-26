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

use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::Path;

use quick_xml::Reader;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use zip::ZipArchive;

use crate::xls_layout::{ColSpan, RowSpec, SheetLayout, WorkbookLayout};

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
const MAX_STYLES: usize = 65_536;
const MAX_COLUMN_ENTRIES_PER_SHEET: usize = 4_096;
/// BIFF sheets carry 65 536 ROW records at most; the same bound holds the
/// .ods row table.
const MAX_ROW_SPECS_PER_SHEET: usize = 65_536;
/// The converter's implicit row height in points — a custom height equal to
/// it has nothing to say (same rule as the BIFF walk).
const DEFAULT_ROW_HEIGHT_PT: f64 = 15.0;

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

/// Column/row styles from `<office:automatic-styles>` — the style tables
/// the BIFF walk reads from FONT/XF records, sourced from content.xml here.
/// They precede the body in document order, so one forward pass fills them
/// before the tables that reference them.
#[derive(Default)]
struct StyleTablesWalk {
    columns: HashMap<String, ColumnStyle>,
    rows: HashMap<String, RowStyle>,
    /// The `style:style` element in flight: (name, family).
    current: Option<(String, ColumnOrRow)>,
}

#[derive(Clone, Copy, PartialEq)]
enum ColumnOrRow {
    Column,
    Row,
}

/// `style:family="table-column"`: the width the converter maps into `<col>`
/// runs, plus the visibility flag.
#[derive(Clone, Copy)]
struct ColumnStyle {
    /// Width in Excel character units (what the `width=` attribute takes).
    width_chars: f64,
    hidden: bool,
}

/// `style:family="table-row"`: the height in points; `optimal` rows are
/// auto-fit, so their stored height is the producer's measurement, not a
/// formatting decision.
#[derive(Clone, Copy)]
struct RowStyle {
    height_pt: f64,
    optimal: bool,
}

/// One `table:table-column` run: a style applied to `repeat` columns.
struct ColumnEntry {
    repeat: u32,
    style: Option<ColumnStyle>,
}

/// Walks content.xml and returns one layout per `table:table`, in document
/// order — the same order calamine numbers sheets by.
fn walk(xml: &str) -> Vec<SheetLayout> {
    let mut reader = Reader::from_str(xml);
    let mut sheets: Vec<SheetLayout> = Vec::new();
    let mut styles = StyleTablesWalk::default();
    // Current table state; None outside any table (automatic-styles,
    // named expressions, ...).
    let mut table: Option<TableWalk> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) => match local_name(element.name()) {
                b"style" => styles.begin_style(&reader, &element),
                b"table-column-properties" => styles.column_properties(&reader, &element),
                b"table-row-properties" => styles.row_properties(&reader, &element),
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
                b"table-column" => {
                    if let Some(walk) = table.as_mut() {
                        walk.column(&reader, &element, &styles);
                    }
                }
                b"table-row" => {
                    if let Some(walk) = table.as_mut() {
                        walk.begin_row(&reader, &element, &styles);
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
                b"style" => styles.current = None,
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
                b"style" => {
                    styles.begin_style(&reader, &element);
                    styles.current = None;
                }
                b"table-column-properties" => styles.column_properties(&reader, &element),
                b"table-row-properties" => styles.row_properties(&reader, &element),
                b"table-column" => {
                    if let Some(walk) = table.as_mut() {
                        walk.column(&reader, &element, &styles);
                    }
                }
                b"table-row" => {
                    if let Some(walk) = table.as_mut() {
                        walk.begin_row(&reader, &element, &styles);
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

impl StyleTablesWalk {
    fn begin_style(&mut self, reader: &Reader<&[u8]>, element: &BytesStart<'_>) {
        let name = read_attr(reader, element, b"name");
        let family = read_attr(reader, element, b"family");
        self.current = name.zip(family).and_then(|(name, family)| {
            let slot = match family.as_str() {
                "table-column" => ColumnOrRow::Column,
                "table-row" => ColumnOrRow::Row,
                _ => return None,
            };
            Some((name, slot))
        });
    }

    fn column_properties(&mut self, reader: &Reader<&[u8]>, element: &BytesStart<'_>) {
        let Some((name, ColumnOrRow::Column)) = &self.current else {
            return;
        };
        let Some(inches) = read_attr(reader, element, b"column-width").and_then(|value| parse_length(&value))
        else {
            return;
        };
        let hidden = read_attr(reader, element, b"visibility")
            .is_some_and(|value| value == "collapse" || value == "filter");
        if self.columns.len() < MAX_STYLES {
            self.columns.insert(
                name.clone(),
                ColumnStyle {
                    width_chars: inches_to_char_width(inches),
                    hidden,
                },
            );
        }
    }

    fn row_properties(&mut self, reader: &Reader<&[u8]>, element: &BytesStart<'_>) {
        let Some((name, ColumnOrRow::Row)) = &self.current else {
            return;
        };
        let Some(inches) = read_attr(reader, element, b"row-height").and_then(|value| parse_length(&value))
        else {
            return;
        };
        let optimal = read_attr(reader, element, b"use-optimal-row-height")
            .is_some_and(|value| value == "true");
        if self.rows.len() < MAX_STYLES {
            self.rows.insert(
                name.clone(),
                RowStyle {
                    height_pt: inches * 72.0,
                    optimal,
                },
            );
        }
    }
}

/// ODF lengths (`0.889in`, `2.258cm`, `22.58mm`, `64pt`, `1pc`, `96px`)
/// in inches — the common unit both the column and the row mapping need.
fn parse_length(text: &str) -> Option<f64> {
    let digits = text.trim_end_matches(|ch: char| ch.is_ascii_alphabetic());
    let unit = &text[digits.len()..];
    let value: f64 = digits.parse().ok()?;
    let inches = match unit {
        "in" => value,
        "cm" => value / 2.54,
        "mm" => value / 25.4,
        "pt" => value / 72.0,
        "pc" => value / 6.0,
        "px" => value / 96.0,
        _ => return None,
    };
    (inches > 0.0 && inches < 100.0).then_some(inches)
}

/// Excel character width for a physical column width: the app's Calibri 11
/// body font has a 7px maximum-digit width and 5px of padding, so
/// `chars = (px - 5) / 7`. This inverts the mapping LibreOffice applies on
/// its own ods->xlsx export, keeping a converted form's proportions.
fn inches_to_char_width(inches: f64) -> f64 {
    ((inches * 96.0 - 5.0) / 7.0).clamp(0.0, 255.0)
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

/// Per-table accumulator: the merge ranges, column runs and row specs of
/// one sheet.
struct TableWalk {
    merges: Vec<([u32; 2], [u32; 2])>,
    col_entries: Vec<ColumnEntry>,
    rows: Vec<(u32, RowSpec)>,
    row: u32,
    column: u32,
    /// The repeat count of the row in flight (1 outside a row).
    row_repeat: u32,
    /// The spec the row in flight carries (height and/or hidden).
    row_spec: Option<RowSpec>,
}

impl Default for TableWalk {
    fn default() -> Self {
        Self {
            merges: Vec::new(),
            col_entries: Vec::new(),
            rows: Vec::new(),
            row: 0,
            column: 0,
            row_repeat: 1,
            row_spec: None,
        }
    }
}

impl TableWalk {
    fn column<'a>(
        &mut self,
        reader: &Reader<&'a [u8]>,
        element: &BytesStart<'a>,
        styles: &StyleTablesWalk,
    ) {
        if self.col_entries.len() >= MAX_COLUMN_ENTRIES_PER_SHEET {
            return;
        }
        let repeat = read_count(reader, element, b"number-columns-repeated", MAX_COLUMN_REPEAT);
        let style = read_attr(reader, element, b"style-name")
            .and_then(|name| styles.columns.get(&name))
            .copied();
        // Column visibility lives on the element itself when there is no
        // style (`table:visibility="collapse"` on a filtered-out column).
        let hidden = read_attr(reader, element, b"visibility")
            .is_some_and(|value| value == "collapse" || value == "filter");
        let style = style.map(|mut style| {
            style.hidden |= hidden;
            style
        });
        self.col_entries.push(ColumnEntry { repeat, style });
    }

    fn begin_row<'a>(
        &mut self,
        reader: &Reader<&'a [u8]>,
        element: &BytesStart<'a>,
        styles: &StyleTablesWalk,
    ) {
        self.row_repeat = read_count(reader, element, b"number-rows-repeated", MAX_ROW_REPEAT);
        // A new row restarts the column cursor.
        self.column = 0;
        let hidden = read_attr(reader, element, b"visibility")
            .is_some_and(|value| value == "collapse" || value == "filter");
        let height = read_attr(reader, element, b"style-name")
            .and_then(|name| styles.rows.get(&name))
            .copied()
            .and_then(Self::row_height);
        self.row_spec = (height.is_some() || hidden).then_some(RowSpec { height, hidden });
    }

    /// A carried height must be manually set (not the producer's auto-fit
    /// measurement) and say something the conversion cannot guess — the
    /// same filter the BIFF walk applies to ROW records.
    fn row_height(style: RowStyle) -> Option<f64> {
        (!style.optimal
            && style.height_pt > 0.0
            && style.height_pt <= 409.5
            && style.height_pt != DEFAULT_ROW_HEIGHT_PT)
            .then_some((style.height_pt * 100.0).round() / 100.0)
    }

    fn end_row(&mut self) {
        // A repeated row block repeats its spec (a filtered-out band hides
        // every row it covers); merge anchors are still recorded only in
        // the first repeat.
        if let Some(spec) = self.row_spec {
            for offset in 0..self.row_repeat {
                if self.rows.len() >= MAX_ROW_SPECS_PER_SHEET {
                    break;
                }
                let row = self.row.saturating_add(offset).min(MAX_ROWS - 1);
                self.rows.push((row, spec));
            }
        }
        self.row = self.row.saturating_add(self.row_repeat).min(MAX_ROWS);
        self.row_repeat = 1;
        self.row_spec = None;
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
        let cols = column_runs(&self.col_entries);
        self.rows.sort_unstable_by_key(|(row, _)| *row);
        self.rows.dedup_by_key(|(row, _)| *row);
        SheetLayout {
            merges: self.merges,
            cols,
            rows: self.rows,
            ..Default::default()
        }
    }
}

/// Collapses `table:table-column` runs into emitted `<col>` spans. The
/// untouched tail is a repeated default-width run (LibreOffice closes every
/// sheet with thousands of default columns): it is treated as the sheet
/// default and only spans that differ from it — or are hidden — emit.
fn column_runs(entries: &[ColumnEntry]) -> Vec<ColSpan> {
    // Expand into per-column styles, clamped to the grid.
    let mut columns: Vec<Option<ColumnStyle>> = vec![None; MAX_COLUMNS as usize];
    let mut cursor = 0usize;
    for entry in entries {
        for _ in 0..entry.repeat {
            if cursor >= columns.len() {
                break;
            }
            columns[cursor] = entry.style;
            cursor += 1;
        }
    }
    let default_width = columns[columns.len() - 1].map(|style| style.width_chars);
    let uniform_tail = columns[columns.len() - 1]
        .map(|tail| {
            columns[..columns.len() - 1].iter().all(|style| {
                style.is_some_and(|style| {
                    same_width(style.width_chars, tail.width_chars) && style.hidden == tail.hidden
                })
            })
        })
        .unwrap_or(false);
    let default_width = if uniform_tail { None } else { default_width };

    let mut spans = Vec::new();
    let mut index = 0usize;
    while index < columns.len() {
        let Some(style) = columns[index] else {
            index += 1;
            continue;
        };
        if !style.hidden && default_width.is_some_and(|width| same_width(width, style.width_chars)) {
            index += 1;
            continue;
        }
        let start = index;
        while index < columns.len()
            && columns[index]
                .is_some_and(|next| same_width(next.width_chars, style.width_chars) && next.hidden == style.hidden)
        {
            index += 1;
        }
        spans.push(ColSpan {
            first: start as u16,
            last: (index - 1) as u16,
            width: style.width_chars,
            hidden: style.hidden,
        });
    }
    spans
}

/// Column widths compare at the precision they are emitted at.
fn same_width(a: f64, b: f64) -> bool {
    (a - b).abs() < 0.005
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

    /// ODF lengths parse across the unit zoo and refuse nonsense.
    #[test]
    fn lengths_parse_and_refuse() {
        assert!((parse_length("0.889in").unwrap() - 0.889).abs() < 1e-12);
        assert!((parse_length("2.54cm").unwrap() - 1.0).abs() < 1e-12);
        assert!((parse_length("25.4mm").unwrap() - 1.0).abs() < 1e-12);
        assert!((parse_length("72pt").unwrap() - 1.0).abs() < 1e-12);
        assert!((parse_length("6pc").unwrap() - 1.0).abs() < 1e-12);
        assert!((parse_length("96px").unwrap() - 1.0).abs() < 1e-12);
        assert_eq!(parse_length("wide"), None);
        assert_eq!(parse_length("-1in"), None);
        assert_eq!(parse_length("0in"), None);
    }

    /// Physical widths map onto the app's Calibri-11 character units.
    #[test]
    fn inches_map_to_char_widths() {
        assert!((inches_to_char_width(1.85) - 24.657_142_857).abs() < 1e-6);
        // The LibreOffice default column: ~8.46 characters.
        assert!((inches_to_char_width(0.6689) - 8.4592).abs() < 1e-9);
        assert_eq!(inches_to_char_width(0.0), 0.0);
    }

    /// Column runs collapse to the user-set spans; the repeated
    /// default-width tail (and hidden-at-default columns) stay out.
    #[test]
    fn column_runs_drop_the_default_tail() {
        let default = ColumnStyle {
            width_chars: 8.46,
            hidden: false,
        };
        let wide = ColumnStyle {
            width_chars: 24.66,
            hidden: false,
        };
        let hidden = ColumnStyle {
            width_chars: 8.46,
            hidden: true,
        };
        let entries = |styles: Vec<(u32, Option<ColumnStyle>)>| -> Vec<ColumnEntry> {
            styles
                .into_iter()
                .map(|(repeat, style)| ColumnEntry { repeat, style })
                .collect()
        };
        // A=wide, B=default, rest=default tail -> only A emits.
        let spans = column_runs(&entries(vec![
            (1, Some(wide)),
            (1, Some(default)),
            (MAX_COLUMN_REPEAT, Some(default)),
        ]));
        assert_eq!(spans.len(), 1);
        assert_eq!((spans[0].first, spans[0].last), (0, 0));
        assert!((spans[0].width - 24.66).abs() < 1e-9);
        // A hidden column survives even at the default width.
        let spans = column_runs(&entries(vec![
            (1, Some(default)),
            (1, Some(hidden)),
            (MAX_COLUMN_REPEAT, Some(default)),
        ]));
        assert_eq!(spans.len(), 1);
        assert!(spans[0].hidden);
        assert_eq!(spans[0].first, 1);
    }

    /// Row specs: an auto-fit height is not a decision (dropped), a custom
    /// height carries (rounded), hidden rows carry without a height, and
    /// repeated row blocks repeat their spec.
    #[test]
    fn row_specs_filter_and_repeat() {
        assert_eq!(
            TableWalk::row_height(RowStyle {
                height_pt: 30.0,
                optimal: false
            }),
            Some(30.0)
        );
        // Auto-fit and default heights are not formatting decisions.
        assert_eq!(
            TableWalk::row_height(RowStyle {
                height_pt: 15.0,
                optimal: false
            }),
            None
        );
        assert_eq!(
            TableWalk::row_height(RowStyle {
                height_pt: 12.0,
                optimal: true
            }),
            None
        );
        assert_eq!(
            TableWalk::row_height(RowStyle {
                height_pt: 10_000.0,
                optimal: false
            }),
            None
        );
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

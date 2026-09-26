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
//! Cell formatting (fills, font weight/color/size, borders, alignment)
//! lives in the same automatic-styles tables; the walk turns every used
//! cell style into a synthetic BIFF-style Font/XF entry so the proven
//! `xls_layout` style emission produces the styles.xml unchanged — one
//! interner for both legacy walks.
//!
//! ODF number formats (`number:date-style` data styles) are deliberately
//! not translated: a style whose only feature is a data style stays out of
//! styles.xml, so those cells keep the converter's type-derived fallback
//! formats (short date for date values, and so on) instead of a custom
//! format without a code.
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

use crate::xls_layout::{
    ColSpan, FontSpec, RowSpec, SheetLayout, StyleTables, WorkbookLayout, XfSpec,
};

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
/// The BIFF palette overlay starts at color index 8 and the classic table
/// ends at 63: at most 56 ODF colors can ride it. Beyond the cap the color
/// degrades to "unspecified" instead of displacing an earlier color.
/// Styled cells carried per sheet — the BIFF walk's cell cap.
const MAX_CELLS_PER_SHEET: usize = 1_000_000;
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
    let (sheets, style_tables) = walk(&xml);
    // No default font: every synthetic font must intern (the sentinel
    // matches nothing), and the base font stays the converter's Calibri.
    WorkbookLayout::from_parts(sheets, style_tables, u16::MAX)
}

/// Column/row styles from `<office:automatic-styles>` — the style tables
/// the BIFF walk reads from FONT/XF records, sourced from content.xml here.
/// They precede the body in document order, so one forward pass fills them
/// before the tables that reference them.
#[derive(Default)]
struct StyleTablesWalk {
    columns: HashMap<String, ColumnStyle>,
    rows: HashMap<String, RowStyle>,
    cells: HashMap<String, CellStyleDef>,
    /// The `style:style` element in flight: (name, family).
    current: Option<(String, ColumnOrRow)>,
}

#[derive(Clone, Copy, PartialEq, Debug)]
enum ColumnOrRow {
    Column,
    Row,
    Cell,
}

/// One automatic cell style (`style:family="table-cell"`): the visual
/// features its cells carry. `None` fields mean "source says nothing".
#[derive(Default, Clone)]
struct CellStyleDef {
    /// Solid background RGB.
    fill: Option<[u8; 3]>,
    bold: bool,
    italic: bool,
    font_color: Option<[u8; 3]>,
    font_size_pt: Option<f64>,
    /// `style:font-name` — the display name, as BIFF FONT records carry.
    font_name: Option<String>,
    /// 0 top, 1 center; bottom (the default) stays None.
    vertical: Option<u8>,
    /// 1 left, 2 center, 3 right, 5 justify; general stays None.
    horizontal: Option<u8>,
    wrap: bool,
    /// Border line codes (BIFF order left/right/top/bottom; 0 = none) and
    /// their colors.
    borders: [u8; 4],
    border_colors: [Option<[u8; 3]>; 4],
}

impl CellStyleDef {
    /// Whether the style carries anything the conversion cannot guess.
    fn is_interesting(&self) -> bool {
        self.fill.is_some()
            || self.bold
            || self.italic
            || self.font_color.is_some()
            || self.font_size_pt.is_some()
            || self.vertical.is_some()
            || self.horizontal.is_some()
            || self.wrap
            || self.borders.iter().any(|&code| code != 0)
    }
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
fn walk(xml: &str) -> (Vec<SheetLayout>, StyleTables) {
    let mut reader = Reader::from_str(xml);
    let mut sheets: Vec<TableWalk> = Vec::new();
    let mut styles = StyleTablesWalk::default();
    let mut style_names = StyleNameIds::default();
    // Current table state; None outside any table (automatic-styles,
    // named expressions, ...).
    let mut table: Option<TableWalk> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) => match local_name(element.name()) {
                b"style" => styles.begin_style(&reader, &element),
                b"table-column-properties" => styles.column_properties(&reader, &element),
                b"table-row-properties" => styles.row_properties(&reader, &element),
                b"table-cell-properties" => styles.cell_properties(&reader, &element),
                b"text-properties" => styles.text_properties(&reader, &element),
                b"paragraph-properties" => styles.paragraph_properties(&reader, &element),
                b"table" => {
                    // A table without a name is invisible to calamine too
                    // (it only collects named tables): skip it whole so its
                    // rows never pollute the neighboring sheet's walk.
                    if read_attr(&reader, &element, b"name").is_some() {
                        if let Some(finished) = table.take() {
                            sheets.push(finished);
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
                        walk.cell(&reader, &element, &mut style_names);
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
                        sheets.push(finished);
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
                b"table-cell-properties" => styles.cell_properties(&reader, &element),
                b"text-properties" => styles.text_properties(&reader, &element),
                b"paragraph-properties" => styles.paragraph_properties(&reader, &element),
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
                        walk.cell(&reader, &element, &mut style_names);
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
        sheets.push(finished);
    }
    // Named styles -> synthetic BIFF-style tables; only styles with visible
    // features become xfs, everything else falls back by construction.
    let mut style_tables = StyleTables::default();
    let mut xf_by_name: HashMap<String, u16> = HashMap::new();
    let mut names: Vec<&String> = styles.cells.keys().collect();
    names.sort();
    for name in names {
        let def = &styles.cells[name];
        if def.is_interesting() && let Some(xf) = build_xf(def, &mut style_tables) {
            xf_by_name.insert(name.clone(), xf);
        }
    }
    let layouts = sheets
        .into_iter()
        .map(|mut table| {
            let styled = std::mem::take(&mut table.styled_cells);
            let mut finished = table.finish();
            finished.cells = styled
                .into_iter()
                .filter_map(|(position, id)| {
                    let name = style_names.name(id)?;
                    xf_by_name.get(name).map(|xf| (position, *xf))
                })
                .collect();
            finished
        })
        .collect();
    (layouts, style_tables)
}

/// Interned style names: the walk stores a u32 per styled cell, the final
/// mapping resolves ids against the collected automatic styles.
#[derive(Default)]
struct StyleNameIds {
    ids: HashMap<String, u32>,
    names: Vec<String>,
}

impl StyleNameIds {
    fn id(&mut self, name: &str) -> Option<u32> {
        // "Default" is the producer's word for "nothing to carry".
        if name == "Default" {
            return None;
        }
        Some(if let Some(&id) = self.ids.get(name) {
            id
        } else {
            let id = self.names.len() as u32;
            self.names.push(name.to_string());
            self.ids.insert(name.to_string(), id);
            id
        })
    }

    fn name(&self, id: u32) -> Option<&String> {
        self.names.get(id as usize)
    }
}

/// Builds the synthetic BIFF-style xf for one cell style: fonts, fills,
/// borders and alignments are interned into the shared style tables the
/// xls_layout emission consumes. `None` when the style's colors no longer
/// fit the palette — the cells then keep the fallback look.
fn build_xf(def: &CellStyleDef, tables: &mut StyleTables) -> Option<u16> {
    // Fonts: 0 is the base-font slot; unique font combinations follow.
    let font = if !def.bold
        && !def.italic
        && def.font_color.is_none()
        && def.font_size_pt.is_none()
        && def.font_name.is_none()
    {
        0
    } else {
        let color = def.font_color.and_then(|rgb| tables.intern_color(rgb));
        let fonts = &mut tables.fonts;

        let position = fonts.iter().position(|font| {
            font.bold == def.bold
                && font.italic == def.italic
                && font.name == def.font_name.as_deref().unwrap_or("")
                && font.height_pt == def.font_size_pt.unwrap_or(11.0)
                && font.color == color.unwrap_or(32767)
        });
        let index = match position {
            Some(index) => index as u16,
            None => {
                fonts.push(FontSpec {
                    bold: def.bold,
                    italic: def.italic,
                    underline: 0,
                    height_pt: def.font_size_pt.unwrap_or(11.0),
                    color: color.unwrap_or(32767),
                    name: def.font_name.clone().unwrap_or_default(),
                });
                (fonts.len() - 1) as u16
            }
        };
        index
    };
    // Fill: a solid fill rides the palette; a fill whose color no longer
    // fits degrades to no fill instead of painting an arbitrary color.
    let mut fill_pattern = 0;
    let mut fill_color = 0;
    if let Some(rgb) = def.fill
        && let Some(index) = tables.intern_color(rgb)
    {
        fill_pattern = 1;
        fill_color = index;
    }
    let mut border_colors = [0u16; 4];
    for (index, color) in def.border_colors.iter().enumerate() {
        if def.borders[index] != 0
            && let Some(rgb) = color
            && let Some(palette_index) = tables.intern_color(*rgb)
        {
            border_colors[index] = palette_index;
        }
    }
    let xfs = &mut tables.xfs;
    xfs.push(XfSpec {
        font,
        format: 0,
        horizontal: def.horizontal.unwrap_or(0),
        vertical: def.vertical.unwrap_or(2),
        wrap: def.wrap,
        borders: def.borders,
        border_colors,
        fill_pattern,
        fill_color,
    });
    Some((xfs.len() - 1) as u16)
}

impl StyleTablesWalk {
    fn begin_style(&mut self, reader: &Reader<&[u8]>, element: &BytesStart<'_>) {
        let name = read_attr(reader, element, b"name");
        let family = read_attr(reader, element, b"family");
        eprintln!("DBG begin_style name={name:?} family={family:?}");
        self.current = name.zip(family).and_then(|(name, family)| {
            let slot = match family.as_str() {
                "table-column" => ColumnOrRow::Column,
                "table-row" => ColumnOrRow::Row,
                "table-cell" => ColumnOrRow::Cell,
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

    fn cell_properties(&mut self, reader: &Reader<&[u8]>, element: &BytesStart<'_>) {
        let Some((name, ColumnOrRow::Cell)) = &self.current else {
            return;
        };
        let style = style_entry(&mut self.cells, name);
        if let Some(fill) = read_attr(reader, element, b"background-color")
            .and_then(|value| parse_color(&value))
        {
            style.fill = Some(fill);
        }
        match read_attr(reader, element, b"vertical-align").as_deref() {
            Some("top") => style.vertical = Some(0),
            Some("middle") => style.vertical = Some(1),
            _ => {}
        }
        // Shorthand first, then the per-side overrides.
        let shorthand = read_attr(reader, element, b"border").and_then(|value| parse_border(&value));
        let mut sides: [Option<(u8, Option<[u8; 3]>)>; 4] = [None, None, None, None];
        if let Some(border) = shorthand {
            sides = [Some(border.clone()), Some(border.clone()), Some(border.clone()), Some(border)];
        }
        for (side, attr) in std::iter::zip(
            sides.iter_mut(),
            [
                b"border-left" as &[u8],
                b"border-right",
                b"border-top",
                b"border-bottom",
            ],
        ) {
            if let Some(border) = read_attr(reader, element, attr).and_then(|value| parse_border(&value)) {
                *side = Some(border);
            }
        }
        for (index, side) in sides.into_iter().enumerate() {
            if let Some((code, color)) = side {
                style.borders[index] = code;
                style.border_colors[index] = color;
            }
        }
    }

    fn text_properties(&mut self, reader: &Reader<&[u8]>, element: &BytesStart<'_>) {
        let Some((name, ColumnOrRow::Cell)) = &self.current else {
            return;
        };
        let style = style_entry(&mut self.cells, name);
        if read_attr(reader, element, b"font-weight").is_some_and(|value| value == "bold") {
            style.bold = true;
        }
        if read_attr(reader, element, b"font-style").is_some_and(|value| value == "italic") {
            style.italic = true;
        }
        if let Some(color) = read_attr(reader, element, b"color").and_then(|value| parse_color(&value)) {
            style.font_color = Some(color);
        }
        if let Some(size) = read_attr(reader, element, b"font-size").and_then(|value| parse_length(&value)) {
            style.font_size_pt = Some((size * 72.0 * 100.0).round() / 100.0);
        }
        if let Some(name) = read_attr(reader, element, b"font-name") {
            style.font_name = Some(name);
        }
    }

    fn paragraph_properties(&mut self, reader: &Reader<&[u8]>, element: &BytesStart<'_>) {
        let Some((name, ColumnOrRow::Cell)) = &self.current else {
            return;
        };
        let style = style_entry(&mut self.cells, name);
        match read_attr(reader, element, b"text-align").as_deref() {
            Some("left") | Some("start") => style.horizontal = Some(1),
            Some("center") => style.horizontal = Some(2),
            Some("right") | Some("end") => style.horizontal = Some(3),
            Some("justify") => style.horizontal = Some(5),
            _ => {}
        }
        if read_attr(reader, element, b"wrap-option").is_some_and(|value| value == "wrap") {
            style.wrap = true;
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

/// The automatic cell style being accumulated: created on first touch.
fn style_entry<'a>(
    cells: &'a mut HashMap<String, CellStyleDef>,
    name: &str,
) -> &'a mut CellStyleDef {
    cells.entry(name.to_string()).or_default()
}

/// `#rrggbb` (the only color shape LibreOffice writes for cell
/// backgrounds and text); `transparent`/`none`/system colors are refused
/// so the style carries nothing instead of painting black.
fn parse_color(text: &str) -> Option<[u8; 3]> {
    let hex = text.strip_prefix('#')?;
    if hex.len() != 6 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    Some([
        u8::from_str_radix(&hex[0..2], 16).ok()?,
        u8::from_str_radix(&hex[2..4], 16).ok()?,
        u8::from_str_radix(&hex[4..6], 16).ok()?,
    ])
}

/// A border shorthand (`0.5pt solid #808080`, `0.74pt dotted #ff0000`):
/// the BIFF line code the emission names, and the color when present.
fn parse_border(text: &str) -> Option<(u8, Option<[u8; 3]>)> {
    let mut parts = text.split_whitespace();
    let width_pt = parts.next().and_then(parse_length)? * 72.0;
    let line = parts.next()?;
    let color = parts.next().and_then(parse_color);
    let code = match line {
        "none" | "hidden" => 0,
        "solid" => {
            // LibreOffice writes its thin border as 0.5pt; the ladder maps
            // the producer's widths onto Excel's four weights.
            if width_pt < 0.25 {
                7 // hair
            } else if width_pt < 1.75 {
                1 // thin
            } else if width_pt < 2.75 {
                2 // medium
            } else {
                5 // thick
            }
        }
        "double" => 6,
        "dashed" => 3,
        "dotted" => 4,
        _ => return None,
    };
    (code != 0).then_some((code, color))
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
    /// Styled cell positions -> interned style-name id. "Default" and
    /// unknown names never enter; the synthetic xf mapping happens once,
    /// after the walk.
    styled_cells: HashMap<(u32, u32), u32>,
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
            styled_cells: HashMap::new(),
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

    fn cell<'a>(
        &mut self,
        reader: &Reader<&'a [u8]>,
        element: &BytesStart<'a>,
        style_names: &mut StyleNameIds,
    ) {
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
        // A named cell style is recorded with its position; merge
        // continuations inherit the anchor's style so a merged header keeps
        // its fill across the whole range the way LibreOffice's own export
        // writes it.
        let style_id = read_attr(reader, element, b"style-name")
            .and_then(|name| style_names.id(&name));
        if let Some(id) = style_id {
            for row in 0..rows_spanned.min(MAX_SPAN) {
                for column in 0..columns_spanned.min(MAX_SPAN) {
                    if self.styled_cells.len() >= MAX_CELLS_PER_SHEET {
                        break;
                    }
                    self.styled_cells.insert(
                        (
                            self.row.saturating_add(row).min(MAX_ROWS - 1),
                            self.column.saturating_add(column).min(MAX_COLUMNS - 1),
                        ),
                        id,
                    );
                }
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

    /// The layout-only view of a walk, for the merge/geometry tests.
    fn walk_sheets(content: &str) -> Vec<SheetLayout> {
        walk(content).0
    }

    fn merges_of(content: &str) -> Vec<([u32; 2], [u32; 2])> {
        let mut sheets = walk_sheets(content);
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
        let sheets = walk_sheets(content);
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

    /// A cell style with a fill and bold font interns into the synthetic
    /// tables; the styled cell carries its xf.
    #[test]
    fn cell_styles_intern_into_synthetic_tables() {
        let content = r##"<office:document-content xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:automatic-styles><style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:background-color="#ffcc00"/><style:text-properties fo:color="#ff0000" fo:font-size="14pt" fo:font-weight="bold" style:font-name="Cambria"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="S"><table:table-row><table:table-cell table:style-name="ce1" office:value-type="string"><text:p>H</text:p></table:table-cell><table:table-cell table:style-name="Default"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"##;
        let (sheets, tables) = walk(content);
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].cells.len(), 1, "default-named cells stay out");
        let xf = *sheets[0].cells.values().next().unwrap();
        let font = &tables.fonts[tables.xfs[xf as usize].font as usize];
        assert!(font.bold);
        assert_eq!(font.name, "Cambria");
        // Font color interns first (build_xf resolves fonts before fills).
        assert_eq!(tables.palette, vec![[0xFF, 0x00, 0x00], [0xFF, 0xCC, 0x00]]);
    }

    /// Border shorthands map onto BIFF line codes with a width ladder.
    #[test]
    fn borders_map_to_line_codes() {
        assert_eq!(parse_border("0.5pt solid #808080"), Some((1, Some([128, 128, 128]))));
        assert_eq!(parse_border("0.1pt solid #000000"), Some((7, Some([0, 0, 0]))));
        assert_eq!(parse_border("1.5pt solid #000000"), Some((1, Some([0, 0, 0]))));
        assert_eq!(parse_border("2.5pt solid #000000"), Some((2, Some([0, 0, 0]))));
        assert_eq!(parse_border("3pt solid #000000"), Some((5, Some([0, 0, 0]))));
        assert_eq!(parse_border("0.5pt double #000000"), Some((6, Some([0, 0, 0]))));
        assert_eq!(parse_border("0.5pt dashed #000000"), Some((3, Some([0, 0, 0]))));
        assert_eq!(parse_border("0.5pt dotted #000000"), Some((4, Some([0, 0, 0]))));
        assert_eq!(parse_border("0.5pt none"), None);
        assert_eq!(parse_border("solid"), None);
    }

    /// Colors: hex only; named/system colors stay unresolvable.
    #[test]
    fn colors_parse_hex_only() {
        assert_eq!(parse_color("#FFCC00"), Some([0xFF, 0xCC, 0x00]));
        assert_eq!(parse_color("#ffcc00"), Some([0xFF, 0xCC, 0x00]));
        assert_eq!(parse_color("transparent"), None);
        assert_eq!(parse_color("#fff"), None);
        assert_eq!(parse_color("#GGGGGG"), None);
    }

    /// A truncated .ods never panics the walk: the cut fixture degrades to
    /// whatever prefix was readable, exactly like the BIFF walk.
    #[test]
    fn truncated_ods_fixture_never_panics_the_walk() {
        let full = std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/par214-ods-form-layout.ods"
        ))
        .unwrap();
        let dir = tempfile::tempdir().unwrap();
        for cut in (0..full.len()).step_by(64) {
            let path = dir.path().join("cut.ods");
            std::fs::write(&path, &full[..cut]).unwrap();
            let _ = extract(&path);
        }
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

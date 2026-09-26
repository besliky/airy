//! Legacy .xls layout and style carry-over for the convert path (BUG-1602).
//!
//! calamine exposes cell values of a .xls book but nothing about its shape:
//! merged ranges (MERGEDCELLS), column widths (COLINFO) and cell formats
//! (FONT/XF/FORMAT/PALETTE plus per-cell format indexes) are all dropped, so
//! a tabular form — a two-tier merged header, bordered boxes, sized columns —
//! converts into a flat grid of default-width cells. LibreOffice keeps all of
//! it when converting the same file, and this module brings the converter to
//! parity for the parts that make a form look like a form.
//!
//! The extraction is a second, independent BIFF walk over the same Workbook
//! stream the BUG-1600 string overlay reads (shared compound-file reader and
//! record walker). Workbook globals contribute FONT, XF, FORMAT and PALETTE
//! records (the style tables) plus BOUNDSHEET offsets; each sheet substream
//! contributes MERGEDCELLS, COLINFO and the XF index of every cell record
//! (RK/MULRK/NUMBER/LABELSST/FORMULA/BLANK/MULBLANK/...), plus the ROW
//! heights and hidden flags that keep a form's header rows tall and its
//! filtered rows out of sight (BUG-1607). FONT names and FORMAT codes are
//! decoded only after the walk: CODEPAGE is "last record wins", exactly
//! like in the string overlay.
//!
//! Everything is best-effort, exactly like the string overlay: any structural
//! surprise yields an empty layout and the conversion falls back to the plain
//! value-only output. BIFF5 books are gated out (their XF record has a
//! different layout); records are only read through bounds-checked slices and
//! every collection is capped, so hostile input cannot size allocations.

use std::collections::HashMap;
use std::io::Read as _;
use std::path::Path;

use codepage::to_encoding;
use encoding_rs::Encoding;

use crate::legacy_xls::{self, CompoundFile, Record};
use crate::visuals::colors::INDEXED_COLORS;

/// The one named cell style every SpreadsheetML styles.xml must carry. A
/// macro, not a const, because every writer embeds it inside a `concat!`
/// literal. IronCalc's importer indexes the `cellStyles` section
/// unconditionally and panics on a book without it (BUG-1659), which used to
/// leave every legacy-converted (.ods/.xls) workbook dead for the formula
/// engine — the same books Excel flags as needing repair.
macro_rules! default_cell_styles_xml {
    () => {
        r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#
    };
}
// The re-export serves the recalc.rs engine fixture (test-only today); the
// in-module writers below reach the macro through its textual scope.
#[allow(unused_imports)]
pub(crate) use default_cell_styles_xml;

/// BIFF8 cell records that carry an XF index at offset 4 (row @0, column @2).
const CELL_XF_RECORDS: [u16; 8] = [
    0x0006, // FORMULA
    0x00FD, // LABELSST
    0x0203, // NUMBER
    0x0201, // BLANK
    0x0205, // BOOLERR
    0x027E, // RK
    0x0204, // LABEL (BIFF5-style string; rare in BIFF8 but legal)
    0x00D6, // RSTRING
];
/// MULRK/MULBLANK: row @0, colFirst @2, then per column (XF, RK) or XF
/// entries, colLast in the final 2 bytes.
const REC_MULRK: u16 = 0x00BD;
const REC_MULBLANK: u16 = 0x00BE;
const REC_MERGEDCELLS: u16 = 0x00E5;
const REC_COLINFO: u16 = 0x007D;
const REC_ROW: u16 = 0x0208;
const REC_DEFAULTROWHEIGHT: u16 = 0x0225;
const REC_FONT: u16 = 0x0031;
const REC_XF: u16 = 0x00E0;
const REC_FORMAT: u16 = 0x041E;
const REC_PALETTE: u16 = 0x0092;
const REC_BOUNDSHEET: u16 = 0x0085;
const REC_CODEPAGE: u16 = 0x0042;
const REC_BOF: u16 = 0x0809;

/// Caps: a legitimate form carries tens of merges and a handful of styles.
/// The caps only bind for hostile or pathological input, where the walk
/// degrades to a partial layout instead of buffering unbounded state.
const MAX_MERGES_PER_SHEET: usize = 65_536;
const MAX_COLINFO_RECORDS_PER_SHEET: usize = 4_096;
const MAX_CELLS_PER_SHEET: usize = 1_000_000;
/// BIFF8 sheets have at most 65 536 rows, and one ROW record each.
const MAX_ROWS_PER_SHEET: usize = 65_536;
/// Excel's hard maximum row height, in twips (409.5 pt); anything larger
/// is hostile input, not a form.
const MAX_ROW_HEIGHT_TWIPS: u16 = 8_190;
/// The converter's implicit xlsx default row height, in points. A book
/// whose own default differs carries it into `<sheetFormatPr>`; a row
/// whose custom height equals it has nothing to say.
const DEFAULT_ROW_HEIGHT_PT: f64 = 15.0;
const MAX_FONTS: usize = 65_536;
const MAX_XFS: usize = 65_536;
const MAX_FORMATS: usize = 65_536;
/// BIFF sheets have at most 256 columns.
const MAX_COLUMNS: usize = 256;

/// BIFF border line style codes map 1:1 onto the xlsx border style names
/// ([MS-XLS] 2.5.11 BorderStyle; codes past slantDashDot do not exist).
const BORDER_STYLES: [&str; 14] = [
    "none",
    "thin",
    "medium",
    "dashed",
    "dotted",
    "thick",
    "double",
    "hair",
    "mediumDashed",
    "dashDot",
    "mediumDashDot",
    "dashDotDot",
    "mediumDashDotDot",
    "slantDashDot",
];

/// Extracted layout of one sheet, aligned with calamine's sheet order.
#[derive(Default)]
pub(crate) struct SheetLayout {
    /// Merged ranges, 0-based inclusive (row first/last, column first/last),
    /// sorted and deduplicated.
    pub(crate) merges: Vec<([u16; 2], [u16; 2])>,
    /// User-set column spans (merged runs), ascending.
    pub(crate) cols: Vec<ColSpan>,
    /// Source XF index per styled cell. Plain unformatted cells are absent:
    /// they keep the converter's fallback look, which is what they had.
    pub(crate) cells: HashMap<(u32, u32), u16>,
    /// Rows worth carrying over (custom height and/or hidden), ascending.
    pub(crate) rows: Vec<(u32, RowSpec)>,
    /// The sheet's default row height in points when the book states one
    /// other than the converter's implicit 15pt.
    pub(crate) default_row_height: Option<f64>,
}

impl SheetLayout {
    /// Nothing worth carrying over for this sheet.
    pub(crate) fn is_empty(&self) -> bool {
        self.merges.is_empty()
            && self.cols.is_empty()
            && self.cells.is_empty()
            && self.rows.is_empty()
            && self.default_row_height.is_none()
    }

    /// What to carry over for one row, if anything.
    pub(crate) fn row_spec(&self, row: u32) -> Option<&RowSpec> {
        self.rows
            .binary_search_by_key(&row, |(row, _)| *row)
            .ok()
            .map(|index| &self.rows[index].1)
    }
}

/// ROW record subset worth carrying over for one row.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct RowSpec {
    /// Manually-set row height in points, for the `ht=` attribute (always
    /// with `customHeight="1"`). None keeps the sheet default — a hidden
    /// row at the producer's default height needs only the flag.
    pub(crate) height: Option<f64>,
    pub(crate) hidden: bool,
}

/// One emitted `<col min max width hidden>` run.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct ColSpan {
    pub(crate) first: u16,
    pub(crate) last: u16,
    /// Width in characters, as written into the xlsx `width=` attribute.
    pub(crate) width: f64,
    pub(crate) hidden: bool,
}

/// FONT record subset the converter maps into an xlsx `<font>`.
struct FontSpec {
    bold: bool,
    italic: bool,
    /// BIFF underline code: 0 none, 1 single, 2 double, 33/34 accounting.
    underline: u8,
    /// dyHeight is in twips (1/20 pt); kept in points for `<sz val=>`.
    height_pt: f64,
    color: u16,
    name: String,
}

/// XF record subset the converter maps into an xlsx `<xf>`.
struct XfSpec {
    font: u16,
    format: u16,
    /// Horizontal: 0 general, 1 left, 2 center, 3 right, 4 fill, 5 justify,
    /// 6 center-continuous ([MS-XLS] 2.5.8 AlignH).
    horizontal: u8,
    /// Vertical: 0 top, 1 center, 2 bottom, 3 justify, 4 distributed.
    vertical: u8,
    wrap: bool,
    /// Border line styles, order left/right/top/bottom (0 = none).
    borders: [u8; 4],
    /// Border color indexes in the same order.
    border_colors: [u16; 4],
    /// 0 = no fill, 1 = solid; 2..=18 are the standard pattern fills.
    fill_pattern: u8,
    fill_color: u16,
}

/// Style tables from the workbook globals substream.
#[derive(Default)]
struct StyleTables {
    fonts: Vec<FontSpec>,
    xfs: Vec<XfSpec>,
    /// Custom number formats as (id, code) in FORMAT record order.
    formats: Vec<(u16, String)>,
    /// PALETTE overrides: RGB for color indexes 8, 9, ... ([MS-XLS] 2.4.125:
    /// the overrides start at index 8; lower slots are fixed system colors).
    palette: Vec<[u8; 3]>,
}

/// Best-effort layout of a legacy workbook. Empty unless the source is a
/// BIFF8 OLE2 compound document the walk can vouch for.
pub(crate) struct WorkbookLayout {
    sheets: Vec<SheetLayout>,
    styles: StyleTables,
    /// The modal font across a book's cells is its body font: the base font
    /// plain converted cells should keep. Bold header fonts lose the vote by
    /// design — "unstyled" simply means "like the surrounding data".
    default_font: u16,
}

impl WorkbookLayout {
    /// Read the layout tables out of a legacy .xls. Never fails: everything
    /// unexpected yields an empty layout and the previous conversion shape.
    pub(crate) fn extract(source: &Path) -> Self {
        let mut layout = Self {
            sheets: Vec::new(),
            styles: StyleTables::default(),
            default_font: 0,
        };
        // Content-based, like the string overlay: only an OLE2 compound
        // document can carry BIFF records; zip-based sources skip the walk.
        let Ok(mut file) = std::fs::File::open(source) else {
            return layout;
        };
        let mut magic = [0u8; 8];
        if file.read_exact(&mut magic).is_err() || magic != legacy_xls::CFB_MAGIC {
            return layout;
        }
        let Some((sheets, styles, default_font)) = walk(file) else {
            return layout;
        };
        layout.sheets = sheets;
        layout.styles = styles;
        layout.default_font = default_font;
        layout
    }

    /// Layout for the sheet at calamine's index (BOUNDSHEET order).
    pub(crate) fn sheet(&self, index: usize) -> Option<&SheetLayout> {
        self.sheets.get(index).filter(|sheet| !(*sheet).is_empty())
    }

    pub(crate) fn default_font(&self) -> u16 {
        self.default_font
    }

    /// A layout with the given sheets, for emission tests.
    #[cfg(test)]
    pub(crate) fn for_test(sheets: Vec<SheetLayout>) -> Self {
        Self {
            sheets,
            styles: StyleTables::default(),
            default_font: 0,
        }
    }
}

// ---------------------------------------------------------------------------
// The BIFF walk
// ---------------------------------------------------------------------------

fn walk(file: std::fs::File) -> Option<(Vec<SheetLayout>, StyleTables, u16)> {
    let mut compound = CompoundFile::open(file)?;
    let stream = compound.read_workbook_stream()?;
    parse_stream(&stream)
}

fn parse_stream(stream: &[u8]) -> Option<(Vec<SheetLayout>, StyleTables, u16)> {
    // BIFF8 only: the XF record has a different layout (16 bytes, packed
    // fields) in BIFF5, and its cell records carry no merge information in
    // the shapes used here.
    let first = Record::at(stream, 0)?;
    if first.id != REC_BOF
        || legacy_xls::read_u16(first.body) != Some(legacy_xls::BIFF8_BOF_VERSION)
    {
        return None;
    }

    let mut styles = StyleTables::default();
    // Raw FONT/FORMAT bodies, decoded once the final codepage is known.
    let mut font_bodies: Vec<Vec<u8>> = Vec::new();
    let mut format_bodies: Vec<Vec<u8>> = Vec::new();
    let mut codepage = 1200u16;
    // Sheet substream start offsets in BOUNDSHEET order — the same order
    // calamine numbers sheets by.
    let mut sheet_starts: Vec<usize> = Vec::new();
    let mut sheet_accs: Vec<SheetAccumulator> = Vec::new();

    let mut position = 0usize;
    while let Some(record) = Record::at(stream, position) {
        let in_globals = sheet_starts.first().is_none_or(|start| position < *start);
        if in_globals {
            match record.id {
                REC_CODEPAGE => {
                    if let Some(value) = legacy_xls::read_u16(record.body) {
                        codepage = value;
                    }
                }
                REC_FONT => {
                    if font_bodies.len() < MAX_FONTS {
                        font_bodies.push(record.body.to_vec());
                    }
                }
                REC_XF => match parse_xf(record.body) {
                    Some(xf) => {
                        if styles.xfs.len() < MAX_XFS {
                            styles.xfs.push(xf);
                        }
                    }
                    // A truncated XF cannot desynchronize the record walk
                    // (length-based), but it ends the trustworthy table.
                    None => styles.xfs.clear(),
                },
                REC_FORMAT => {
                    if format_bodies.len() < MAX_FORMATS {
                        format_bodies.push(record.body.to_vec());
                    }
                }
                REC_PALETTE => styles.palette = parse_palette(record.body),
                REC_BOUNDSHEET => {
                    if let Some(offset) = legacy_xls::read_u32(record.body).map(|v| v as usize) {
                        sheet_starts.push(offset);
                        sheet_accs.push(SheetAccumulator::default());
                    }
                }
                _ => {}
            }
        } else if let Some(sheet) = sheet_starts
            .iter()
            .rposition(|start| *start <= position)
            .and_then(|index| sheet_accs.get_mut(index))
        {
            sheet.accumulate(record.id, record.body, &styles);
        }
        position = record.end;
    }

    // The codepage is final now: decode the string-bearing tables, work out
    // the body font, and fold the raw accumulators into sheet layouts.
    // Compressed (fHighByte=0) strings are byte-per-character; a UTF-16
    // codepage (1200, what Excel and LibreOffice usually write) has no byte
    // table for them, so those fall back to Latin-1.
    let encoding = to_encoding(codepage).filter(|encoding| encoding.is_single_byte());
    styles.fonts = font_bodies
        .iter()
        .enumerate()
        .flat_map(|(index, body)| {
            // BIFF font index 4 is reserved and never referenced ([MS-XLS]
            // XLUnicodeFont): the 5th FONT record belongs to index 5. A
            // placeholder keeps every later record's index what cells say.
            let gap = (index == 4).then(|| FontSpec {
                bold: false,
                italic: false,
                underline: 0,
                height_pt: 11.0,
                color: 32767,
                name: String::new(),
            });
            gap.into_iter()
                .chain(parse_font(body, encoding).into_iter().collect::<Vec<_>>())
        })
        .collect();
    styles.formats = format_bodies
        .iter()
        .filter_map(|body| parse_format(body, encoding))
        .collect();
    let default_font = default_font_of(&sheet_accs);
    let sheets = sheet_accs
        .into_iter()
        .map(|acc| acc.finish(&styles, default_font))
        .collect();
    Some((sheets, styles, default_font))
}

/// Per-sheet record accumulator: merges, COLINFO spans, cell XF indexes,
/// and the ROW table.
#[derive(Default)]
struct SheetAccumulator {
    merges: Vec<([u16; 2], [u16; 2])>,
    colinfos: Vec<(u16, u16, u16, bool)>,
    cells: HashMap<(u32, u32), u16>,
    /// Raw ROW records: row -> (height twips, grbit). One per row in
    /// practice; a repeat simply overwrites.
    rows: HashMap<u16, (u16, u16)>,
    /// DEFAULTROWHEIGHT's height, in twips, when the sheet states one.
    default_row_height_twips: Option<u16>,
    /// Font histogram over value cells only — blank/bordered decoration
    /// cells must not out-vote the body font.
    font_votes: HashMap<u16, u32>,
}

impl SheetAccumulator {
    fn accumulate(&mut self, id: u16, body: &[u8], styles: &StyleTables) {
        match id {
            REC_MERGEDCELLS => self.push_merges(body),
            REC_ROW => self.push_row(body),
            REC_DEFAULTROWHEIGHT => self.note_default_row_height(body),
            REC_COLINFO => {
                if self.colinfos.len() < MAX_COLINFO_RECORDS_PER_SHEET
                    && let Some(span) = parse_colinfo(body)
                {
                    self.colinfos.push(span);
                }
            }
            REC_MULRK | REC_MULBLANK => {
                if self.cells.len() >= MAX_CELLS_PER_SHEET {
                    return;
                }
                let is_value = id == REC_MULRK;
                let stride = if id == REC_MULRK { 6 } else { 2 };
                let (Some(row), Some(first)) = (
                    body.get(0..2)
                        .and_then(|b| b.try_into().ok())
                        .map(u16_from_le),
                    body.get(2..4)
                        .and_then(|b| b.try_into().ok())
                        .map(u16_from_le),
                ) else {
                    return;
                };
                // 4 bytes of (row, colFirst) + 2 bytes of colLast surround
                // the per-column entries; anything shorter is truncated.
                let pairs = body.len().saturating_sub(6) / stride;
                for index in 0..pairs.min(MAX_COLUMNS) {
                    let offset = 4 + index * stride;
                    let Some(xf) = body
                        .get(offset..offset + 2)
                        .and_then(|b| b.try_into().ok())
                        .map(u16_from_le)
                    else {
                        break;
                    };
                    if is_value {
                        self.vote_font(styles, xf);
                    }
                    if styles.cell_is_interesting(xf, None) {
                        self.cells
                            .insert((u32::from(row), u32::from(first) + index as u32), xf);
                    }
                }
            }
            _ => {
                if !CELL_XF_RECORDS.contains(&id) || self.cells.len() >= MAX_CELLS_PER_SHEET {
                    return;
                }
                let (Some(row), Some(column), Some(xf)) = (
                    body.get(0..2)
                        .and_then(|b| b.try_into().ok())
                        .map(u16_from_le),
                    body.get(2..4)
                        .and_then(|b| b.try_into().ok())
                        .map(u16_from_le),
                    body.get(4..6)
                        .and_then(|b| b.try_into().ok())
                        .map(u16_from_le),
                ) else {
                    return;
                };
                if !matches!(id, 0x0201 | 0x00BE) {
                    // Only value records vote: a bordered blank is
                    // decoration, not the body voice of the sheet.
                    self.vote_font(styles, xf);
                }
                if styles.cell_is_interesting(xf, None) {
                    self.cells.insert((u32::from(row), u32::from(column)), xf);
                }
            }
        }
    }

    fn vote_font(&mut self, styles: &StyleTables, xf: u16) {
        if let Some(spec) = styles.xfs.get(xf as usize) {
            *self.font_votes.entry(spec.font).or_insert(0) += 1;
        }
    }

    /// MERGEDCELLS body: u16 range count, then that many inclusive
    /// (rowFirst, rowLast, colFirst, colLast) u16 quadruples. Truncated tails
    /// keep the ranges decoded so far.
    fn push_merges(&mut self, body: &[u8]) {
        let Some(count) = body
            .get(0..2)
            .and_then(|b| b.try_into().ok())
            .map(u16_from_le)
        else {
            return;
        };
        for index in 0..usize::from(count) {
            if self.merges.len() >= MAX_MERGES_PER_SHEET {
                return;
            }
            let Some(range) = body
                .get(2 + index * 8..10 + index * 8)
                .and_then(|bytes| TryInto::<&[u8; 8]>::try_into(bytes).ok())
            else {
                return;
            };
            let (row_first, row_last, col_first, col_last) = (
                u16_from_le(range[0..2].try_into().unwrap()),
                u16_from_le(range[2..4].try_into().unwrap()),
                u16_from_le(range[4..6].try_into().unwrap()),
                u16_from_le(range[6..8].try_into().unwrap()),
            );
            if row_first <= row_last && col_first <= col_last {
                self.merges
                    .push(([row_first, row_last], [col_first, col_last]));
            }
        }
    }

    /// ROW record (16 bytes): row number @0, height in twips @6 (its high
    /// bit marks "the default height"), row flags @12 — bit 5 hides the
    /// row, bit 6 marks a manually-set height (fUnsynced). Layout verified
    /// against xlrd's ROW parsing and the fixture's raw records.
    fn push_row(&mut self, body: &[u8]) {
        if self.rows.len() >= MAX_ROWS_PER_SHEET {
            return;
        }
        let Some(fields) = body.first_chunk::<14>() else {
            return;
        };
        let row = u16_from_le(fields[0..2].try_into().expect("checked above"));
        let height = u16_from_le(fields[6..8].try_into().expect("checked above"));
        let flags = u16_from_le(fields[12..14].try_into().expect("checked above"));
        self.rows.insert(row, (height, flags));
    }

    /// DEFAULTROWHEIGHT body: flags @0 (bit 0 fUnsynced, bit 4 fDyZero),
    /// default height in twips @2 — order verified against xlrd's handling
    /// and the fixture bytes.
    fn note_default_row_height(&mut self, body: &[u8]) {
        let Some(fields) = body.first_chunk::<4>() else {
            return;
        };
        let height = u16_from_le(fields[2..4].try_into().expect("checked above"));
        if height > 0 && height <= MAX_ROW_HEIGHT_TWIPS {
            self.default_row_height_twips = Some(height);
        }
    }

    fn finish(mut self, styles: &StyleTables, default_font: u16) -> SheetLayout {
        self.cells
            .retain(|_, xf| styles.cell_is_interesting(*xf, Some(default_font)));
        let default_twips = self.default_row_height_twips;
        let mut merges = self.merges;
        merges.sort_unstable();
        merges.dedup();
        let mut rows: Vec<(u32, RowSpec)> = self
            .rows
            .into_iter()
            .filter_map(|(row, (height, flags))| {
                row_spec(height, flags, default_twips).map(|spec| (u32::from(row), spec))
            })
            .collect();
        // HashMap order is arbitrary; the xlsx schema wants ascending rows.
        rows.sort_unstable_by_key(|(row, _)| *row);
        SheetLayout {
            merges,
            cols: column_spans(&self.colinfos),
            cells: self.cells,
            rows,
            default_row_height: default_twips
                .filter(|&twips| f64::from(twips) != DEFAULT_ROW_HEIGHT_PT * 20.0)
                .map(|twips| f64::from(twips) / 20.0),
        }
    }
}

/// Which ROW records are worth carrying over: hidden rows always (they must
/// not reappear in the conversion); a height only when it was manually set
/// (fUnsynced) and says something the conversion cannot guess — equal to the
/// sheet default or to our own implicit 15pt, or missing entirely, it would
/// only add noise attributes.
fn row_spec(height_twips: u16, flags: u16, default_twips: Option<u16>) -> Option<RowSpec> {
    let hidden = flags & 0x0020 != 0;
    let height = height_twips & 0x7FFF; // the high bit marks "default height"
    let custom = flags & 0x0040 != 0
        && height != 0
        && height <= MAX_ROW_HEIGHT_TWIPS
        && Some(height) != default_twips
        && f64::from(height) / 20.0 != DEFAULT_ROW_HEIGHT_PT;
    if custom {
        return Some(RowSpec {
            height: Some(f64::from(height) / 20.0),
            hidden,
        });
    }
    hidden.then_some(RowSpec {
        height: None,
        hidden: true,
    })
}

fn u16_from_le(bytes: &[u8; 2]) -> u16 {
    u16::from_le_bytes(*bytes)
}

/// COLINFO body: colFirst, colLast, width (1/256 char), XF index, flags
/// (bit 0 = hidden).
fn parse_colinfo(body: &[u8]) -> Option<(u16, u16, u16, bool)> {
    let fields: &[u8; 10] = body.first_chunk()?;
    let first = u16_from_le(fields[0..2].try_into().ok()?);
    let last = u16_from_le(fields[2..4].try_into().ok()?);
    let width = u16_from_le(fields[4..6].try_into().ok()?);
    let hidden = fields[8] & 0x01 != 0;
    if first > last {
        return None;
    }
    Some((first, last, width, hidden))
}

/// Collapse COLINFO records into emitted `<col>` runs. Producers write a
/// catch-all record for the untouched tail (LibreOffice: columns 3..=255 at
/// the default width); those must not become 256 customWidth columns, so the
/// tail's width is treated as the sheet default and only spans that differ
/// from it (or are hidden) are emitted. A catch-all that nothing differs
/// from is a genuinely uniform sheet — then every column is emitted, because
/// "default" there is the producer's width, not ours.
fn column_spans(colinfos: &[(u16, u16, u16, bool)]) -> Vec<ColSpan> {
    let mut columns: Vec<Option<(u16, bool)>> = vec![None; MAX_COLUMNS];
    for &(first, last, width, hidden) in colinfos {
        for column in first..=last.min((MAX_COLUMNS - 1) as u16) {
            columns[column as usize] = Some((width, hidden));
        }
    }
    let default_width = columns[MAX_COLUMNS - 1].map(|(width, _)| width);
    let uniform_catch_all = columns[MAX_COLUMNS - 1].is_some_and(|catch_all| {
        columns[..MAX_COLUMNS - 1]
            .iter()
            .all(|entry| entry.map(|column| column == catch_all).unwrap_or(false))
    });
    let default_width = if uniform_catch_all {
        None
    } else {
        default_width
    };

    let mut spans = Vec::new();
    let mut index = 0;
    while index < MAX_COLUMNS {
        let Some((width, hidden)) = columns[index] else {
            index += 1;
            continue;
        };
        if !hidden && Some(width) == default_width {
            index += 1;
            continue;
        }
        let start = index;
        while index < MAX_COLUMNS && columns[index].is_some_and(|(w, h)| w == width && h == hidden)
        {
            index += 1;
        }
        spans.push(ColSpan {
            first: start as u16,
            last: (index - 1) as u16,
            width: f64::from(width) / 256.0,
            hidden,
        });
    }
    spans
}

/// FONT record ([MS-XLS] 2.4.50): height (twips), grbit, color, weight,
/// escapement, then underline/family/charset bytes and the name at byte 14
/// as an XLUnicodeString with a 1-byte length. Bold is NOT a grbit bit in
/// BIFF8 — it lives in the weight field (700 = bold); LibreOffice and Excel
/// both write bold headers as grbit=0 weight=700. Verified against xlrd and
/// LibreOffice output.
fn parse_font(body: &[u8], encoding: Option<&'static Encoding>) -> Option<FontSpec> {
    let head: &[u8; 13] = body.first_chunk()?;
    let height = u16_from_le(head[0..2].try_into().ok()?);
    let grbit = u16_from_le(head[2..4].try_into().ok()?);
    let color = u16_from_le(head[4..6].try_into().ok()?);
    let weight = u16_from_le(head[6..8].try_into().ok()?);
    let underline = head[10];
    let name = legacy_xls::parse_short_xl_unicode_string(
        body.get(14..)?,
        encoding.unwrap_or(encoding_rs::WINDOWS_1252),
    )
    .unwrap_or_default();
    Some(FontSpec {
        bold: weight >= 600,
        italic: grbit & 0x0002 != 0,
        underline,
        height_pt: f64::from(height) / 20.0,
        color,
        name,
    })
}

/// XF record (BIFF8, 20 bytes). Layout verified against xlrd's formatting.py
/// and LibreOffice output: font index @0, number format @2, protection and
/// parent style @4, alignment byte @6 (horizontal 0-2, wrap 3, vertical 4-6),
/// rotation @7, indent/shrink @8, attribute-usage flags @9, then two border/
/// background words @10 and @14 and a closing u16 @18.
fn parse_xf(body: &[u8]) -> Option<XfSpec> {
    let words: &[u8; 20] = body.first_chunk()?;
    let u16_at = |offset: usize| -> Option<u16> {
        Some(u16_from_le(words.get(offset..offset + 2)?.try_into().ok()?))
    };
    let u32_at = |offset: usize| -> Option<u32> {
        Some(u32::from_le_bytes(
            words.get(offset..offset + 4)?.try_into().ok()?,
        ))
    };
    let align = words[6];
    let border1 = u32_at(10)?;
    let border2 = u32_at(14)?;
    let closing = u16_at(18)?;
    let line = |shift: u32| ((border1 >> shift) & 0x0F) as u8;
    Some(XfSpec {
        font: u16_at(0)?,
        format: u16_at(2)?,
        horizontal: align & 0x07,
        vertical: (align >> 4) & 0x07,
        wrap: align & 0x08 != 0,
        borders: [line(0), line(4), line(8), line(12)],
        border_colors: [
            ((border1 >> 16) & 0x7F) as u16,
            ((border1 >> 23) & 0x7F) as u16,
            (border2 & 0x7F) as u16,
            ((border2 >> 7) & 0x7F) as u16,
        ],
        fill_pattern: ((border2 >> 26) & 0x3F) as u8,
        fill_color: closing & 0x7F,
    })
}

/// FORMAT record: format id, then an XLUnicodeString (2-byte length, flags,
/// characters). The code is already escaped the way BIFF stores it
/// (`yyyy\-mm\-dd`) and must reach the xlsx unchanged.
fn parse_format(body: &[u8], encoding: Option<&'static Encoding>) -> Option<(u16, String)> {
    let head: &[u8; 4] = body.first_chunk()?;
    let id = u16_from_le(head[0..2].try_into().ok()?);
    let count = usize::from(u16_from_le(head[2..4].try_into().ok()?));
    let flags = *body.get(4)?;
    let bytes_per_char = if flags & 0x1 != 0 { 2 } else { 1 };
    let bytes = body.get(5..5 + count.checked_mul(bytes_per_char)?)?;
    let code = if flags & 0x1 != 0 {
        legacy_xls::decode_utf16le(bytes)
    } else {
        legacy_xls::decode_single_byte(bytes, encoding)
    };
    Some((id, code))
}

/// PALETTE record: entry count, then that many RGB entries (red, green,
/// blue and a reserved zero byte) overriding palette indexes 8, 9, ...
fn parse_palette(body: &[u8]) -> Vec<[u8; 3]> {
    let Some(count) = body
        .get(0..2)
        .and_then(|b| b.try_into().ok())
        .map(u16_from_le)
    else {
        return Vec::new();
    };
    if usize::from(count) > MAX_COLUMNS / 4 {
        // 64 entries is the real maximum; anything beyond is hostile.
        return Vec::new();
    }
    body.get(2..2 + usize::from(count) * 4)
        .map(|bytes| {
            bytes
                .as_chunks::<4>()
                .0
                .iter()
                .map(|rgb| [rgb[0], rgb[1], rgb[2]])
                .collect()
        })
        .unwrap_or_default()
}

/// The modal font across a book's value cells is its body font: the font
/// plain converted cells should keep. Header fonts lose the vote by design.
fn default_font_of(accs: &[SheetAccumulator]) -> u16 {
    let mut total: HashMap<u16, u32> = HashMap::new();
    for sheet in accs {
        for (&font, count) in &sheet.font_votes {
            *total.entry(font).or_insert(0) += count;
        }
    }
    total
        .into_iter()
        .max_by_key(|(font, count)| (*count, std::cmp::Reverse(*font)))
        .map(|(font, _)| font)
        .unwrap_or(0)
}

impl StyleTables {
    /// Whether one cell's format carries something worth emitting: a fill,
    /// any border, non-default alignment, a non-general number format
    /// (dates!), or a decorated font. `default_font` (known only after the
    /// walk) additionally requires decorated fonts to differ from the body
    /// font, so the book's plain body style stays out of styles.xml.
    fn cell_is_interesting(&self, xf: u16, default_font: Option<u16>) -> bool {
        let Some(spec) = self.xfs.get(xf as usize) else {
            return false;
        };
        if spec.fill_pattern != 0
            || spec.borders.iter().any(|&style| style != 0)
            || spec.horizontal != 0
            || spec.vertical != 2
            || spec.wrap
            || !self.format_is_general(spec.format)
        {
            return true;
        }
        self.font_is_decorated(spec.font) && default_font.is_none_or(|default| spec.font != default)
    }

    fn font_is_decorated(&self, font: u16) -> bool {
        self.fonts
            .get(font as usize)
            .is_some_and(|spec| spec.bold || spec.italic || spec.underline != 0)
    }

    /// "General-like" number format: builtin 0, or a FORMAT record spelling
    /// General (LibreOffice writes its default cells as custom 164 =
    /// "General"; those must not count as number formatting).
    fn format_is_general(&self, format: u16) -> bool {
        format == 0
            || self
                .formats
                .iter()
                .any(|(id, code)| *id == format && code.eq_ignore_ascii_case("general"))
    }

    /// Resolve a BIFF color index to `RRGGBB`. System slots (64, 65, 32767)
    /// resolve to None; the caller omits the color attribute.
    fn color_hex(&self, index: u16) -> Option<String> {
        match index {
            0..=63 => {
                if index >= 8
                    && let Some(rgb) = self.palette.get(index as usize - 8)
                {
                    return Some(format!("{:02X}{:02X}{:02X}", rgb[0], rgb[1], rgb[2]));
                }
                INDEXED_COLORS
                    .get(index as usize)
                    .map(|hex| (*hex).to_string())
            }
            _ => None,
        }
    }

    fn custom_format(&self, format: u16) -> Option<&str> {
        self.formats
            .iter()
            .find(|(id, _)| *id == format)
            .map(|(_, code)| code.as_str())
    }
}

// ---------------------------------------------------------------------------
// Emission: intern source XF indexes into xlsx styles.xml entries
// ---------------------------------------------------------------------------

/// One composed cellXfs entry: everything the xlsx xf needs, resolved.
struct ComposedXf {
    font: usize,
    fill: usize,
    border: usize,
    format: u16,
    horizontal: u8,
    vertical: u8,
    wrap: bool,
}

/// Interner turning source XF indexes into cellXfs entries. Sheet XML is
/// built first (cells reference cellXfs by index), so `cell_style` runs
/// during sheet serialization and `styles_xml` only afterwards.
pub(crate) struct StyleInterner<'a> {
    layout: &'a WorkbookLayout,
    /// source XF index -> emitted cellXfs index; None = unknown source XF,
    /// the cell keeps the converter's fallback style.
    xfs: HashMap<u16, Option<usize>>,
    composed: Vec<ComposedXf>,
    fonts: Vec<u16>,
    fills: Vec<(u8, Option<String>)>,
    borders: Vec<([u8; 4], [Option<String>; 4])>,
    /// Custom number formats actually referenced, in FORMAT order.
    num_fmts: Vec<u16>,
}

impl<'a> StyleInterner<'a> {
    pub(crate) fn new(layout: &'a WorkbookLayout) -> Self {
        Self {
            layout,
            xfs: HashMap::new(),
            composed: Vec::new(),
            fonts: Vec::new(),
            fills: Vec::new(),
            borders: Vec::new(),
            num_fmts: Vec::new(),
        }
    }

    /// cellXfs index for a source XF.
    pub(crate) fn cell_style(&mut self, source_xf: u16) -> Option<usize> {
        let composed = match self.xfs.get(&source_xf) {
            Some(cached) => *cached,
            None => {
                let fresh = self.compose(source_xf);
                self.xfs.insert(source_xf, fresh);
                fresh
            }
        }?;
        Some(composed)
    }

    fn compose(&mut self, source_xf: u16) -> Option<usize> {
        let spec = self.layout.styles.xfs.get(source_xf as usize)?;
        // Copy the fields out first: interning below borrows self mutably.
        let (font_index, format) = (spec.font, spec.format);
        let (fill_pattern, fill_color) = (spec.fill_pattern, spec.fill_color);
        let (borders, border_colors) = (spec.borders, spec.border_colors);
        let (horizontal, vertical, wrap) = (spec.horizontal, spec.vertical, spec.wrap);

        let font = self.intern_font(font_index);
        let fill = self.intern_fill(fill_pattern, fill_color);
        let border = self.intern_border(borders, border_colors);
        if format != 0
            && self.layout.styles.custom_format(format).is_some()
            && !self.num_fmts.contains(&format)
        {
            self.num_fmts.push(format);
        }
        self.composed.push(ComposedXf {
            font,
            fill,
            border,
            format,
            horizontal,
            vertical,
            wrap,
        });
        // After the four fallback styles (general, short date, date+time,
        // elapsed time).
        Some(self.composed.len() - 1 + 4)
    }

    /// Emitted font index for a source font (0 = the base/body font slot).
    fn intern_font(&mut self, font: u16) -> usize {
        // Font 0 is always the base font slot; a source font matching the
        // body font — or missing from the table — maps there.
        if font == self.layout.default_font()
            || self.layout.styles.fonts.get(font as usize).is_none()
        {
            return 0;
        }
        if let Some(index) = self.fonts.iter().position(|&seen| seen == font) {
            return index + 1;
        }
        self.fonts.push(font);
        self.fonts.len()
    }

    /// Emitted fill index (0 = none, 1 = gray125, custom fills from 2).
    fn intern_fill(&mut self, pattern: u8, color: u16) -> usize {
        if pattern == 0 {
            return 0;
        }
        // A solid fill without a resolvable color paints nothing visible;
        // keep those cells fill-less instead of defaulting to some color.
        let rgb = if pattern == 1 {
            self.layout.styles.color_hex(color)
        } else {
            None
        };
        let entry = (pattern, rgb);
        if let Some(index) = self.fills.iter().position(|seen| *seen == entry) {
            return index + 2;
        }
        self.fills.push(entry);
        self.fills.len() + 1
    }

    /// Emitted border index (0 = the default borderless one).
    fn intern_border(&mut self, styles: [u8; 4], colors: [u16; 4]) -> usize {
        if styles.iter().all(|&style| style == 0) {
            return 0;
        }
        let resolved: [Option<String>; 4] = std::array::from_fn(|side| {
            if styles[side] == 0 {
                None
            } else {
                self.layout.styles.color_hex(colors[side])
            }
        });
        let entry = (styles, resolved);
        if let Some(index) = self.borders.iter().position(|seen| *seen == entry) {
            return index + 1;
        }
        self.borders.push(entry);
        self.borders.len()
    }

    /// The complete styles.xml: the converter's fallback styles (general,
    /// short date, date+time, elapsed time — indexes 0-3) followed by every
    /// style the walk picked up.
    pub(crate) fn styles_xml(&self) -> String {
        if self.composed.is_empty() {
            return BASE_STYLES_XML.into();
        }
        let styles = &self.layout.styles;
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        );
        if !self.num_fmts.is_empty() {
            xml.push_str(&format!(r#"<numFmts count="{}">"#, self.num_fmts.len()));
            for &format in &self.num_fmts {
                let code = styles.custom_format(format).unwrap_or("General");
                xml.push_str(&format!(
                    r#"<numFmt numFmtId="{format}" formatCode="{}"/>"#,
                    escape(code),
                ));
            }
            xml.push_str("</numFmts>");
        }
        xml.push_str(&format!(r#"<fonts count="{}">"#, self.fonts.len() + 1));
        xml.push_str(&self.base_font_xml());
        for &font in &self.fonts {
            xml.push_str(&self.font_xml(font));
        }
        xml.push_str("</fonts>");
        xml.push_str(&format!(
            r#"<fills count="{}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>"#,
            self.fills.len() + 2,
        ));
        for &(pattern, ref rgb) in &self.fills {
            let color = rgb
                .as_deref()
                .map(|hex| format!(r#"<fgColor rgb="FF{hex}"/>"#))
                .unwrap_or_default();
            xml.push_str(&format!(
                r#"<fill><patternFill patternType="{}">{color}</patternFill></fill>"#,
                FILL_NAMES.get(pattern as usize).copied().unwrap_or("solid"),
            ));
        }
        xml.push_str("</fills>");
        xml.push_str(&format!(r#"<borders count="{}">"#, self.borders.len() + 1));
        xml.push_str(r#"<border><left/><right/><top/><bottom/><diagonal/></border>"#);
        for &(side_styles, ref side_colors) in &self.borders {
            xml.push_str("<border>");
            for (side, (style, color)) in std::iter::zip(
                ["left", "right", "top", "bottom"],
                std::iter::zip(side_styles, side_colors),
            ) {
                if style != 0 {
                    let color_xml = color
                        .as_deref()
                        .map(|hex| format!(r#"<color rgb="FF{hex}"/>"#))
                        .unwrap_or_default();
                    let name = BORDER_STYLES.get(style as usize).copied().unwrap_or("thin");
                    xml.push_str(&format!(r#"<{side} style="{name}">{color_xml}</{side}>"#));
                } else {
                    xml.push_str(&format!("<{side}/>"));
                }
            }
            xml.push_str("<diagonal/></border>");
        }
        xml.push_str("</borders>");
        xml.push_str(
            r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
        );
        xml.push_str(&format!(r#"<cellXfs count="{}">"#, self.composed.len() + 4));
        xml.push_str(FALLBACK_CELL_XFS);
        for entry in &self.composed {
            let alignment = alignment_xml(entry.horizontal, entry.vertical, entry.wrap);
            xml.push_str(&format!(
                r#"<xf numFmtId="{}" fontId="{}" fillId="{}" borderId="{}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"{}>{}</xf>"#,
                entry.format,
                entry.font,
                entry.fill,
                entry.border,
                if alignment.is_some() {
                    r#" applyAlignment="1""#
                } else {
                    ""
                },
                alignment.unwrap_or_default(),
            ));
        }
        xml.push_str("</cellXfs>");
        xml.push_str(default_cell_styles_xml!());
        xml.push_str("</styleSheet>");
        xml
    }

    /// The base font: the book's body font when known, Calibri 11 otherwise
    /// (the minimal converter's font, kept for non-.xls sources).
    fn base_font_xml(&self) -> String {
        let default = self.layout.default_font();
        match self.layout.styles.fonts.get(default as usize) {
            Some(spec) if !spec.name.is_empty() => self.font_xml(default),
            _ => r#"<font><sz val="11"/><name val="Calibri"/></font>"#.into(),
        }
    }

    fn font_xml(&self, font: u16) -> String {
        let Some(spec) = self.layout.styles.fonts.get(font as usize) else {
            return r#"<font><sz val="11"/><name val="Calibri"/></font>"#.into();
        };
        let mut xml = String::from("<font>");
        if spec.bold {
            xml.push_str("<b/>");
        }
        if spec.italic {
            xml.push_str("<i/>");
        }
        match spec.underline {
            1 => xml.push_str(r#"<u val="single"/>"#),
            2 => xml.push_str(r#"<u val="double"/>"#),
            33 => xml.push_str(r#"<u val="singleAccounting"/>"#),
            34 => xml.push_str(r#"<u val="doubleAccounting"/>"#),
            _ => {}
        }
        xml.push_str(&format!(r#"<sz val="{}"/>"#, format_size(spec.height_pt)));
        if let Some(hex) = self.layout.styles.color_hex(spec.color) {
            xml.push_str(&format!(r#"<color rgb="FF{hex}"/>"#));
        }
        if !spec.name.is_empty() {
            xml.push_str(&format!(r#"<name val="{}"/>"#, escape(&spec.name)));
        }
        xml.push_str("</font>");
        xml
    }
}

/// Filling pattern names shared by BIFF and the xlsx pattern enum; solid is
/// the only one a form realistically uses, the rest pass through by name.
const FILL_NAMES: [&str; 19] = [
    "none",
    "solid",
    "mediumGray",
    "darkGray",
    "lightGray",
    "darkHorizontal",
    "darkVertical",
    "darkDown",
    "darkUp",
    "darkGrid",
    "darkTrellis",
    "lightHorizontal",
    "lightVertical",
    "lightDown",
    "lightUp",
    "lightGrid",
    "lightTrellis",
    "gray125",
    "gray0625",
];

/// xlsx `<alignment>` for non-default alignments only (default: general
/// horizontal, bottom vertical, no wrap).
fn alignment_xml(horizontal: u8, vertical: u8, wrap: bool) -> Option<String> {
    let horizontal = match horizontal {
        1 => "left",
        2 => "center",
        3 => "right",
        4 => "fill",
        5 => "justify",
        6 => "centerContinuous",
        _ => "",
    };
    let vertical = match vertical {
        0 => "top",
        1 => "center",
        3 => "justify",
        4 => "distributed",
        _ => "",
    };
    if horizontal.is_empty() && vertical.is_empty() && !wrap {
        return None;
    }
    let mut xml = String::new();
    if !horizontal.is_empty() {
        xml.push_str(&format!(r#" horizontal="{horizontal}""#));
    }
    if !vertical.is_empty() {
        xml.push_str(&format!(r#" vertical="{vertical}""#));
    }
    if wrap {
        xml.push_str(r#" wrapText="1""#);
    }
    Some(format!("<alignment{xml}/>"))
}

/// Font size without a trailing decimal point (14, not 14).
fn format_size(height_pt: f64) -> String {
    let rounded = (height_pt * 100.0).round() / 100.0;
    if (rounded - rounded.trunc()).abs() < f64::EPSILON {
        format!("{}", rounded as i64)
    } else {
        format!("{rounded}")
    }
}

/// XML attribute/text escaping for the pieces emitted here.
fn escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// The styles.xml of the minimal converter (no .xls styles found):
/// xf 0 general, xf 1 short date (numFmt 14), xf 2 date+time (numFmt 22),
/// xf 3 elapsed time (numFmt 46, what ODF `PT…` durations land on), closed
/// by the mandatory `<cellStyles>` section (BUG-1659).
const BASE_STYLES_XML: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
"#,
    r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    r#"<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>"#,
    r#"<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>"#,
    r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"#,
    r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
    r#"<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
    r#"<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#,
    r#"<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#,
    r#"<xf numFmtId="46" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>"#,
    default_cell_styles_xml!(),
    "</styleSheet>",
);

/// The fallback cellXfs entries (verbatim prefix of BASE_STYLES_XML).
const FALLBACK_CELL_XFS: &str = concat!(
    r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
    r#"<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#,
    r#"<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#,
    r#"<xf numFmtId="46" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#,
);

#[cfg(test)]
mod tests {
    use super::*;

    /// LibreOffice writes a catch-all COLINFO for the untouched tail; only
    /// the columns that actually differ from it become `<col>` runs.
    #[test]
    fn column_spans_drop_the_default_width_catch_all() {
        // The committed fixture's exact records: A=8, B=32, C=15, rest=8.68.
        let spans = column_spans(&[
            (0, 0, 2049, false),
            (1, 1, 8193, false),
            (2, 2, 3841, false),
            (3, 255, 2222, false),
        ]);
        assert_eq!(
            spans,
            vec![
                ColSpan {
                    first: 0,
                    last: 0,
                    width: 2049.0 / 256.0,
                    hidden: false
                },
                ColSpan {
                    first: 1,
                    last: 1,
                    width: 8193.0 / 256.0,
                    hidden: false
                },
                ColSpan {
                    first: 2,
                    last: 2,
                    width: 3841.0 / 256.0,
                    hidden: false
                },
            ]
        );
    }

    /// BUG-1659: styles.xml always carries the `<cellStyles>` section —
    /// IronCalc's importer indexes it unconditionally and panics on a book
    /// without it, dead-ending every converted workbook.
    #[test]
    fn styles_xml_always_carries_a_cell_styles_section() {
        let layout = WorkbookLayout::for_test(Vec::new());
        let styles = StyleInterner::new(&layout).styles_xml();
        assert!(styles.contains(
            r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#
        ));
        assert!(styles.ends_with("</cellStyles></styleSheet>"));
        // Schema order: the named styles come after cellXfs.
        assert!(styles.find("</cellXfs>").unwrap() < styles.find("<cellStyles").unwrap());
    }

    /// A genuinely uniform sheet (every column at the same custom width) has
    /// no "untouched tail" to filter: one run for the whole width, or the
    /// form would fall back to the app's default columns.
    #[test]
    fn a_uniform_catch_all_is_not_treated_as_the_default() {
        let spans = column_spans(&[(0, 255, 7680, false)]);
        assert_eq!(
            spans,
            vec![ColSpan {
                first: 0,
                last: 255,
                width: 30.0,
                hidden: false
            }]
        );
    }

    /// Excel-style records without a catch-all: every COLINFO is user-set.
    #[test]
    fn records_without_a_catch_all_are_all_emitted() {
        let spans = column_spans(&[(0, 5, 4096, false)]);
        assert_eq!(
            spans,
            vec![ColSpan {
                first: 0,
                last: 5,
                width: 16.0,
                hidden: false
            }]
        );
    }

    /// Hidden columns survive even at the catch-all width — and the
    /// catch-all stays filtered (its columns differ only by the hidden flag
    /// on one column, so it is still the "untouched tail" record).
    #[test]
    fn hidden_columns_are_kept() {
        let spans = column_spans(&[(0, 255, 2222, false), (4, 4, 2222, true)]);
        assert_eq!(
            spans,
            vec![ColSpan {
                first: 4,
                last: 4,
                width: 2222.0 / 256.0,
                hidden: true
            }]
        );
    }

    /// Overlapping records: the later one wins per column, and runs only
    /// merge when the effective (width, hidden) pair is identical.
    #[test]
    fn later_records_override_earlier_columns() {
        let spans = column_spans(&[
            (0, 255, 2222, false),
            (1, 3, 8193, false),
            (2, 2, 3841, false),
        ]);
        assert_eq!(
            spans,
            vec![
                ColSpan {
                    first: 1,
                    last: 1,
                    width: 8193.0 / 256.0,
                    hidden: false
                },
                ColSpan {
                    first: 2,
                    last: 2,
                    width: 3841.0 / 256.0,
                    hidden: false
                },
                ColSpan {
                    first: 3,
                    last: 3,
                    width: 8193.0 / 256.0,
                    hidden: false
                },
            ]
        );
    }

    /// Invalid ranges are refused at parse time.
    #[test]
    fn inverted_colinfo_ranges_are_refused() {
        assert!(parse_colinfo(&[5, 0, 0, 0, 1, 0, 0, 0, 0, 0]).is_none());
    }

    /// MERGEDCELLS keeps the ranges decoded before a truncated tail.
    #[test]
    fn truncated_merge_record_keeps_complete_ranges_only() {
        let mut acc = SheetAccumulator::default();
        let mut body = 2u16.to_le_bytes().to_vec();
        body.extend_from_slice(&[0, 0, 0, 0, 0, 0, 2, 0]); // A1:C1 complete
        body.extend_from_slice(&[2, 0, 4]); // truncated tail
        acc.push_merges(&body);
        assert_eq!(acc.merges, vec![([0, 0], [0, 2])]);
    }

    /// Inverted merge ranges (first > last) are dropped, valid ones kept.
    #[test]
    fn inverted_merge_ranges_are_dropped() {
        let mut acc = SheetAccumulator::default();
        let mut body = 2u16.to_le_bytes().to_vec();
        body.extend_from_slice(&[5, 0, 3, 0, 0, 0, 2, 0]); // rows 5..3 invalid
        body.extend_from_slice(&[2, 0, 3, 0, 0, 0, 1, 0]); // A3:B4 valid
        acc.push_merges(&body);
        assert_eq!(acc.merges, vec![([2, 3], [0, 1])]);
    }

    /// A ROW record body (16 bytes): row @0, height twips @6, flags @12 —
    /// bit 5 hidden, bit 6 manually-set height. Truncated bodies are
    /// ignored, and a later record for the same row overwrites the first.
    #[test]
    fn row_records_fill_the_table_last_write_wins() {
        let mut acc = SheetAccumulator::default();
        let body = |number: u16, height: u16, flags: u16| {
            let mut bytes = number.to_le_bytes().to_vec();
            bytes.extend_from_slice(&[0u8; 4]); // colMic, colMac
            bytes.extend_from_slice(&height.to_le_bytes());
            bytes.extend_from_slice(&[0u8; 4]); // irwMac, reserved
            bytes.extend_from_slice(&flags.to_le_bytes());
            bytes.extend_from_slice(&[0u8; 2]); // ixfe
            bytes
        };
        acc.push_row(&body(3, 600, 0x0140));
        acc.push_row(&body(3, 900, 0x0100));
        acc.push_row(&body(3, 600, 0x0140)[..8]); // truncated: ignored
        assert_eq!(acc.rows.get(&3), Some(&(900, 0x0100)));
    }

    /// Custom heights carry over with their points value; a custom height
    /// equal to the sheet default (or to our implicit 15pt) is noise and
    /// stays out.
    #[test]
    fn custom_heights_are_kept_only_when_they_differ_from_defaults() {
        // The fixture's tall header: 600 twips = 30pt, fUnsynced.
        assert_eq!(
            row_spec(600, 0x0140, Some(300)),
            Some(RowSpec {
                height: Some(30.0),
                hidden: false
            })
        );
        // Custom but equal to the sheet default: nothing to add.
        assert_eq!(row_spec(300, 0x0040, Some(300)), None);
        // Custom but equal to the converter's implicit default.
        assert_eq!(row_spec(300, 0x0040, None), None);
        // No default height known: a differing custom height is kept.
        assert_eq!(
            row_spec(255, 0x0040, None),
            Some(RowSpec {
                height: Some(12.75),
                hidden: false
            })
        );
    }

    /// Hidden rows survive even at the default height, without gaining a
    /// customHeight; a custom height and hidden combine into one spec.
    #[test]
    fn hidden_rows_are_kept_with_or_without_a_height() {
        assert_eq!(
            row_spec(300, 0x0120, Some(300)),
            Some(RowSpec {
                height: None,
                hidden: true
            })
        );
        assert_eq!(
            row_spec(600, 0x0160, Some(300)),
            Some(RowSpec {
                height: Some(30.0),
                hidden: true
            })
        );
    }

    /// Auto-fit rows and hostile heights stay out: a zero height, a height
    /// past Excel's 409.5pt maximum, and the plain default-height flag.
    #[test]
    fn plain_and_hostile_heights_are_dropped() {
        // LibreOffice writes every row; untouched ones mean nothing.
        assert_eq!(row_spec(300, 0x0100, Some(300)), None);
        // Height bit unset, zero height: hidden would still carry, custom not.
        assert_eq!(row_spec(0, 0x0040, Some(300)), None);
        // Past the 409.5pt maximum: not a form, hostile input.
        assert_eq!(row_spec(9_000, 0x0140, Some(300)), None);
        // The "row has the default height" high bit on its own.
        assert_eq!(row_spec(0x8300, 0x0100, Some(300)), None);
    }

    /// DEFAULTROWHEIGHT carries the sheet default (flags @0, height @2) but
    /// only when it differs from the converter's implicit 15pt — 300 twips
    /// (15pt, what both fixtures' producers write) means nothing to say,
    /// 255 twips (12.75pt, the Excel 97-2003 default) does.
    #[test]
    fn the_book_default_row_height_is_kept_when_it_differs() {
        let mut acc = SheetAccumulator::default();
        acc.note_default_row_height(&[0x00, 0x00, 0xFF, 0x00]);
        assert_eq!(acc.default_row_height_twips, Some(255));
        let layout = acc.finish(&StyleTables::default(), 0);
        assert_eq!(layout.default_row_height, Some(12.75));
        assert!(!layout.is_empty());

        let mut acc = SheetAccumulator::default();
        acc.note_default_row_height(&[0x00, 0x00, 0x2C, 0x01]); // 300tw = 15pt
        let layout = acc.finish(&StyleTables::default(), 0);
        assert_eq!(layout.default_row_height, None);
        assert!(layout.is_empty());
    }

    /// Carried rows come out sorted by row number whatever the record
    /// order was, and empty rows keep the whole layout non-empty so the
    /// conversion still emits them.
    #[test]
    fn carried_rows_are_sorted_and_keep_the_layout_interesting() {
        let mut acc = SheetAccumulator::default();
        let mut body = 7u16.to_le_bytes().to_vec();
        body.extend_from_slice(&[0u8; 4]);
        body.extend_from_slice(&600u16.to_le_bytes());
        body.extend_from_slice(&[0u8; 4]);
        body.extend_from_slice(&0x0140u16.to_le_bytes());
        body.extend_from_slice(&[0u8; 2]);
        acc.push_row(&body);
        let mut body = 1u16.to_le_bytes().to_vec();
        body.extend_from_slice(&[0u8; 4]);
        body.extend_from_slice(&300u16.to_le_bytes());
        body.extend_from_slice(&[0u8; 4]);
        body.extend_from_slice(&0x0120u16.to_le_bytes()); // hidden
        body.extend_from_slice(&[0u8; 2]);
        acc.push_row(&body);
        let layout = acc.finish(&StyleTables::default(), 0);
        assert_eq!(
            layout.rows,
            vec![
                (1, RowSpec {
                    height: None,
                    hidden: true
                }),
                (7, RowSpec {
                    height: Some(30.0),
                    hidden: false
                }),
            ]
        );
        assert_eq!(layout.row_spec(7).map(|spec| spec.height), Some(Some(30.0)));
        assert!(layout.row_spec(2).is_none());
        assert!(!layout.is_empty());
    }
}

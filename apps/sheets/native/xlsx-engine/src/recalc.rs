//! IronCalc-backed formula recalculation with a session-resident model cache.
//!
//! The first request for a workbook loads it into IronCalc's model; later
//! requests reuse that model and only apply the edits that changed, so a
//! debounced keystroke pays evaluate() instead of a full file import. The
//! cache invalidates itself when the file on disk changes (mtime+size) and
//! rebuilds when the edit set shrinks (an undo must restore file content the
//! model no longer has). The caller stays fail-soft: any load or evaluation
//! problem (including panics from IronCalc's strict importer) surfaces as a
//! normal sidecar error and the renderer keeps showing cached values.

use std::collections::HashMap;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::{Path, PathBuf};
use std::time::SystemTime;

use ironcalc::base::Model;
use ironcalc::base::formatter::format::format_number;
use ironcalc::base::locale::{get_default_locale, get_locale};
use ironcalc::import::load_from_xlsx;
use serde::{Deserialize, Serialize};

use crate::{CellRange, SidecarError};

pub const MAX_RECALC_EDITS: usize = 10_000;
pub const MAX_RECALC_READ_CELLS: usize = 20_000;
/// Resident models are large (the whole workbook's cell graph); the sidecar
/// serves one document, so two covers the active file plus one recently
/// closed-and-reopened neighbour.
const MAX_RESIDENT_MODELS: usize = 2;
/// Source files above this size (compressed bytes) count as heavy for the
/// residency rule in `evict_beyond_cap`.
const HEAVY_SOURCE_BYTES: u64 = 8_000_000;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RecalcEdit {
    pub sheet: String,
    /// 0-based coordinates on the wire (IronCalc itself is 1-based).
    pub row: u32,
    pub column: u32,
    /// User input: `=SUM(A1:A3)`, `42`, `text`; empty clears.
    pub input: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RecalcRead {
    pub sheet: String,
    pub range: CellRange,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RecalcCell {
    pub sheet: String,
    pub row: u32,
    pub column: u32,
    /// Display string after evaluation (number formats applied).
    pub formatted: String,
    /// Raw numeric value when the cell evaluates to a number.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number: Option<f64>,
    pub is_formula: bool,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RecalcResult {
    pub cells: Vec<RecalcCell>,
    /// True when a resident model served this request (no file re-import).
    pub cached: bool,
}

type EditKey = (String, u32, u32);

struct ResidentModel {
    model: Model<'static>,
    mtime: SystemTime,
    size: u64,
    /// Edits already in the model, keyed by cell; a request whose edit set no
    /// longer covers these keys forces a rebuild (only the file knows the
    /// original content of a reverted cell).
    applied: HashMap<EditKey, String>,
    last_used: u64,
}

#[derive(Default)]
pub struct RecalcCache {
    entries: HashMap<PathBuf, ResidentModel>,
    tick: u64,
}

impl RecalcCache {
    pub fn new() -> Self {
        Self::default()
    }

    /// Drop the model for a path whose bytes are about to change (save) or
    /// whose session closed.
    pub fn purge(&mut self, path: &Path) {
        self.entries.remove(&cache_key(path));
    }

    fn evict_beyond_cap(&mut self) {
        while self.entries.len() > MAX_RESIDENT_MODELS {
            let Some(oldest) = self
                .entries
                .iter()
                .min_by_key(|(_, entry)| entry.last_used)
                .map(|(path, _)| path.clone())
            else {
                return;
            };
            self.entries.remove(&oldest);
        }
        // Heavy sources import into models that dwarf everything else in the
        // process (a 31MB workbook's model holds ~1.2GB resident); keep at
        // most one of those — the most recently used.
        loop {
            let mut heavy: Vec<(PathBuf, u64)> = self
                .entries
                .iter()
                .filter(|(_, entry)| entry.size > HEAVY_SOURCE_BYTES)
                .map(|(path, entry)| (path.clone(), entry.last_used))
                .collect();
            if heavy.len() <= 1 {
                return;
            }
            heavy.sort_by_key(|(_, last_used)| *last_used);
            let Some((oldest, _)) = heavy.first() else {
                return;
            };
            self.entries.remove(&oldest.clone());
        }
    }
}

/// Cache keys are canonicalized: sessions store the canonical workbook path
/// (open() resolves it) while recalc requests carry the renderer's raw path,
/// and on macOS temp files those differ (/var vs /private/var) — a raw key
/// would make the close/save purge miss the resident model.
fn cache_key(path: &Path) -> PathBuf {
    path.canonicalize().unwrap_or_else(|_| path.to_path_buf())
}

pub fn recalc_cells(
    cache: &mut RecalcCache,
    path: &Path,
    edits: &[RecalcEdit],
    reads: &[RecalcRead],
) -> Result<RecalcResult, SidecarError> {
    if edits.len() > MAX_RECALC_EDITS {
        return Err(SidecarError::InvalidRequest(format!(
            "Too many recalc edits (limit {MAX_RECALC_EDITS})."
        )));
    }
    let read_cells: usize = reads
        .iter()
        .map(|read| {
            let rows = read.range.end_row.saturating_sub(read.range.start_row) + 1;
            let columns = read
                .range
                .end_column
                .saturating_sub(read.range.start_column)
                + 1;
            rows.saturating_mul(columns)
        })
        .sum();
    if read_cells > MAX_RECALC_READ_CELLS {
        return Err(SidecarError::InvalidRequest(format!(
            "Recalc read exceeds {MAX_RECALC_READ_CELLS} cells."
        )));
    }

    let metadata = std::fs::metadata(path).map_err(|error| SidecarError::Io(error.to_string()))?;
    let mtime = metadata
        .modified()
        .map_err(|error| SidecarError::Io(error.to_string()))?;
    let size = metadata.len();

    let key = cache_key(path);
    // Take the entry out while working on it: a panic or error mid-apply
    // leaves the model in an unknown state, and a removed entry can never be
    // reused by the next request.
    let resident = cache.entries.remove(&key).filter(|entry| {
        entry.mtime == mtime
            && entry.size == size
            && entry.applied.keys().all(|key| {
                edits
                    .iter()
                    .any(|edit| edit.sheet == key.0 && edit.row == key.1 && edit.column == key.2)
            })
    });
    let cached = resident.is_some();

    // IronCalc's importer is strict and can panic on non-Excel producers
    // (missing cellStyles, whitespace nodes); contain that to this request.
    let (entry, cells) = catch_unwind(AssertUnwindSafe(|| {
        run(resident, path, mtime, size, edits, reads)
    }))
    .map_err(|_| {
        SidecarError::Workbook("The formula engine could not process this workbook.".into())
    })??;

    cache.tick += 1;
    let mut entry = entry;
    entry.last_used = cache.tick;
    cache.entries.insert(key, entry);
    cache.evict_beyond_cap();

    Ok(RecalcResult { cells, cached })
}

fn run(
    resident: Option<ResidentModel>,
    path: &Path,
    mtime: SystemTime,
    size: u64,
    edits: &[RecalcEdit],
    reads: &[RecalcRead],
) -> Result<(ResidentModel, Vec<RecalcCell>), SidecarError> {
    let mut entry = match resident {
        Some(entry) => entry,
        None => {
            let path_text = path.to_str().ok_or_else(|| {
                SidecarError::InvalidRequest("Workbook path is not valid UTF-8.".into())
            })?;
            let mut model = load_from_xlsx(path_text, "en", "UTC", "en").map_err(|error| {
                SidecarError::Workbook(format!("Formula engine import failed: {error}"))
            })?;
            pin_unparsable_formulas(&mut model);
            ResidentModel {
                model,
                mtime,
                size,
                applied: HashMap::new(),
                last_used: 0,
            }
        }
    };

    let mut dirty = false;
    for edit in edits {
        let key = (edit.sheet.clone(), edit.row, edit.column);
        if entry.applied.get(&key) == Some(&edit.input) {
            continue;
        }
        let sheet = sheet_index(&entry.model, &edit.sheet)?;
        entry
            .model
            .set_user_input(
                sheet,
                edit.row as i32 + 1,
                edit.column as i32 + 1,
                edit.input.clone(),
            )
            .map_err(|error| {
                SidecarError::Workbook(format!("Formula engine rejected an edit: {error}"))
            })?;
        entry.applied.insert(key, edit.input.clone());
        dirty = true;
    }
    if dirty || entry.last_used == 0 {
        entry.model.evaluate();
    }

    let mut cells = Vec::new();
    for read in reads {
        let sheet = sheet_index(&entry.model, &read.sheet)?;
        for row in read.range.start_row..=read.range.end_row {
            for column in read.range.start_column..=read.range.end_column {
                let row_1 = row as i32 + 1;
                let column_1 = column as i32 + 1;
                let formatted = entry
                    .model
                    .get_formatted_cell_value(sheet, row_1, column_1)
                    .unwrap_or_default();
                let formula = entry
                    .model
                    .get_cell_formula(sheet, row_1, column_1)
                    .ok()
                    .flatten();
                let is_formula = formula.is_some();
                if formatted.is_empty() && !is_formula {
                    continue;
                }
                // Known engine gap where the file's cached value beats the
                // error: RATE's Newton solver used to die near -100% where
                // Excel converges (#185; the renderer computes RATE itself,
                // the skip keeps a cached value alive if the engine still
                // errors). CELL("filename") needed the same treatment until
                // IronCalc implemented it in 0.8 (#175); it now computes, so
                // only RATE is covered.
                if formatted.starts_with('#') {
                    if let Some(text) = &formula {
                        let upper = text.to_uppercase();
                        if formatted == "#NUM!" && upper.contains("RATE(") {
                            continue;
                        }
                    }
                }
                let number = raw_number(&entry.model, sheet, row_1, column_1);
                // BUG-1507: a date-returning function under General shows a
                // date, like the format Excel applies at entry time.
                let formatted = match (number, formula.as_deref()) {
                    (Some(value), Some(text)) => {
                        auto_date_display(&entry.model, sheet, row_1, column_1, value, text)
                            .unwrap_or(formatted)
                    }
                    _ => formatted,
                };
                cells.push(RecalcCell {
                    sheet: read.sheet.clone(),
                    row: row as u32,
                    column: column as u32,
                    formatted,
                    number,
                    is_formula,
                });
            }
        }
    }
    Ok((entry, cells))
}

enum PinnedValue {
    Number(f64),
    Text(String),
    Bool(bool),
}

/// IronCalc cannot parse external-workbook references (`[1]Sheet1!A1`):
/// such a formula evaluates to #ERROR! and the error cascades through every
/// dependent. Excel keeps the cached values when the linked workbook is
/// unreachable, so pin those cells to the value the file carries and let the
/// dependents compute against it. The renderer never sees the cell as a
/// formula afterwards, so its own cached copy stays on screen.
fn pin_unparsable_formulas(model: &mut Model) {
    use ironcalc::base::expressions::parser::Node;
    use ironcalc::base::types::{Cell, FormulaValue};
    let mut pins = Vec::new();
    for (sheet, worksheet) in model.workbook.worksheets.iter().enumerate() {
        let Some(parsed) = model.parsed_formulas.get(sheet) else {
            continue;
        };
        for (row, columns) in &worksheet.sheet_data {
            for (column, cell) in columns {
                let (formula, value) = match cell {
                    Cell::CellFormula {
                        f,
                        v: FormulaValue::Number(value),
                        ..
                    } => (*f, PinnedValue::Number(*value)),
                    Cell::CellFormula {
                        f,
                        v: FormulaValue::Text(value),
                        ..
                    } => (*f, PinnedValue::Text(value.clone())),
                    Cell::CellFormula {
                        f,
                        v: FormulaValue::Boolean(value),
                        ..
                    } => (*f, PinnedValue::Bool(*value)),
                    _ => continue,
                };
                if matches!(
                    parsed.get(formula as usize),
                    Some((Node::ParseErrorKind { .. }, _))
                ) {
                    pins.push((sheet as u32, *row, *column, value));
                }
            }
        }
    }
    for (sheet, row, column, value) in pins {
        // A pin that fails leaves the cell as it was: #ERROR! on that cell,
        // which the renderer already declines to display.
        let _ = match value {
            PinnedValue::Number(number) => {
                model.update_cell_with_number(sheet, row, column, number)
            }
            PinnedValue::Text(text) => model.update_cell_with_text(sheet, row, column, &text),
            PinnedValue::Bool(flag) => model.update_cell_with_bool(sheet, row, column, flag),
        };
    }
}

fn sheet_index(model: &Model, name: &str) -> Result<u32, SidecarError> {
    model
        .workbook
        .worksheets
        .iter()
        .position(|worksheet| worksheet.name == name)
        .map(|index| index as u32)
        .ok_or_else(|| SidecarError::InvalidRequest(format!("Unknown sheet: {name}")))
}

/// The root function name of a formula that is a single call, uppercased
/// ("DATE" for "=DATE(2023,12,31)"); None for literals, references and
/// anything else — an arithmetic tail ("=DATE(..)+0") or a wrapper
/// ("=SUM(DATE(..),0)") is a numeric expression, not a date entry. The `=`
/// is always there (get_cell_formula adds it).
fn root_function(formula: &str) -> Option<String> {
    let body = formula.strip_prefix('=')?.trim();
    let name_end = body
        .find(|c: char| !(c.is_ascii_alphanumeric() || c == '.' || c == '_'))
        .unwrap_or(body.len());
    let (name, rest) = body.split_at(name_end);
    if name.is_empty() || !rest.trim_start().starts_with('(') {
        return None;
    }
    if !call_closes_at_end(rest.trim_start()) {
        return None;
    }
    Some(name.to_ascii_uppercase())
}

/// Whether the first `(` closes as the last character of the call: the call
/// is the whole expression, no operator or extra token follows it. String
/// literals are skipped so a `)` inside them cannot fake the close.
fn call_closes_at_end(call: &str) -> bool {
    let mut depth = 0usize;
    let mut in_string = false;
    let mut close = None;
    for (index, ch) in call.char_indices() {
        if in_string {
            if ch == '"' {
                in_string = false;
            }
        } else {
            match ch {
                '"' => in_string = true,
                '(' => depth += 1,
                ')' => {
                    if depth == 0 {
                        return false;
                    }
                    depth -= 1;
                    if depth == 0 && close.is_none() {
                        close = Some(index);
                    }
                }
                _ => {}
            }
        }
    }
    !in_string && depth == 0 && close.is_some_and(|index| call[index + 1..].trim().is_empty())
}

/// Excel applies an automatic date format when a date-returning function is
/// entered into a General cell: =DATE(2023,12,31) shows 12/31/23, not the
/// raw serial 45291. Files saved without cached values (openpyxl and other
/// third-party generators) surface through this recalc path, so the display
/// mirrors that automatic format: the locale's short date for the date
/// functions, date and time for NOW, a time for TIME and TIMEVALUE. Only a
/// single root call counts (numeric expressions stay numeric), and an
/// explicit cell format is never overridden — Excel applies the automatic
/// format only at entry, and the file's own formats always win.
fn auto_date_display(
    model: &Model,
    sheet: u32,
    row: i32,
    column: i32,
    value: f64,
    formula: &str,
) -> Option<String> {
    if !value.is_finite() || value < 0.0 {
        return None;
    }
    let is_general = model
        .get_style_for_cell(sheet, row, column)
        .map(|style| style.num_fmt.eq_ignore_ascii_case("general"))
        .unwrap_or(false);
    if !is_general {
        return None;
    }
    let locale = get_locale(&model.get_locale()).unwrap_or_else(|_| get_default_locale());
    let short_date = &locale.dates.date_formats.short;
    let pattern = match root_function(formula)?.as_str() {
        "DATE" | "DATEVALUE" | "TODAY" | "EDATE" | "EOMONTH" | "WORKDAY" | "WORKDAY.INTL" => {
            short_date.clone()
        }
        "NOW" => format!("{short_date} h:mm"),
        "TIME" | "TIMEVALUE" => "h:mm AM/PM".to_string(),
        _ => return None,
    };
    Some(format_number(value, &pattern, locale).text)
}

fn raw_number(model: &Model, sheet: u32, row: i32, column: i32) -> Option<f64> {
    use ironcalc::base::cell::CellValue;
    match model.get_cell_value_by_index(sheet, row, column) {
        Ok(CellValue::Number(value)) => Some(value),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use ironcalc::base::types::Style;
    use ironcalc::export::save_to_xlsx;

    fn fixture(path: &Path) {
        write_fixture(path, &[("A1", "10"), ("A2", "20"), ("A3", "=SUM(A1:A2)")]);
    }

    fn write_fixture(path: &Path, cells: &[(&str, &str)]) {
        let mut model = Model::new_empty("fixture", "en", "UTC", "en").unwrap();
        for (address, input) in cells {
            let column = (address.as_bytes()[0] - b'A' + 1) as i32;
            let row: i32 = address[1..].parse().unwrap();
            model
                .set_user_input(0, row, column, (*input).to_string())
                .unwrap();
        }
        model.evaluate();
        if path.exists() {
            std::fs::remove_file(path).unwrap();
        }
        save_to_xlsx(&model, path.to_str().unwrap()).unwrap();
    }

    fn edit(address: &str, input: &str) -> RecalcEdit {
        RecalcEdit {
            sheet: "Sheet1".into(),
            row: address[1..].parse::<u32>().unwrap() - 1,
            column: (address.as_bytes()[0] - b'A') as u32,
            input: input.into(),
        }
    }

    fn read_a1_a3() -> RecalcRead {
        RecalcRead {
            sheet: "Sheet1".into(),
            range: CellRange {
                start_row: 0,
                end_row: 2,
                start_column: 0,
                end_column: 0,
            },
        }
    }

    fn sum_value(result: &RecalcResult) -> String {
        result
            .cells
            .iter()
            .find(|cell| cell.row == 2)
            .unwrap()
            .formatted
            .clone()
    }

    /// BUG-1507: a date-returning function in a General cell gets Excel's
    /// automatic date format on display (DATE/TODAY/EDATE/EOMONTH/WORKDAY as
    /// the locale short date, NOW with a time, TIME/TIMEVALUE as a time);
    /// numeric expressions and literal serials keep the General number
    /// display, and an explicit number format is never overridden. The
    /// fixture models a third-party generator (openpyxl & co): those write
    /// the formula with the cell left General, so the entry-time format
    /// IronCalc applies inside set_user_input is stripped again.
    #[test]
    fn date_functions_display_as_dates_under_general() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("dates.xlsx");
        let mut model = Model::new_empty("fixture", "en", "UTC", "en").unwrap();
        let cells: [(&str, &str); 8] = [
            ("A1", "=DATE(2023,12,31)"),
            ("A2", "=EDATE(DATE(2023,12,31),2)"),
            ("A3", "=EOMONTH(DATE(2023,12,1),0)"),
            ("A4", "=DATE(2023,12,31)+0"),
            ("A5", "45291"),
            ("A6", "=DATE(2023,12,31)"),
            ("A7", "=TIME(12,0,0)"),
            ("A8", "=SUM(DATE(2023,12,31),0)"),
        ];
        for (address, input) in cells {
            let column = (address.as_bytes()[0] - b'A' + 1) as i32;
            let row: i32 = address[1..].parse().unwrap();
            model
                .set_user_input(0, row, column, input.to_string())
                .unwrap();
        }
        for row in [1, 2, 3, 4, 7, 8] {
            model.set_cell_style(0, row, 1, &Style::default()).unwrap();
        }
        // A6 carries an explicit numeric format: the raw serial rounds, the
        // automatic date format must not touch it.
        let mut explicit = Style::default();
        explicit.num_fmt = "0.00".to_string();
        model.set_cell_style(0, 6, 1, &explicit).unwrap();
        model.evaluate();
        save_to_xlsx(&model, path.to_str().unwrap()).unwrap();

        let mut cache = RecalcCache::new();
        let result = recalc_cells(
            &mut cache,
            &path,
            &[],
            &[RecalcRead {
                sheet: "Sheet1".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 7,
                    start_column: 0,
                    end_column: 0,
                },
            }],
        )
        .unwrap();
        let at = |row: u32| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row)
                .map(|cell| (cell.formatted.as_str(), cell.number))
                .unwrap()
        };
        // Date functions: the short-date/time rendering Excel auto-applies.
        assert_eq!(at(0), ("12/31/23", Some(45291.0)));
        assert_eq!(at(1), ("2/29/24", Some(45351.0)));
        assert_eq!(at(2), ("12/31/23", Some(45291.0)));
        assert_eq!(at(6), ("12:00 PM", Some(0.5)));
        // Numeric expressions and plain serials keep the General display.
        assert_eq!(at(3), ("45291", Some(45291.0)));
        assert_eq!(at(4), ("45291", Some(45291.0)));
        assert_eq!(at(7), ("45291", Some(45291.0)));
        // An explicit cell format always wins over the automatic date.
        assert_eq!(at(5), ("45291.00", Some(45291.0)));
    }

    /// The automatic format only applies to a formula that is a single date
    /// call: wrappers, arithmetic tails and literals stay numeric.
    #[test]
    fn root_function_matches_only_a_single_root_call() {
        assert_eq!(root_function("=DATE(2023,12,31)").as_deref(), Some("DATE"));
        assert_eq!(root_function("=today()  ").as_deref(), Some("TODAY"));
        assert_eq!(
            root_function("=WORKDAY.INTL(DATE(2023,12,31),1)").as_deref(),
            Some("WORKDAY.INTL")
        );
        // A ")" inside a string literal cannot fake the close; the real one
        // still ends the call.
        assert_eq!(
            root_function(r#"=DATEVALUE("12/31/2023 (file)")"#).as_deref(),
            Some("DATEVALUE")
        );
        assert_eq!(root_function("=DATE(2023,12,31)+0"), None);
        assert_eq!(root_function("=DATE(2023,12,31)*TODAY()"), None);
        // A wrapper is a single root call, but not a date one: the
        // auto-date allowlist in auto_date_display skips it.
        assert_eq!(
            root_function("=SUM(DATE(2023,12,31),0)").as_deref(),
            Some("SUM")
        );
        assert_eq!(root_function("=IF(TRUE,TODAY(),0)").as_deref(), Some("IF"));
        assert_eq!(root_function("45291"), None);
        assert_eq!(root_function("=A1"), None);
    }

    /// TIME's time pattern renders with the AM/PM marker Excel shows on
    /// entry (kept deterministic by formatting the exact serial 0.5: the
    /// formatter floors near-inexact second fractions one minute down).
    #[test]
    fn time_pattern_renders_am_pm() {
        let locale = get_locale("en").unwrap();
        let text = format_number(0.5, "h:mm AM/PM", locale).text;
        assert_eq!(text, "12:00 PM");
    }

    /// NOW()'s combined date+time pattern renders through the same formatter
    /// (kept deterministic by formatting a fixed serial directly).
    #[test]
    fn now_pattern_renders_date_and_time() {
        let locale = get_locale("en").unwrap();
        let text = format_number(45291.5, "m/d/yy h:mm", locale).text;
        assert_eq!(text, "12/31/23 12:00");
    }

    /// Deleting rows/columns that formulas reference now succeeds and writes
    /// Excel-style #REF! into the affected formulas (the gateway rewrites
    /// orphaned references to the bare #REF! token — the qualified
    /// `Sheet!#REF!` form is not parseable here). The sidecar must round-trip
    /// those formulas: parse them, evaluate to the #REF! error, and keep
    /// reporting them as formulas on recalc.
    #[test]
    fn ref_error_formulas_survive_the_recalc_round_trip() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("ref-error.xlsx");
        write_fixture(
            &path,
            &[
                ("A1", "10"),
                ("A2", "20"),
                ("A3", "=SUM(#REF!)"),
                ("A4", "=#REF!+1"),
                ("A5", "=SUM(#REF!)+Sheet1!A1"),
            ],
        );
        let mut cache = RecalcCache::new();
        let result = recalc_cells(
            &mut cache,
            &path,
            &[],
            &[RecalcRead {
                sheet: "Sheet1".into(),
                range: CellRange {
                    start_row: 2,
                    end_row: 4,
                    start_column: 0,
                    end_column: 0,
                },
            }],
        )
        .unwrap();
        let by_row = |row: u32| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row)
                .map(|cell| (cell.formatted.clone(), cell.is_formula))
                .unwrap()
        };
        assert_eq!(by_row(2), ("#REF!".to_owned(), true));
        assert_eq!(by_row(3), ("#REF!".to_owned(), true));
        assert_eq!(by_row(4), ("#REF!".to_owned(), true));
    }

    #[test]
    fn keeps_at_most_one_heavy_model_resident() {
        let mut cache = RecalcCache::new();
        for (name, size) in [
            ("heavy-old", HEAVY_SOURCE_BYTES + 1),
            ("heavy-new", HEAVY_SOURCE_BYTES + 2),
            ("light", 1),
        ] {
            cache.tick += 1;
            cache.entries.insert(
                PathBuf::from(name),
                ResidentModel {
                    model: Model::new_empty("fixture", "en", "UTC", "en").unwrap(),
                    mtime: SystemTime::now(),
                    size,
                    applied: HashMap::new(),
                    last_used: cache.tick,
                },
            );
            cache.evict_beyond_cap();
        }
        assert!(!cache.entries.contains_key(Path::new("heavy-old")));
        assert!(cache.entries.contains_key(Path::new("heavy-new")));
        assert!(cache.entries.contains_key(Path::new("light")));
    }

    #[test]
    fn recalculates_after_edits() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        let result =
            recalc_cells(&mut cache, &path, &[edit("A1", "100")], &[read_a1_a3()]).unwrap();
        let sum = result.cells.iter().find(|cell| cell.row == 2).unwrap();
        assert_eq!(sum.formatted, "120");
        assert_eq!(sum.number, Some(120.0));
        assert!(sum.is_formula);
        assert!(!result.cached);
    }

    #[test]
    fn reuses_the_resident_model_and_applies_edits_incrementally() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        let first = recalc_cells(&mut cache, &path, &[edit("A1", "100")], &[read_a1_a3()]).unwrap();
        assert!(!first.cached);
        assert_eq!(sum_value(&first), "120");
        // same edit again: reused model, nothing re-applied
        let second =
            recalc_cells(&mut cache, &path, &[edit("A1", "100")], &[read_a1_a3()]).unwrap();
        assert!(second.cached);
        assert_eq!(sum_value(&second), "120");
        // changed edit: reused model, incremental set_user_input
        let third = recalc_cells(&mut cache, &path, &[edit("A1", "200")], &[read_a1_a3()]).unwrap();
        assert!(third.cached);
        assert_eq!(sum_value(&third), "220");
    }

    #[test]
    fn rebuilds_when_an_edit_is_reverted() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        let first = recalc_cells(&mut cache, &path, &[edit("A1", "100")], &[read_a1_a3()]).unwrap();
        assert_eq!(sum_value(&first), "120");
        // undo removed the A1 edit: only the file knows A1's original value
        let second = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert!(!second.cached);
        assert_eq!(sum_value(&second), "30");
    }

    #[test]
    fn rebuilds_when_the_file_changes_on_disk() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        let first = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert!(!first.cached);
        assert_eq!(sum_value(&first), "30");
        write_fixture(
            &path,
            &[("A1", "11"), ("A2", "20"), ("A3", "=SUM(A1:A2)+1000")],
        );
        let second = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert!(!second.cached);
        assert_eq!(sum_value(&second), "1031");
    }

    #[test]
    fn purge_drops_the_resident_model() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        cache.purge(&path);
        let result = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert!(!result.cached);
    }

    #[test]
    fn an_errored_request_does_not_poison_the_cache() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        let error = recalc_cells(
            &mut cache,
            &path,
            &[],
            &[RecalcRead {
                sheet: "Nope".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 0,
                    start_column: 0,
                    end_column: 0,
                },
            }],
        )
        .unwrap_err();
        assert!(matches!(error, SidecarError::InvalidRequest(_)));
        // the entry was taken out and not re-inserted; next call reloads cleanly
        let after = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert!(!after.cached);
        assert_eq!(sum_value(&after), "30");
    }

    #[test]
    fn caps_resident_models() {
        let dir = tempfile::tempdir().unwrap();
        let mut cache = RecalcCache::new();
        let mut paths = Vec::new();
        for index in 0..3 {
            let path = dir.path().join(format!("recalc-{index}.xlsx"));
            fixture(&path);
            recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
            paths.push(path);
        }
        assert_eq!(cache.entries.len(), MAX_RESIDENT_MODELS);
        // the oldest was evicted, the newest survives (entries key by
        // canonical path)
        assert!(!cache.entries.contains_key(&cache_key(&paths[0])));
        assert!(cache.entries.contains_key(&cache_key(&paths[2])));
    }

    #[test]
    fn purge_hits_regardless_of_path_spelling() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert_eq!(cache.entries.len(), 1);
        // The raw and canonical spellings differ on macOS temp dirs
        // (/var vs /private/var); the close/save purge may hold either.
        cache.purge(&path);
        assert!(cache.entries.is_empty());
        recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        cache.purge(&cache_key(&path));
        assert!(cache.entries.is_empty());
    }

    #[test]
    fn rejects_unknown_sheets() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let error = recalc_cells(
            &mut RecalcCache::new(),
            &path,
            &[],
            &[RecalcRead {
                sheet: "Nope".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 0,
                    start_column: 0,
                    end_column: 0,
                },
            }],
        )
        .unwrap_err();
        assert!(matches!(error, SidecarError::InvalidRequest(_)));
    }

    #[test]
    fn caps_read_volume() {
        let error = recalc_cells(
            &mut RecalcCache::new(),
            Path::new("/nonexistent.xlsx"),
            &[],
            &[RecalcRead {
                sheet: "Sheet1".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 999,
                    start_column: 0,
                    end_column: 999,
                },
            }],
        )
        .unwrap_err();
        assert!(matches!(error, SidecarError::InvalidRequest(_)));
    }

    fn write_external_link_fixture(path: &Path) {
        use std::io::Write;
        let entries: [(&str, &str); 8] = [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><externalReferences><externalReference r:id="rId3"/></externalReferences></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>"#,
            ),
            (
                "xl/styles.xml",
                concat!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
                    crate::xls_layout::default_cell_styles_xml!(),
                    "</styleSheet>",
                ),
            ),
            (
                "xl/externalLinks/externalLink1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>42</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#,
            ),
            (
                "xl/externalLinks/_rels/externalLink1.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///C:/data/source.xlsx" TargetMode="External"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1"><f>[1]Sheet1!A1*2</f><v>84</v></c><c r="B1"><f>A1+1</f><v>85</v></c><c r="C1" t="str"><f>'[1]Sheet1'!A1&amp;" units"</f><v>42 units</v></c></row><row r="2"><c r="A2"><v>5</v></c><c r="B2"><f>SUM(A1:A2)</f><v>89</v></c></row></sheetData></worksheet>"#,
            ),
        ];
        let mut writer = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, content) in entries {
            writer.start_file(name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
    }

    /// External-workbook references never parse in IronCalc; the file's
    /// cached values must stand in for them so dependents keep computing
    /// instead of cascading #ERROR! (public issue 235).
    #[test]
    fn external_link_formulas_keep_their_cached_values() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("external.xlsx");
        write_external_link_fixture(&path);
        let mut cache = RecalcCache::new();
        let result = recalc_cells(
            &mut cache,
            &path,
            &[edit("A2", "10")],
            &[RecalcRead {
                sheet: "Sheet1".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 1,
                    start_column: 0,
                    end_column: 2,
                },
            }],
        )
        .unwrap();
        let at = |row: u32, column: u32| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row && cell.column == column)
                .unwrap()
        };
        assert!(result.cells.iter().all(|cell| cell.formatted != "#ERROR!"));
        // pinned to the cache and no longer reported as formulas
        assert_eq!(
            (at(0, 0).formatted.as_str(), at(0, 0).is_formula),
            ("84", false)
        );
        assert_eq!(
            (at(0, 2).formatted.as_str(), at(0, 2).is_formula),
            ("42 units", false)
        );
        // dependents compute against the pinned values
        assert_eq!(
            (at(0, 1).formatted.as_str(), at(0, 1).is_formula),
            ("85", true)
        );
        assert_eq!(
            (at(1, 1).formatted.as_str(), at(1, 1).is_formula),
            ("94", true)
        );
    }

    /// BUG-1659 end to end: a legacy-converted workbook reaches the formula
    /// engine and its formulas compute. The converter used to emit styles.xml
    /// without the mandatory `<cellStyles>` section, IronCalc's importer
    /// panicked on the empty section, and every imported .ods/.xls book
    /// answered workbook_error — formulas silently empty in the UI.
    #[test]
    fn a_converted_legacy_book_computes_formulas_in_the_engine() {
        let source = std::path::Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/bug-1659-legacy-formulas.xls"
        ));
        let dir = tempfile::tempdir().unwrap();
        let target = dir.path().join("converted.xlsx");
        crate::convert::convert_to_xlsx(source, &target).unwrap();

        let mut cache = RecalcCache::new();
        // B2 is =A1*B1 (10*4): an edit to A1 must recompute it live.
        let result = recalc_cells(
            &mut cache,
            &target,
            &[edit("A1", "100")],
            &[RecalcRead {
                sheet: "Sheet1".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 1,
                    start_column: 0,
                    end_column: 1,
                },
            }],
        )
        .unwrap();
        let at = |row: u32, column: u32| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row && cell.column == column)
                .unwrap()
        };
        assert_eq!(at(0, 0).formatted.as_str(), "100");
        assert_eq!(
            (at(1, 1).formatted.as_str(), at(1, 1).is_formula),
            ("400", true)
        );
        assert_eq!(at(1, 1).number, Some(400.0));
    }

    /// BUG-1659: the .ods conversion must reach the engine too — recalc used
    /// to answer workbook_error there on the same missing-`<cellStyles>`
    /// panic. The verbatim ODF `of:=` formula syntax stays a known separate
    /// gap, so this pins loading and the literal cells, not formula results.
    #[test]
    fn a_converted_ods_book_loads_in_the_formula_engine() {
        let source = std::path::Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/bug-1659-legacy-formulas.ods"
        ));
        let dir = tempfile::tempdir().unwrap();
        let target = dir.path().join("converted.xlsx");
        crate::convert::convert_to_xlsx(source, &target).unwrap();

        let mut cache = RecalcCache::new();
        let result = recalc_cells(
            &mut cache,
            &target,
            &[],
            &[RecalcRead {
                sheet: "Sheet1".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 1,
                    start_column: 0,
                    end_column: 1,
                },
            }],
        )
        .unwrap();
        let at = |row: u32, column: u32| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row && cell.column == column)
                .map(|cell| cell.formatted.as_str().to_owned())
        };
        assert_eq!(at(0, 0).as_deref(), Some("10"));
        assert_eq!(at(0, 1).as_deref(), Some("4"));
        assert_eq!(at(1, 0).as_deref(), Some("20"));
    }

    /// BUG-1660 reference suite against Excel 365 values. IronCalc 0.7.1 had
    /// no SUMPRODUCT and no LET (both evaluated to #NAME?) and comparisons/IF
    /// did not lift element-wise over ranges, so SUM(IF(A1:A3>10,1,0)) summed
    /// a single implicit-intersection value (1 instead of Excel's 2). The
    /// upgraded engine computes every case with Excel semantics: non-numeric
    /// entries count as 0, dimension mismatches yield #VALUE!, and an unknown
    /// function still yields #NAME? — which is Excel's own behavior for
    /// unknown names, unlike the audit cases, which are valid Excel formulas.
    #[test]
    fn sumproduct_let_and_array_if_match_excel_semantics() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("bug-1660.xlsx");
        write_fixture(
            &path,
            &[
                ("A1", "5"),
                ("A2", "15"),
                ("A3", "20"),
                ("B1", "text"),
                ("B2", "2"),
                ("B3", "4"),
                ("D1", "=SUMPRODUCT(A1:A3)"),
                ("D2", "=SUMPRODUCT(A1:A3,A1:A3)"),
                ("D3", "=SUMPRODUCT((A1:A3>10)*1)"),
                ("D4", "=SUMPRODUCT((A1:A3>10)*A1:A3)"),
                ("D5", "=SUMPRODUCT(A1:A3*2)"),
                ("D6", "=SUMPRODUCT(A1:A3,B1:B3)"),
                ("D7", "=SUMPRODUCT((A1:A3>10)*(A1:A3<18))"),
                ("D8", "=SUMPRODUCT((A1:A3>10)*1,B1:B3)"),
                ("D9", "=SUMPRODUCT(A1:A2,B1:B3)"),
                ("D10", "=LET(x,5,x*2)"),
                ("D11", "=LET(s,SUM(A1:A3),s+10)"),
                ("D12", "=LET(r,A1:A3,SUMPRODUCT((r>10)*1))"),
                ("D13", "=SUM(IF(A1:A3>10,1,0))"),
                ("D14", "=THIS_IS_NOT_A_FUNCTION(1)"),
            ],
        );
        let mut cache = RecalcCache::new();
        let result = recalc_cells(
            &mut cache,
            &path,
            &[],
            &[RecalcRead {
                sheet: "Sheet1".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 13,
                    start_column: 3,
                    end_column: 3,
                },
            }],
        )
        .unwrap();
        let at = |row: u32| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row && cell.column == 3)
                .map(|cell| cell.formatted.clone())
                .unwrap_or_else(|| String::from("<missing>"))
        };
        // SUMPRODUCT with ranges and conditions (Excel 365 values).
        assert_eq!(at(0), "40"); // 5+15+20
        assert_eq!(at(1), "650"); // 25+225+400
        assert_eq!(at(2), "2"); // the audit repro: two cells above 10
        assert_eq!(at(3), "35"); // 15+20
        assert_eq!(at(4), "80"); // (5+15+20)*2
        assert_eq!(at(5), "110"); // B1 text counts as 0: 15*2+20*4
        assert_eq!(at(6), "1"); // only 15 satisfies both conditions
        assert_eq!(at(7), "6"); // (0,1,1)·(0,2,4)
        assert_eq!(at(8), "#VALUE!"); // 2x1 vs 3x1 dimension mismatch
        // LET (the audit repro returned #NAME? for both spellings).
        assert_eq!(at(9), "10");
        assert_eq!(at(10), "50"); // s=40, s+10
        assert_eq!(at(11), "2"); // a bound range feeds SUMPRODUCT
        // Array-context IF: Excel 365 computes 2, the engine used to give 1.
        assert_eq!(at(12), "2");
        // Unknown names keep Excel's #NAME? — valid formulas no longer hit it.
        assert_eq!(at(13), "#NAME?");
    }

    /// The audit's book was generated by a third-party tool (openpyxl), so
    /// the formulas reach the engine through file import with no cached
    /// values, and LET is stored under its `_xlfn.` compatibility prefix.
    /// Both spellings must compute on that path too.
    #[test]
    fn a_third_party_book_with_xlfn_let_computes_through_the_file_path() {
        use std::io::Write;
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("bug-1660-openpyxl.xlsx");
        let entries: [(&str, &str); 6] = [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
            ),
            (
                "xl/styles.xml",
                concat!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
                    crate::xls_layout::default_cell_styles_xml!(),
                    "</styleSheet>",
                ),
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1"><v>5</v></c><c r="C1"><f>_xlfn.LET(x,5,x*2)</f></c></row><row r="2"><c r="A2"><v>15</v></c><c r="C2"><f>SUMPRODUCT((A1:A3&gt;10)*1)</f></c></row><row r="3"><c r="A3"><v>20</v></c><c r="C3"><f>SUM(IF(A1:A3&gt;10,1,0))</f></c></row></sheetData></worksheet>"#,
            ),
        ];
        let mut writer = zip::ZipWriter::new(std::fs::File::create(&path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, content) in entries {
            writer.start_file(name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();

        let mut cache = RecalcCache::new();
        let result = recalc_cells(
            &mut cache,
            &path,
            &[],
            &[RecalcRead {
                sheet: "Sheet1".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 2,
                    start_column: 2,
                    end_column: 2,
                },
            }],
        )
        .unwrap();
        let at = |row: u32| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row && cell.column == 2)
                .map(|cell| cell.formatted.clone())
                .unwrap()
        };
        assert_eq!(at(0), "10"); // _xlfn.LET stored form
        assert_eq!(at(1), "2"); // SUMPRODUCT over a condition
        assert_eq!(at(2), "2"); // array-context IF
    }
}

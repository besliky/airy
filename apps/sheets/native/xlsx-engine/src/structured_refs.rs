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
//! column names), so no second resolver is maintained here. Two gaps in
//! 0.8.3 are bridged by rewriting the formula text before IronCalc sees it —
//! on the recalc side only; stored formulas, the renderer model and saved
//! files keep the user-facing text, so files round-trip exactly as Excel
//! wrote them:
//!
//! - The `@` this-row shorthand Excel writes by default (`Sales[@Amount]`,
//!   `Sales[@[Unit Price]]`, `Sales[@[Jan]:[Dec]]`): IronCalc's lexer only
//!   accepts the spelled-out `[#This Row]` form. `normalize_at_shorthand`
//!   rewrites those formulas into the spelled-out form.
//! - The doubled right-bracket column escape (`Odd]]Col` for the column
//!   `Odd]Col`): IronCalc only implements the `'`-escape dialect.
//!   `normalize_escaped_brackets` rewrites the doubled form into it; left
//!   alone the formula fails to parse and the unparsable-formula pin turns
//!   the cell into a silent zero (BUG-1754).
//!
//! Defined names over structured references (`SalesTotal = SUM(
//! Sales[Amount])`) are resolved by inlining the name's body into the cell
//! formulas that use it (BUG-1753): IronCalc's importer re-parses
//! defined-name formulas against an empty table set, and its defined-name
//! parser only models cell/range references and LAMBDA definitions, so the
//! name evaluates to #NAME? wherever a cell uses it. See
//! `expand_defined_names_in_formula` for the Excel name-priority rules the
//! inlining follows.

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

/// Whether the `]` at `close` opens a doubled-bracket escape rather than
/// terminating the bracket item. Excel escapes a `]` inside a column name
/// either as `']` or as `]]` (ECMA-376 §18.17.2.5); a `]]` pair is the
/// escape when what follows it continues the column name, and structural
/// when the characters after it close the selector (`,`/`:` separators,
/// operators, whitespace, end of text). The one ambiguous shape — a column
/// name ending in `]`, written `Odd]]]` — reads the first pair as the
/// escape and the final `]` as the terminator.
fn is_escaped_bracket(chars: &[char], close: usize) -> bool {
    if chars.get(close + 1) != Some(&']') {
        return false;
    }
    match chars.get(close + 2) {
        None => false,
        Some(']') => true,
        Some(next) => !matches!(
            next,
            ',' | ':'
                | ';'
                | ')'
                | '+'
                | '-'
                | '*'
                | '/'
                | '^'
                | '&'
                | '='
                | '<'
                | '>'
                | '%'
                | ' '
                | '\t'
                | '\n'
                | '\r'
        ),
    }
}

/// Scans the selector group whose `[` sits at `open`: bracketed items
/// separated by `,`/`:` and closed by the `]` that expression text follows
/// (`SUM(Tbl[[#Data],[Odd]]Col]])` — the scan must reach the final `]`,
/// not stop at the first item's). Bare group forms — the empty `[]`, the
/// unbracketed specifier `[#All]` and a bare column name — carry no
/// doubled-bracket escapes; `None` when the group never closes.
fn scan_selector_group(chars: &[char], open: usize) -> Option<(usize, Vec<usize>)> {
    let mut cursor = open + 1;
    let mut escapes = Vec::new();
    loop {
        match chars.get(cursor)? {
            ']' => return Some((cursor + 1, escapes)),
            '#' => {
                // Bare specifier: the group's `]` directly follows it.
                while chars.get(cursor) != Some(&']') {
                    cursor += 1;
                }
                return Some((cursor + 1, escapes));
            }
            '[' => {
                let (after_item, mut item_escapes) = scan_bracketed_item_excel(chars, cursor)?;
                escapes.append(&mut item_escapes);
                cursor = after_item;
                match chars.get(cursor) {
                    Some(',') | Some(':') => cursor += 1,
                    Some(']') => return Some((cursor + 1, escapes)),
                    _ => return None,
                }
            }
            _ => {
                // Bare column name at group level (`Tbl[Odd]]Col]`): run to
                // the group's `]`, skipping `'`-escaped characters and
                // collecting the doubled `]]` escapes.
                loop {
                    match chars.get(cursor)? {
                        ']' => {
                            if is_escaped_bracket(chars, cursor) {
                                escapes.push(cursor);
                                cursor += 2;
                                continue;
                            }
                            return Some((cursor + 1, escapes));
                        }
                        '\'' => cursor += 2,
                        _ => cursor += 1,
                    }
                }
            }
        }
    }
}

/// Scans the `]`-terminated bracket item whose `[` sits at `open`, honoring
/// both escape dialects: the `'`-prefixed character (skip two) and the
/// doubled `]]` (see `is_escaped_bracket`). Returns the index just past the
/// closing `]` plus the positions of the escape pairs; `None` when the
/// formula ends inside the item.
fn scan_bracketed_item_excel(chars: &[char], open: usize) -> Option<(usize, Vec<usize>)> {
    let mut cursor = open + 1;
    let mut escapes = Vec::new();
    while cursor < chars.len() {
        match chars[cursor] {
            ']' => {
                if is_escaped_bracket(chars, cursor) {
                    escapes.push(cursor);
                    cursor += 2;
                    continue;
                }
                return Some((cursor + 1, escapes));
            }
            '\'' => cursor += 1,
            _ => {}
        }
        cursor += 1;
    }
    None
}

/// The item text with every doubled `]]` escape rewritten as the `'`
/// dialect's `']`, the form IronCalc's lexer unescapes back to `]`.
fn item_with_quote_escapes(
    chars: &[char],
    open: usize,
    after_item: usize,
    escapes: &[usize],
) -> String {
    let mut out = String::with_capacity(after_item - open);
    let mut cursor = open;
    let mut next_escape = 0usize;
    while cursor < after_item {
        if next_escape < escapes.len() && escapes[next_escape] == cursor {
            out.push_str("']");
            cursor += 2;
            next_escape += 1;
        } else {
            out.push(chars[cursor]);
            cursor += 1;
        }
    }
    out
}

/// Rewrites Excel's doubled right-bracket escape inside structured-reference
/// column items into the `'`-escape IronCalc 0.8.3 lexes:
/// `SUM(Tbl[[#Data],[Odd]]Col]])` for the column `Odd]Col` becomes
/// `SUM(Tbl[[#Data],[Odd']Col]])`. Only items that a table-name token opens
/// are touched, and a plain `[[#Data],[Amount]])` close pair is structural
/// (its second bracket is followed by the expression, not a name) — string
/// literals pass through untouched.
pub(crate) fn normalize_escaped_brackets(formula: &str) -> Cow<'_, str> {
    if !formula.contains("]]") {
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
                if ends_with_table_name(&chars, position) {
                    let Some((after_group, escapes)) = scan_selector_group(&chars, position) else {
                        // Unterminated group: copy the rest verbatim.
                        out.push(ch);
                        position += 1;
                        continue;
                    };
                    if escapes.is_empty() {
                        out.push_str(&item_copy(&chars, position, after_group));
                    } else {
                        out.push_str(&item_with_quote_escapes(
                            &chars,
                            position,
                            after_group,
                            &escapes,
                        ));
                        rewritten = true;
                    }
                    position = after_group;
                } else {
                    out.push(ch);
                    position += 1;
                }
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

/// Whether a formula text contains a structured reference: a `[` that a
/// table-name token opens a selector with. External-workbook references
/// (`[1]Sheet1!A1`, file paths in `'...'`) have no such token and read
/// false. IronCalc stores formulas internally in R1C1, where the relative
/// forms `R[0]` and `C[-2]` carry brackets too — `R` and `C` are reserved
/// names Excel never gives a table, so a bracketed digit directly after the
/// lone token reads as an R1C1 reference.
pub(crate) fn has_structured_reference(formula: &str) -> bool {
    let is_name_char = |ch: char| ch.is_alphanumeric() || matches!(ch, '_' | '\\' | '.');
    let chars: Vec<char> = formula.chars().collect();
    let mut in_quote = false;
    for (index, ch) in chars.iter().enumerate() {
        match ch {
            '\'' => in_quote = !in_quote,
            '[' if !in_quote && ends_with_table_name(&chars, index) => {
                let reserved_r1c1_token = matches!(chars[index - 1], 'R' | 'C' | 'r' | 'c')
                    && (index == 1 || !is_name_char(chars[index - 2]));
                let r1c1_relative = reserved_r1c1_token
                    && chars.get(index + 1).is_some_and(|next| {
                        next.is_ascii_digit() || matches!(next, '+' | '-' | ']')
                    });
                if !r1c1_relative {
                    return true;
                }
            }
            _ => {}
        }
    }
    false
}

/// The workbook facts the defined-name inlining needs: structured-reference
/// defined names as `(lowercased name, body)` pairs, plus the lowercased
/// table names that shadow same-named tokens.
pub(crate) struct SrefContext {
    bodies: Vec<(String, String)>,
    tables: Vec<String>,
}

impl SrefContext {
    fn is_empty(&self) -> bool {
        self.bodies.is_empty()
    }
}

/// Collects the structured-reference defined names of a workbook. IronCalc's
/// importer re-parses defined-name formulas against an empty table set and
/// its defined-name parser only models plain cell/range references and
/// LAMBDA definitions, so Excel's `SalesTotal = SUM(Sales[Amount])` —
/// entirely valid there — evaluates to #NAME? wherever a cell uses the
/// name.
pub(crate) fn sref_context(model: &Model) -> SrefContext {
    let tables = model
        .workbook
        .tables
        .keys()
        .map(|name| name.to_lowercase())
        .collect();
    let bodies = model
        .workbook
        .defined_names
        .iter()
        .filter(|defined_name| has_structured_reference(&defined_name.formula))
        .map(|defined_name| {
            (
                defined_name.name.to_lowercase(),
                defined_name.formula.trim_start_matches('=').to_string(),
            )
        })
        .collect();
    SrefContext { bodies, tables }
}

/// Replaces every use of a structured-reference defined name in a formula
/// with its parenthesized body: `=SalesTotal*2` → `=(SUM(Sales[Amount]))*2`.
///
/// Excel keeps table names and defined names in one namespace and refuses
/// duplicates, and the selector syntax (`Name[...]`) always means a table,
/// so tokens followed by `[` never expand; a token that a loaded table
/// shadows (third-party files can carry what Excel would reject) also stays
/// with the native structured-reference resolution, which already handles
/// the bare whole-table form. Tokens followed by `(` are calls and by `!`
/// sheet qualifications; string literals are skipped. Bodies that reference
/// further names resolve on the next engine pass only when the referenced
/// name itself carries a structured reference.
pub(crate) fn expand_defined_names_in_formula<'a>(
    formula: &'a str,
    context: &SrefContext,
) -> Cow<'a, str> {
    if context.bodies.is_empty() {
        return Cow::Borrowed(formula);
    }
    let is_token_char = |ch: char| ch.is_alphanumeric() || matches!(ch, '_' | '\\' | '.');
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
        if ch == '"' {
            in_string = true;
            out.push(ch);
            position += 1;
            continue;
        }
        let token_start =
            is_token_char(ch) && (position == 0 || !is_token_char(chars[position - 1]));
        if token_start {
            let mut end = position;
            while end < chars.len() && is_token_char(chars[end]) {
                end += 1;
            }
            let token: String = chars[position..end].iter().collect();
            let lower = token.to_lowercase();
            let shadowed = context.tables.iter().any(|name| *name == lower);
            let body = context
                .bodies
                .iter()
                .find(|(name, _)| *name == lower)
                .map(|(_, body)| body.as_str());
            if let Some(body) = body {
                if !shadowed && !matches!(chars.get(end), Some('[') | Some('(') | Some('!')) {
                    out.push('(');
                    out.push_str(body);
                    out.push(')');
                    rewritten = true;
                    position = end;
                    continue;
                }
            }
        }
        out.push(ch);
        position += 1;
    }
    if rewritten {
        Cow::Owned(out)
    } else {
        Cow::Borrowed(formula)
    }
}

/// The full structured-reference text preparation the engine channel runs
/// before IronCalc parses a formula: defined names over structured refs are
/// inlined (BUG-1753), the doubled `]]` column escape is rewritten into the
/// `'` dialect (the `@` pass below relies on that order), and the `@`
/// this-row shorthand is spelled out. All three rewrites are deterministic.
pub(crate) fn prepare_sref_formula<'a>(context: &SrefContext, formula: &'a str) -> Cow<'a, str> {
    match expand_defined_names_in_formula(formula, context) {
        Cow::Borrowed(expanded) => match normalize_escaped_brackets(expanded) {
            Cow::Borrowed(unescaped) => normalize_at_shorthand(unescaped),
            Cow::Owned(unescaped) => Cow::Owned(normalize_at_shorthand(&unescaped).into_owned()),
        },
        Cow::Owned(expanded) => {
            let unescaped = normalize_escaped_brackets(&expanded);
            Cow::Owned(normalize_at_shorthand(&unescaped).into_owned())
        }
    }
}

/// Rewrites every plain formula cell of an imported IronCalc model whose
/// stored text needs structured-reference preparation. Runs before
/// `pin_unparsable_formulas` on the cold-import path, so formulas Excel
/// wrote as `Sales[@Amount]`, `SalesTotal` (a defined name over a table) or
/// with `]]`-escaped columns resolve and evaluate instead of erroring (or
/// being pinned to their cached values). The rewrite reads the stored import
/// text rather than the parsed node: a broken structured reference re-parses
/// into a name-like node whose stringify no longer carries the original
/// text. Array/CSE formulas are left alone: `set_user_input` would drop
/// their CSE/spill semantics. Only the resident compute model changes —
/// never the file.
pub(crate) fn normalize_model_structured_references(model: &mut Model) {
    let context = sref_context(model);
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
                if !text.contains('[') && !text.contains('@') && context.is_empty() {
                    continue;
                }
                let Cow::Owned(normalized) = prepare_sref_formula(&context, text) else {
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

    // -- Doubled-bracket escape rewrite (BUG-1754) ---------------------------

    fn unescape(formula: &str) -> String {
        normalize_escaped_brackets(formula).into_owned()
    }

    #[test]
    fn doubled_bracket_escapes_become_quote_escapes() {
        // The audit's formula: the column `Odd]Col` — the first `]]` is the
        // escape, the final `]` closes the selector group.
        assert_eq!(
            unescape("=SUM(Tbl[[#Data],[Odd]]Col]])"),
            "=SUM(Tbl[[#Data],[Odd']Col]])".to_string()
        );
    }

    #[test]
    fn plain_selector_closes_are_never_rewritten() {
        assert_eq!(
            unescape("=SUM(Sales[[#Data],[Amount]])"),
            "=SUM(Sales[[#Data],[Amount]])".to_string()
        );
        assert_eq!(
            unescape("=COUNTA(Sales[#All])+SUM(Sales[[#Totals],[Amount]])"),
            "=COUNTA(Sales[#All])+SUM(Sales[[#Totals],[Amount]])".to_string()
        );
    }

    #[test]
    fn escape_alone_in_a_selector_and_range_forms_survive() {
        // Un-nested item form, and a column-name range whose first name ends
        // in `]` (`Jan]`): `]]` escape, then the item and group closes.
        assert_eq!(
            unescape("=SUM(Tbl[Odd]]Col])"),
            "=SUM(Tbl[Odd']Col])".to_string()
        );
        assert_eq!(
            unescape("=SUM(Tbl[[Jan]]]:[Dec]])"),
            "=SUM(Tbl[[Jan']]:[Dec]])".to_string()
        );
    }

    #[test]
    fn external_references_and_strings_are_untouched() {
        // `[1]` has no table token before the bracket.
        assert_eq!(
            unescape("=[1]Sheet1!A1+SUM(Tbl[[Amount]])"),
            "=[1]Sheet1!A1+SUM(Tbl[[Amount]])".to_string()
        );
        assert_eq!(
            unescape(r#"=IF(A1="a]]b",0,Tbl[[Amount]])"#),
            r#"=IF(A1="a]]b",0,Tbl[[Amount]])"#.to_string()
        );
    }

    // -- Defined-name inlining (BUG-1753) ------------------------------------

    fn context_with(bodies: &[(&str, &str)], tables: &[&str]) -> SrefContext {
        SrefContext {
            bodies: bodies
                .iter()
                .map(|(name, body)| (name.to_lowercase(), body.to_string()))
                .collect(),
            tables: tables.iter().map(|name| name.to_lowercase()).collect(),
        }
    }

    fn expand(formula: &str, context: &SrefContext) -> String {
        expand_defined_names_in_formula(formula, context).into_owned()
    }

    #[test]
    fn defined_name_bodies_are_inlined() {
        let context = context_with(&[("salestotal", "SUM(Sales[Amount])")], &["Sales"]);
        assert_eq!(
            expand("=SalesTotal", &context),
            "=(SUM(Sales[Amount]))".to_string()
        );
        assert_eq!(
            expand("=SalesTotal*2+1", &context),
            "=(SUM(Sales[Amount]))*2+1".to_string()
        );
        // Case-insensitive, but never a substring of a longer identifier.
        assert_eq!(
            expand("=salestotal+mySalesTotal", &context),
            "=(SUM(Sales[Amount]))+mySalesTotal".to_string()
        );
    }

    #[test]
    fn selector_syntax_and_shadows_keep_the_table() {
        let context = context_with(&[("sales", "SUM(Sales[Amount])")], &["Sales"]);
        // `Sales[Amount]` is the table's structured reference, not the name.
        assert_eq!(
            expand("=SUM(Sales[Amount])", &context),
            "=SUM(Sales[Amount])".to_string()
        );
        // A loaded table shadows a same-named defined name (Excel keeps the
        // namespace duplicate-free; third-party files resolve to the table).
        assert_eq!(expand("=Sales", &context), "=Sales".to_string());
    }

    #[test]
    fn calls_sheets_and_strings_never_expand() {
        let context = context_with(&[("mylambda", "LAMBDA(x,SUM(Sales[Amount]))")], &[]);
        assert_eq!(expand("=MyLambda(3)", &context), "=MyLambda(3)".to_string());
        let context = context_with(&[("sheetlike", "SUM(Sales[Amount])")], &[]);
        assert_eq!(
            expand("=SheetLike!A1+\"sheetlike\"", &context),
            "=SheetLike!A1+\"sheetlike\"".to_string()
        );
    }

    #[test]
    fn full_preparation_chains_inlining_escapes_and_at() {
        let context = context_with(&[("salestotal", "SUM(Sales[Amount])")], &["Sales"]);
        assert_eq!(
            prepare_sref_formula(&context, "=SalesTotal").into_owned(),
            "=(SUM(Sales[Amount]))".to_string()
        );
        let context = context_with(&[], &["Tbl"]);
        assert_eq!(
            prepare_sref_formula(&context, "=SUM(Tbl[@[Odd]]Col]])").into_owned(),
            "=SUM(Tbl[[#This Row],[Odd']Col]])".to_string()
        );
    }

    #[test]
    fn has_structured_reference_distinguishes_selectors_from_external_refs() {
        assert!(has_structured_reference("SUM(Sales[Amount])"));
        assert!(has_structured_reference("SUM(Sales[Amount])+VAT"));
        assert!(!has_structured_reference("SUM(Sheet1!A1:A4)"));
        assert!(!has_structured_reference("[1]Sheet1!A1"));
        assert!(!has_structured_reference(
            r#"'C:\Path\[Book.xlsx]Sheet1'!A1"#
        ));
        assert!(has_structured_reference(
            r#"'C:\Path\[Book.xlsx]Sheet1'!A1+SUM(Sales[Amount])"#
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

    /// Optional parts the BUG-1753/BUG-1754 fixtures add to the base
    /// two-table book: the defined name `SalesTotal = SUM(Sales[Amount])`
    /// and a second sheet whose `Tbl` table carries the `]]`-escaped column
    /// `Odd]Col` (data 2/4/6/8, totals 20).
    struct BookExtras<'a> {
        row2_cell: &'a str,
        row3_cell: &'a str,
        defined_name: bool,
        escape_table: bool,
    }

    impl Default for BookExtras<'_> {
        fn default() -> Self {
            Self {
                row2_cell: "",
                row3_cell: "",
                defined_name: false,
                escape_table: false,
            }
        }
    }

    /// `row2_cell`/`row3_cell` inject optional `<c>` elements at the end of
    /// rows 2/3 (the cold import test carries file-side formulas there).
    fn write_table_workbook(path: &Path, row2_cell: &str, row3_cell: &str) {
        write_workbook(
            path,
            &BookExtras {
                row2_cell,
                row3_cell,
                ..BookExtras::default()
            },
        );
    }

    fn write_workbook(path: &Path, extras: &BookExtras<'_>) {
        let row2_cell = extras.row2_cell;
        let row3_cell = extras.row3_cell;
        let sheet = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C6"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Product</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Unit Price</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>alpha</t></is></c><c r="B2"><v>10</v></c><c r="C2"><v>1.5</v></c>{row2_cell}</row><row r="3"><c r="A3" t="inlineStr"><is><t>beta</t></is></c><c r="B3"><v>20</v></c><c r="C3"><v>2.5</v></c>{row3_cell}</row><row r="4"><c r="A4" t="inlineStr"><is><t>gamma</t></is></c><c r="B4"><v>30</v></c><c r="C4"><v>3.5</v></c></row><row r="5"><c r="A5" t="inlineStr"><is><t>delta</t></is></c><c r="B5"><v>40</v></c><c r="C5"><v>4.5</v></c></row><row r="6"><c r="A6" t="inlineStr"><is><t>Total</t></is></c><c r="B6"><v>100</v></c></row></sheetData><tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts></worksheet>"#
        );
        let sheet2 = r#"<sheet name="Sheet2" sheetId="2" r:id="rId3"/>"#;
        let defined_names = if extras.defined_name {
            r#"<definedNames><definedName name="SalesTotal">SUM(Sales[Amount])</definedName></definedNames>"#
        } else {
            ""
        };
        let workbook_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/>{}</sheets>{}</workbook>"#,
            if extras.escape_table { sheet2 } else { "" },
            defined_names
        );
        let workbook_rels = if extras.escape_table {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#
        } else {
            WORKBOOK_RELS
        };
        let content_types = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>{}</Types>"#,
            if extras.escape_table {
                r#"<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>"#
            } else {
                ""
            }
        );
        let sheet2_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:D4"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Odd]Col</t></is></c></row><row r="2"><c r="A2"><v>2</v></c><c r="D2"><f>SUM(Sales[Amount])</f></c></row><row r="3"><c r="A3"><v>4</v></c><c r="D3"><f>SalesTotal</f></c></row><row r="4"><c r="A4"><v>6</v></c><c r="D4"><f>SUM(Tbl[[#Data],[Odd]]Col]])</f><v>0</v></c></row><row r="5"><c r="A5"><v>8</v></c><c r="D5"><f>SUM(Tbl[[#Data],[Nope]]Col]])</f><v>7</v></c></row><row r="6"><c r="A6"><v>20</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>"#;
        let dir = path.parent().expect("fixture parent");
        std::fs::create_dir_all(dir.join("xl/worksheets/_rels")).unwrap();
        std::fs::create_dir_all(dir.join("xl/tables")).unwrap();
        let mut writer = zip::ZipWriter::new(File::create(path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        let mut entries: Vec<(&str, String)> = vec![
            ("[Content_Types].xml", content_types),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.to_string(),
            ),
            ("xl/workbook.xml", workbook_xml),
            ("xl/_rels/workbook.xml.rels", workbook_rels.to_string()),
            ("xl/styles.xml", STYLES_XML.to_string()),
            ("xl/worksheets/sheet1.xml", sheet),
            ("xl/worksheets/_rels/sheet1.xml.rels", SHEET_RELS.to_string()),
            ("xl/tables/table1.xml", SALES_TABLE_XML.to_string()),
            ("xl/tables/table2.xml", PRICES_TABLE_XML.to_string()),
        ];
        if extras.escape_table {
            entries.push(("xl/worksheets/sheet2.xml", sheet2_xml.to_string()));
            entries.push((
                "xl/worksheets/_rels/sheet2.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table3.xml"/></Relationships>"#.to_string(),
            ));
            entries.push((
                "xl/tables/table3.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="3" name="Tbl" displayName="Tbl" ref="A1:A6" totalsRowCount="1" headerRowCount="1"><autoFilter ref="A1:A5"/><tableColumns count="1"><tableColumn id="1" name="Odd]Col"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>"#.to_string(),
            ));
        }
        for (name, content) in entries {
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

    fn write_bugfix_workbook(path: &Path) {
        write_workbook(
            path,
            &BookExtras {
                defined_name: true,
                escape_table: true,
                ..BookExtras::default()
            },
        );
    }

    fn sheet2_edit(row: u32, column: u32, input: &str) -> RecalcEdit {
        RecalcEdit {
            sheet: "Sheet2".to_string(),
            row,
            column,
            input: input.to_string(),
        }
    }

    /// BUG-1753: the defined name `SalesTotal = SUM(Sales[Amount])` (valid
    /// Excel) must compute through the recalc channel like the same formula
    /// typed directly, on the cold file path (Sheet2!D3) and the edit path.
    #[test]
    fn defined_name_over_a_structured_reference_computes() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("tables.xlsx");
        write_bugfix_workbook(&path);
        let mut cache = RecalcCache::new();
        // Cold import: D2 carries the formula directly (control: 100), D3
        // uses the defined name (Excel resolves it to 100 as well).
        let cold = recalc_cells(
            &mut cache,
            &path,
            &[],
            &[read_cell("Sheet2", 1, 3), read_cell("Sheet2", 2, 3)],
        )
        .unwrap();
        let number = |row: u32| {
            cold.cells
                .iter()
                .find(|cell| cell.row == row)
                .and_then(|cell| cell.number)
        };
        assert_eq!(number(1), Some(100.0), "direct formula control");
        assert_eq!(number(2), Some(100.0), "defined name must not be #NAME?");
        // Edit path: the same name typed into a cell computes too.
        let typed = recalc_cells(
            &mut cache,
            &path,
            &[sheet2_edit(5, 3, "=SalesTotal")],
            &[read_cell("Sheet2", 5, 3)],
        )
        .unwrap();
        assert_eq!(typed.cells[0].number, Some(100.0));
        assert!(typed.cells[0].is_formula);
    }

    /// BUG-1754: the column `Odd]Col` is written with Excel's doubled-bracket
    /// escape. The formula must parse and evaluate to 20 — neither pinned to
    /// a fabricated 0 literal (the file carries no cached value) nor left as
    /// a non-formula.
    #[test]
    fn doubled_bracket_escape_computes_instead_of_a_silent_zero() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("tables.xlsx");
        write_bugfix_workbook(&path);
        let mut cache = RecalcCache::new();
        // Cold import: Sheet2!D4 holds the escaped formula with a cached
        // 0 the audit's corpus carried.
        let cold = recalc_cells(&mut cache, &path, &[], &[read_cell("Sheet2", 3, 3)]).unwrap();
        assert_eq!(cold.cells[0].number, Some(20.0));
        assert!(cold.cells[0].is_formula);
        // A structured-reference formula that still fails to resolve (the
        // column `Nope]Col` does not exist) must not be pinned to its cached
        // value either: the cell stays a formula carrying the engine error —
        // an indication — instead of a silent literal.
        let pinned = recalc_cells(&mut cache, &path, &[], &[read_cell("Sheet2", 4, 3)]).unwrap();
        assert!(pinned.cells[0].is_formula);
        assert_eq!(pinned.cells[0].number, None);
        // Edit path: totals band through the same escaped column.
        let typed = recalc_cells(
            &mut cache,
            &path,
            &[sheet2_edit(6, 3, "=SUM(Tbl[[#Totals],[Odd]]Col]])")],
            &[read_cell("Sheet2", 6, 3)],
        )
        .unwrap();
        assert_eq!(typed.cells[0].number, Some(20.0));
        assert!(typed.cells[0].is_formula);
    }
}

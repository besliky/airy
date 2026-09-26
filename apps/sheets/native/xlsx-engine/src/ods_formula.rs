//! ODF formula -> OOXML formula translation for the .ods convert path
//! (BUG-1661).
//!
//! calamine hands the converter .ods formulas verbatim in ODF formula
//! syntax — `of:=SUM([.A1:.A2])` — which is byte-parity with what
//! LibreOffice itself writes into an .xlsx on export, but nothing on the
//! Airy side evaluates it: IronCalc fails to parse it, the cell goes
//! `#ERROR!`/pinned, and the UI shows dead cells. This module rewrites the
//! ODF subset that has a direct OOXML equivalent and refuses anything
//! else: `translate` returns `None` when it meets a construct it does not
//! fully understand, and the caller keeps the original text verbatim
//! (IronCalc then pins the file's cached value, so the cell still shows
//! the number the source carried).
//!
//! Understood subset, derived from LibreOffice's own ods->xlsx rendering
//! (the pinned reference in `bug-1661-ods-formulas.ods`):
//! - the `of:` formula namespace prefix, with its optional `=`;
//! - bracketed references: `[.A1]` -> `A1`, `[.A1:.A5]` -> `A1:A5`,
//!   `[.$A$1]` -> `$A$1`, `[Data.B1]` -> `Data!B1`,
//!   `[$'My Sheet'.$A$1]` -> `'My Sheet'!$A$1`, whole `[.A:.B]` /
//!   `[.1:.3]` spans;
//! - the `;` argument separator -> `,`;
//! - the `COM.MICROSOFT.` function namespace prefix dropped
//!   (`COM.MICROSOFT.CONCAT` -> `CONCAT`). `TRUE()`/`FALSE()` and every
//!   other function name spell identically in ODF and OOXML, so the
//!   rename table is exactly this one entry.
//!
//! Deliberately NOT translated (verbatim pass-through): external-document
//! references (`['file:///x.ods'#Data.B1]`), 3-D ranges
//! (`[A.A1:B.A5]`), `ORG.OPENOFFICE.*` names, and anything bracketed that
//! does not parse as the reference grammar above.

/// Translates an ODF defined-name payload into OOXML definedName content:
/// a named range arrives as a cell-range-address (`$Data.$A$1:.$A$4`,
/// `$Calc.$A$1`), a named expression as an `of:` formula. `None` = not
/// understood — the caller must skip the name rather than carry ODF syntax
/// into a book the engine has to open.
pub(crate) fn translate_defined_name(raw: &str) -> Option<String> {
    if raw.contains("of:") {
        return translate(raw);
    }
    if raw.contains('!') {
        // The .xls rendering (calamine resolves NAME records to
        // `Sheet!$A$1:$A$4`): validated and re-quoted, carried as-is.
        let (sheet, refs) = raw.split_once('!')?;
        if sheet.is_empty() || refs.contains('!') {
            return None;
        }
        if !refs.split(':').all(is_a1_address) {
            return None;
        }
        let plain = sheet
            .chars()
            .all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
            && !looks_like_address(&sheet);
        let prefix = if plain {
            sheet.to_string()
        } else {
            format!("'{}'", sheet.replace('\'', "''"))
        };
        return Some(format!("{prefix}!{refs}"));
    }
    let (first, second) = match raw.split_once(':') {
        Some((first, second)) => (first, Some(second)),
        None => (raw, None),
    };
    let first = parse_range_part(first)?;
    let sheet = first.sheet?;
    let mut out = format!("{sheet}!{}", first.address);
    if let Some(second) = second {
        let second = parse_range_part(second)?;
        // A different sheet on the second endpoint is a 3-D range — beyond
        // what a defined-name payload should carry, skip it.
        if second.sheet.is_some_and(|other| other != sheet) {
            return None;
        }
        out.push(':');
        out.push_str(&second.address);
    }
    Some(out)
}

/// One endpoint of a cell-range-address: the (already Excel-rendered)
/// optional sheet prefix and the address itself.
struct RangePart {
    sheet: Option<String>,
    address: String,
}

/// Parses one `$?Sheet.?A1` endpoint of a cell-range-address. The sheet
/// name keeps Excel quoting rules (bare when it cannot read as an address,
/// single-quoted otherwise).
fn parse_range_part(part: &str) -> Option<RangePart> {
    let chars: Vec<char> = part.chars().collect();
    let mut cursor = 0;
    if chars.first() == Some(&'$') {
        cursor = 1;
    }
    // Sheet part: quoted `'My Sheet'.` or bare `Data.`; a bare part may
    // also start directly with the address (`.$A$1` continues the range).
    let sheet = if chars.get(cursor) == Some(&'\'') {
        let mut name = String::new();
        cursor += 1;
        loop {
            match chars.get(cursor) {
                Some('\'') if chars.get(cursor + 1) == Some(&'\'') => {
                    name.push('\'');
                    cursor += 2;
                }
                Some('\'') => {
                    cursor += 1;
                    break;
                }
                None => return None,
                Some(ch) => {
                    name.push(*ch);
                    cursor += 1;
                }
            }
        }
        if chars.get(cursor) != Some(&'.') {
            return None;
        }
        cursor += 1;
        Some(format!("'{}'", name.replace('\'', "''")))
    } else {
        let start = cursor;
        while cursor < chars.len()
            && (chars[cursor].is_ascii_alphanumeric() || chars[cursor] == '_')
        {
            cursor += 1;
        }
        if chars.get(cursor) != Some(&'.') {
            // No sheet part at all: everything so far must be the address.
            cursor = 0;
            None
        } else if cursor == start {
            // A bare `.` (the `.$A$4` of a range's second endpoint): the
            // range continues on the first endpoint's sheet.
            cursor += 1;
            None
        } else {
            cursor += 1;
            let name: String = chars[start..cursor - 1].iter().collect();
            // A plain name that cannot read as an address stays bare; a
            // quoted-form name (spaces, or one that reads as an address) is
            // quoted exactly like a formula reference.
            let plain = !name.is_empty()
                && name
                    .chars()
                    .all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
                && !looks_like_address(&name);
            Some(if plain {
                name
            } else {
                format!("'{}'", name.replace('\'', "''"))
            })
        }
    };
    let address = parse_address(&chars, &mut cursor)?;
    if cursor != chars.len() {
        return None;
    }
    Some(RangePart { sheet, address })
}

/// An absolute-ish A1 address (`$A$1`, `A1`) — the only endpoint shape
/// calamine's NAME renderer produces.
fn is_a1_address(address: &str) -> bool {
    let body = address.trim_start_matches('$');
    let (letters, digits) = body.split_at(
        body.find(|ch: char| !ch.is_ascii_alphabetic())
            .unwrap_or(body.len()),
    );
    (1..=3).contains(&letters.len())
        && digits
            .trim_start_matches('$')
            .parse::<u32>()
            .is_ok_and(|row| row >= 1)
}

/// Whether a name is one Excel accepts in `<definedName name="...">`:
/// letters/digits/underscore/period, not starting with a digit, and not
/// reading as an A1 address. Everything else (built-in names, control
/// characters, `R1C1` lookalikes) is skipped rather than mis-emitted.
pub(crate) fn is_valid_defined_name(name: &str) -> bool {
    if name.is_empty() || name.len() > 255 {
        return false;
    }
    if name.to_ascii_lowercase().starts_with("_xlnm") {
        return false;
    }
    let starts = !name.starts_with(|ch: char| ch.is_ascii_digit() || ch == '$')
        && name.starts_with(|ch: char| ch.is_alphanumeric() || matches!(ch, '_' | '\\'));
    starts
        && name
            .chars()
            .all(|ch| ch.is_alphanumeric() || matches!(ch, '_' | '.' | '\\'))
        && !looks_like_address(name)
}

/// Rewrites an ODF formula into OOXML syntax. `None` means "not ODF, or
/// not fully understood" — the original text must be carried verbatim.
pub fn translate(raw: &str) -> Option<String> {
    let body = raw.strip_prefix("of:")?;
    let body = body.strip_prefix('=').unwrap_or(body);
    let chars: Vec<char> = body.chars().collect();
    let mut out = String::with_capacity(body.len());
    let mut index = 0;
    while index < chars.len() {
        match chars[index] {
            // String literals survive untouched, `"` doubled per ODF; a
            // `;` or `[` inside them is data, not syntax.
            '"' => {
                let literal = copy_string_literal(&chars, &mut index)?;
                out.push_str(&literal);
            }
            '[' => translate_reference(&chars, &mut index, &mut out)?,
            ';' => {
                out.push(',');
                index += 1;
            }
            _ => {
                if let Some((identifier, next)) = read_identifier(&chars, index) {
                    // The single rename-table entry: drop the Microsoft
                    // function namespace, keep the bare OOXML name.
                    if let Some(name) = identifier.strip_prefix("COM.MICROSOFT.") {
                        out.push_str(name);
                    } else {
                        out.push_str(&identifier);
                    }
                    index = next;
                } else {
                    out.push(chars[index]);
                    index += 1;
                }
            }
        }
    }
    Some(out)
}

/// Copies a quoted string literal starting at `chars[*index] == '"'`,
/// advancing past its closing quote. `None` on an unterminated literal —
/// the formula is malformed, carry it verbatim.
fn copy_string_literal(chars: &[char], index: &mut usize) -> Option<String> {
    let start = *index;
    let mut cursor = start + 1;
    while cursor < chars.len() {
        if chars[cursor] == '"' {
            // ODF escapes an embedded quote by doubling it.
            if chars.get(cursor + 1) == Some(&'"') {
                cursor += 2;
                continue;
            }
            *index = cursor + 1;
            return Some(chars[start..=cursor].iter().collect());
        }
        cursor += 1;
    }
    None
}

/// Reads a maximal function-name identifier (`SUM`, `ERROR.TYPE`,
/// `COM.MICROSOFT.CONCAT`) at `index`; returns the identifier and the
/// position after it, or `None` when the character cannot start one.
fn read_identifier(chars: &[char], index: usize) -> Option<(String, usize)> {
    let starts = chars[index].is_ascii_alphabetic() || chars[index] == '_' || chars[index] == '.';
    if !starts {
        return None;
    }
    let mut end = index;
    while end < chars.len() && (chars[end].is_ascii_alphanumeric() || chars[end] == '.') {
        end += 1;
    }
    Some((chars[index..end].iter().collect(), end))
}

/// Translates one bracketed reference `[...]` into Excel notation,
/// advancing `index` past the closing `]`. `None` (verbatim) on anything
/// outside the understood grammar — external documents, 3-D ranges,
/// malformed addresses.
fn translate_reference(chars: &[char], index: &mut usize, out: &mut String) -> Option<()> {
    let mut cursor = *index + 1; // past '['
    let mut sheet = parse_sheet_part(chars, &mut cursor)?;
    let first = parse_address(chars, &mut cursor)?;
    let mut range = String::new();
    if chars.get(cursor) == Some(&':') {
        cursor += 1;
        // The second endpoint may repeat the sheet (`[Data.A1:Data.A5]`)
        // or omit it (`[Data.A1:.A5]`); a different sheet is a 3-D range —
        // not understood, the caller carries the formula verbatim.
        let second_sheet = parse_sheet_part(chars, &mut cursor)?;
        let second = parse_address(chars, &mut cursor)?;
        if let Some(other) = second_sheet {
            match &sheet {
                Some(own) if *own == other => {}
                Some(_) => return None,
                None => sheet = Some(other),
            }
        }
        range = format!(":{second}");
    }
    if chars.get(cursor) != Some(&']') {
        return None;
    }
    cursor += 1;
    *index = cursor;
    if let Some(name) = sheet {
        out.push_str(&name);
        out.push('!');
    }
    out.push_str(&first);
    out.push_str(&range);
    Some(())
}

/// Parses the optional sheet part (`Data.`, `$Data.`, `'My Sheet'.`,
/// `[$'My Sheet'.`) and renders it in Excel notation — unquoted for plain
/// names that cannot be read as an address, single-quoted otherwise.
fn parse_sheet_part(chars: &[char], cursor: &mut usize) -> Option<Option<String>> {
    let mut local = *cursor;
    if chars.get(local) == Some(&'$') {
        // ODF absolute-sheet marker; OOXML has no equivalent, drop it.
        local += 1;
    }
    let name = match chars.get(local) {
        Some('\'') => {
            let mut name = String::new();
            local += 1;
            loop {
                match chars.get(local) {
                    Some('\'') => {
                        if chars.get(local + 1) == Some(&'\'') {
                            name.push('\'');
                            local += 2;
                            continue;
                        }
                        local += 1;
                        break;
                    }
                    // Sheet names cannot contain brackets; refuse instead
                    // of guessing where the quoted name ends.
                    Some('[') | Some(']') | None => return None,
                    Some(ch) => {
                        name.push(*ch);
                        local += 1;
                    }
                }
            }
            name
        }
        Some(ch) if ch.is_ascii_alphanumeric() || *ch == '_' => {
            let start = local;
            while local < chars.len()
                && (chars[local].is_ascii_alphanumeric() || chars[local] == '_')
            {
                local += 1;
            }
            chars[start..local].iter().collect()
        }
        _ => String::new(),
    };
    if chars.get(local) != Some(&'.') {
        return None;
    }
    local += 1;
    *cursor = local;
    if name.is_empty() {
        // Bare `[.A1]` / `[.$A$1]` — the current sheet, nothing to emit.
        return Some(None);
    }
    if name
        .chars()
        .all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
        && !looks_like_address(&name)
    {
        Some(Some(name))
    } else {
        // Quoted form; an embedded quote is doubled per both grammars.
        Some(Some(format!("'{}'", name.replace('\'', "''"))))
    }
}

/// Whether a bare name would parse as a cell address (`A1`, `XFD1048576`)
/// and therefore must stay quoted when used as a sheet name.
fn looks_like_address(name: &str) -> bool {
    let split = name
        .char_indices()
        .find(|(_, ch)| ch.is_ascii_digit())
        .map(|(index, _)| index);
    let Some(split) = split else {
        return false;
    };
    let (letters, digits) = name.split_at(split);
    !letters.is_empty()
        && !digits.is_empty()
        && letters.chars().all(|ch| ch.is_ascii_alphabetic())
        && digits.chars().all(|ch| ch.is_ascii_digit())
}

/// Parses one cell/column/row address (`A$1`, `$A$1`, `A`, `3`) after the
/// sheet dot, rendered in Excel notation.
fn parse_address(chars: &[char], cursor: &mut usize) -> Option<String> {
    let mut local = *cursor;
    let mut address = String::new();
    if chars.get(local) == Some(&'$') {
        address.push('$');
        local += 1;
    }
    let mut letters = 0;
    while local < chars.len() && chars[local].is_ascii_alphabetic() {
        address.push(chars[local]);
        local += 1;
        letters += 1;
    }
    if chars.get(local) == Some(&'$') {
        address.push('$');
        local += 1;
    }
    let mut digits = 0;
    while local < chars.len() && chars[local].is_ascii_digit() {
        address.push(chars[local]);
        local += 1;
        digits += 1;
    }
    // Whole column (letters only), whole row (digits only) or a cell
    // (both); anything else (`.` or an empty token) is not understood, and
    // Excel columns stop at XFD.
    let well_formed = letters <= 3
        && ((letters > 0 && digits == 0)
            || (letters == 0 && digits > 0)
            || (letters > 0 && digits > 0));
    if !well_formed {
        return None;
    }
    *cursor = local;
    Some(address)
}

#[cfg(test)]
mod tests {
    use super::translate;

    fn translated(raw: &str) -> String {
        translate(raw).expect("expected the formula to translate")
    }

    fn verbatim(raw: &str) {
        assert_eq!(
            translate(raw),
            None,
            "expected verbatim pass-through: {raw}"
        );
    }

    #[test]
    fn strips_the_of_prefix() {
        assert_eq!(translated("of:=SUM([.A1:.A5])"), "SUM(A1:A5)");
        // The `=` after the namespace is optional per ODF.
        assert_eq!(translated("of:SUM([.A1:.A5])"), "SUM(A1:A5)");
        assert_eq!(translated("of:=[.A1]*[.B1]"), "A1*B1");
    }

    #[test]
    fn translates_absolute_and_relative_addresses() {
        // ODF keeps the current-sheet dot even on absolute addresses —
        // this is exactly how LibreOffice writes `=$A$1*2` (BUG-1661).
        assert_eq!(translated("of:=[.$A$1]*2"), "$A$1*2");
        assert_eq!(translated("of:=[.$A$1]+[.B$2]+[.$C3]"), "$A$1+B$2+$C3");
    }

    #[test]
    fn translates_cross_sheet_references() {
        assert_eq!(
            translated("of:=SUM([.A1:.A2])+[Data.B1]"),
            "SUM(A1:A2)+Data!B1"
        );
        assert_eq!(translated("of:=SUM([Data.A1:.A5])"), "SUM(Data!A1:A5)");
        assert_eq!(translated("of:=SUM([Data.A1:Data.A5])"), "SUM(Data!A1:A5)");
        assert_eq!(translated("of:=[$'My Sheet'.$A$1]"), "'My Sheet'!$A$1");
        // A quoted name that would read as an address stays quoted.
        assert_eq!(translated("of:=['A1'.$B$2]"), "'A1'!$B$2");
        // A plain name containing no digits is safe unquoted.
        assert_eq!(translated("of:=[$Summary_2024.$B$1]"), "Summary_2024!$B$1");
    }

    #[test]
    fn translates_whole_column_and_row_spans() {
        assert_eq!(translated("of:=SUM([.A:.B])"), "SUM(A:B)");
        assert_eq!(translated("of:=SUM([.1:.3])"), "SUM(1:3)");
        assert_eq!(
            translated("of:=SUM([Sheet.A1:Sheet.B3])"),
            "SUM(Sheet!A1:B3)"
        );
    }

    #[test]
    fn swaps_argument_separators_outside_strings() {
        assert_eq!(
            translated("of:=IF([.A3]>25;\"big\";\"small\")"),
            "IF(A3>25,\"big\",\"small\")"
        );
        // A `;` inside a literal is data.
        assert_eq!(
            translated("of:=IF([.A1];\"a;b\";\"c\")"),
            "IF(A1,\"a;b\",\"c\")"
        );
    }

    #[test]
    fn drops_the_microsoft_function_namespace() {
        assert_eq!(
            translated("of:=COM.MICROSOFT.CONCAT([.A1:.A2])"),
            "CONCAT(A1:A2)"
        );
        // Dotted OOXML names pass through whole.
        assert_eq!(translated("of:=ERROR.TYPE([.A1])"), "ERROR.TYPE(A1)");
        assert_eq!(translated("of:=IF(TRUE();1;2)"), "IF(TRUE(),1,2)");
    }

    #[test]
    fn carries_unknown_constructs_verbatim() {
        // External-document references are beyond the grammar.
        verbatim("of:=SUM(['file:///tmp/x.ods'#Data.B1])");
        // 3-D ranges (different sheets per endpoint).
        verbatim("of:=SUM([Data.A1:Other.A5])");
        // Malformed bracket content.
        verbatim("of:=SUM([.])");
        verbatim("of:=X[]");
        // Unterminated string literal.
        verbatim("of:=\"open");
        // A number glued to the bracket tail.
        verbatim("of:=[.A1:.A5x]");
    }

    #[test]
    fn leaves_non_odf_formulas_untouched() {
        // xlsx/.xls sources arrive already in OOXML syntax: without the
        // `of:` prefix nothing is rewritten, so those conversions stay
        // byte-identical.
        verbatim("SUM(A1:A2)");
        verbatim("B1*2");
        verbatim("");
    }
}

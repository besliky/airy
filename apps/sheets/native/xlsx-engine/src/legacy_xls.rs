//! Legacy .xls string repair for the convert path (BUG-1600).
//!
//! calamine builds one workbook-wide decoder from the BIFF CODEPAGE record
//! and applies it to every BIFF8 string. For a single-byte codepage (1252,
//! 1251, ...) that is wrong twice: a string flagged fHighByte=1 (UTF-16LE
//! per [MS-XLS] XLUnicodeString) is fed raw through the single-byte table,
//! so every Cyrillic — or any non-Latin — character becomes mojibake plus
//! a control character, and even 8-bit compressed strings come back
//! NUL-padded. Sheet names mangle the same way, so a whole tender workbook
//! turns into unreadable text (the UR-01 report).
//!
//! The converter therefore re-reads the Workbook stream itself: a minimal,
//! hardened compound-file (OLE2) reader locates the stream, a BIFF walk
//! collects CODEPAGE, SST (+CONTINUE records), BOUNDSHEET and LABELSST
//! records, and the correctly decoded strings are overlaid on calamine's
//! output at conversion time. The overlay activates only for BIFF8 books
//! with a single-byte codepage — UTF-8/UTF-16-codepage books (what
//! LibreOffice and Excel usually write) convert exactly as before.
//!
//! Everything here is best-effort: any structural surprise (short header,
//! broken sector chain, truncated record) yields an empty overlay and the
//! conversion falls back to calamine's own output, exactly as before this
//! module existed. Records are only ever read through bounds-checked
//! slices, chains carry cycle guards, and no allocation is sized by
//! attacker-controlled lengths.

use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;

use codepage::to_encoding;
use encoding_rs::Encoding;

/// BIFF record ids used by the walk ([MS-XLS] 2.4 record enumeration).
const REC_EOF: u16 = 0x000A;
const REC_CONTINUE: u16 = 0x003C;
const REC_CODEPAGE: u16 = 0x0042;
const REC_BOUNDSHEET: u16 = 0x0085;
const REC_LABELSST: u16 = 0x00FD;
const REC_BOF: u16 = 0x0809;
const REC_SST: u16 = 0x00FC;

/// BOF header version for BIFF8 (biffVersion 0x0600); older BIFF books
/// carry no SST and calamine's single-byte path already reads them right.
const BIFF8_BOF_VERSION: u16 = 0x0600;

/// Workbook codepage calamine assumes when CODEPAGE is absent — mirrored
/// so the overlay's gate sees the same decoder calamine uses.
const DEFAULT_CODEPAGE: u16 = 1200;

/// Sector and free-chain sentinels at or above this terminate a chain
/// (ENDOFCHAIN, FREESECT, FATSECT, DIFSECT — [MS-CFB] 2.2).
const CHAIN_END: u32 = 0xFFFF_FFFA;

/// OLE2 compound-file magic bytes ([MS-CFB] 2.2).
const CFB_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// Cap on one decoded stream: beyond it the overlay degrades to calamine's
/// output rather than buffering an absurd allocation.
const MAX_STREAM_BYTES: u64 = 512 * 1024 * 1024;

/// Correctly decoded strings for one legacy workbook, aligned with
/// calamine's sheet order (both follow BOUNDSHEET record order).
#[derive(Default)]
pub struct LegacyXlsStrings {
    sheets: Vec<SheetStrings>,
}

/// Repaired sheet name and LABELSST cell texts for one sheet.
#[derive(Default)]
pub struct SheetStrings {
    pub name: Option<String>,
    pub cells: HashMap<(u32, u32), String>,
}

impl LegacyXlsStrings {
    /// Best-effort extraction; never fails and never panics on hostile
    /// input — anything unexpected yields an empty overlay.
    pub fn extract(source: &Path) -> Self {
        let mut overlay = Self::default();
        // Content-based, not extension-based (SEC-1301 style): only an OLE2
        // compound document — a legacy .xls — can carry the structures the
        // walk understands; zip-based sources skip all of it.
        let Ok(mut file) = File::open(source) else {
            return overlay;
        };
        let mut magic = [0u8; 8];
        if file.read_exact(&mut magic).is_err() || magic != CFB_MAGIC {
            return overlay;
        }
        if let Some(sheets) = extract_repaired_strings(file) {
            overlay.sheets = sheets;
        }
        overlay
    }

    /// Repaired strings for the sheet at calamine's index, if any.
    pub fn sheet(&self, index: usize) -> Option<&SheetStrings> {
        self.sheets.get(index).filter(|sheet| sheet.relevant())
    }
}

impl SheetStrings {
    fn relevant(&self) -> bool {
        self.name.is_some() || !self.cells.is_empty()
    }
}

/// Parse the compound file and walk its Workbook stream. `None` anywhere
/// means "leave the conversion to calamine unchanged".
fn extract_repaired_strings(file: File) -> Option<Vec<SheetStrings>> {
    let mut compound = CompoundFile::open(file)?;
    let stream = compound.read_workbook_stream()?;
    repair_stream_strings(&stream)
}

/// Decode UTF-16LE bytes the way [MS-XLS] fHighByte=1 strings require.
fn decode_utf16le(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| u16::from_le_bytes(*pair))
        .collect();
    String::from_utf16_lossy(&units)
}

fn read_u16(bytes: &[u8]) -> Option<u16> {
    Some(u16::from_le_bytes([*bytes.first()?, *bytes.get(1)?]))
}

fn read_u32(bytes: &[u8]) -> Option<u32> {
    Some(u32::from_le_bytes([
        *bytes.first()?,
        *bytes.get(1)?,
        *bytes.get(2)?,
        *bytes.get(3)?,
    ]))
}

// ---------------------------------------------------------------------------
// Compound file (OLE2) reader — just enough of [MS-CFB] to pull one stream.
// ---------------------------------------------------------------------------

struct CompoundFile {
    file: File,
    sector_size: u64,
    mini_size: u64,
    /// Stream size at or above which a stream lives in regular sectors
    /// instead of the mini stream ([MS-CFB] 3.2.1, typically 4096).
    mini_cutoff: u64,
    fat: Vec<u32>,
    minifat: Vec<u32>,
    /// Regular sectors holding the mini stream container (the root entry's
    /// stream); mini sectors are 64-byte slices of it.
    mini_container: Vec<u32>,
    /// Location of the Workbook/Book stream in the directory.
    workbook: Option<(u32, u64)>,
}

impl CompoundFile {
    fn open(mut file: File) -> Option<Self> {
        let file_len = file.metadata().ok()?.len();
        file.seek(SeekFrom::Start(0)).ok()?;
        let mut header = [0u8; 512];
        file.read_exact(&mut header).ok()?;
        if header[..8] != CFB_MAGIC {
            return None;
        }
        let sector_size = 1u64 << read_u16(&header[30..32])?;
        let mini_size = 1u64 << read_u16(&header[32..34])?;
        // Only the documented geometries; anything else is hostile.
        if sector_size != 512 && sector_size != 4096 {
            return None;
        }
        if mini_size != 64 {
            return None;
        }
        let sectors_in_file = file_len / sector_size;
        let fat_count = read_u32(&header[44..48])? as u64;
        let directory_start = read_u32(&header[48..52])?;
        let mini_cutoff = read_u32(&header[56..60])? as u64;
        let minifat_start = read_u32(&header[60..64])?;
        let minifat_count = read_u32(&header[64..68])? as u64;
        if fat_count > sectors_in_file || minifat_count > sectors_in_file {
            return None;
        }

        // DIFAT: 109 FAT sector numbers in the header, then a chained list.
        let mut difat: Vec<u32> = Vec::new();
        for entry in header[76..512].as_chunks::<4>().0 {
            difat.push(read_u32(entry)?);
        }
        let difat_start = read_u32(&header[68..72])?;
        let difat_count = read_u32(&header[72..76])? as u64;
        if difat_count > sectors_in_file {
            return None;
        }
        let mut next = difat_start;
        for _ in 0..difat_count {
            if next >= CHAIN_END || difat.len() > sectors_in_file as usize {
                return None;
            }
            let sector = read_sector(&mut file, next, sector_size)?;
            let split = sector.len() - 4;
            for entry in sector[..split].as_chunks::<4>().0 {
                difat.push(read_u32(entry)?);
            }
            next = read_u32(&sector[split..])?;
        }

        let mut fat = Vec::new();
        let fat_sectors: Vec<u32> = difat
            .into_iter()
            .filter(|sector| *sector < CHAIN_END)
            .take(fat_count as usize)
            .collect();
        if fat_sectors.len() < fat_count as usize {
            return None;
        }
        for sector in fat_sectors {
            let bytes = read_sector(&mut file, sector, sector_size)?;
            for entry in bytes.as_chunks::<4>().0 {
                fat.push(read_u32(entry)?);
            }
        }

        let mut minifat = Vec::new();
        if minifat_count > 0 {
            let chain = chain_of(minifat_start, &fat)?;
            if chain.len() < minifat_count as usize {
                return None;
            }
            for sector in chain.into_iter().take(minifat_count as usize) {
                let bytes = read_sector(&mut file, sector, sector_size)?;
                for entry in bytes.as_chunks::<4>().0 {
                    minifat.push(read_u32(entry)?);
                }
            }
        }

        let mut compound = Self {
            file,
            sector_size,
            mini_size,
            mini_cutoff,
            fat,
            minifat,
            mini_container: Vec::new(),
            workbook: None,
        };

        // The directory is a chain of 128-byte entries; the root entry (the
        // first of type 5) stores where the mini stream container lives.
        let mut root_start = None;
        let mut visited = 0usize;
        for sector in chain_of(directory_start, &compound.fat)? {
            visited += 1;
            if visited > sectors_in_file as usize + 1 {
                return None;
            }
            let bytes = compound.read_sector(sector)?;
            for entry in bytes.as_chunks::<128>().0 {
                let kind = entry[66];
                if kind == 0 {
                    continue;
                }
                // Name length counts bytes and includes the UTF-16 null.
                let name_len = (read_u16(&entry[64..66])? as usize).min(64) & !1;
                let mut name = decode_utf16le(&entry[..name_len]);
                if name.ends_with('\0') {
                    name.pop();
                }
                let start = read_u32(&entry[116..120])?;
                let size = read_u32(&entry[120..124])? as u64;
                match kind {
                    5 => {
                        if root_start.is_none() {
                            root_start = Some(start);
                        }
                    }
                    2 => {
                        let lower = name.to_lowercase();
                        // Same stream aliases calamine's Xls reader accepts.
                        if compound.workbook.is_none() && (lower == "workbook" || lower == "book") {
                            compound.workbook = Some((start, size));
                        }
                    }
                    _ => {}
                }
            }
        }
        compound.mini_container = chain_of(root_start?, &compound.fat)?;
        Some(compound)
    }

    fn read_sector(&mut self, sector: u32) -> Option<Vec<u8>> {
        read_sector(&mut self.file, sector, self.sector_size)
    }

    fn read_workbook_stream(&mut self) -> Option<Vec<u8>> {
        let (start, size) = self.workbook?;
        if size > MAX_STREAM_BYTES {
            return None;
        }
        let mut bytes = Vec::new();
        if size >= self.mini_cutoff {
            let chain = chain_of(start, &self.fat)?;
            if chain.len() as u64 * self.sector_size < size {
                return None;
            }
            for sector in chain {
                bytes.extend_from_slice(&self.read_sector(sector)?);
                if bytes.len() as u64 >= size {
                    break;
                }
            }
        } else {
            if self.mini_container.is_empty() {
                return None;
            }
            let chain = chain_of(start, &self.minifat)?;
            if chain.len() as u64 * self.mini_size < size {
                return None;
            }
            let per_sector = (self.sector_size / self.mini_size) as usize;
            for mini_sector in chain {
                let container = *self.mini_container.get(mini_sector as usize / per_sector)?;
                let within = (mini_sector as usize % per_sector) as u64 * self.mini_size;
                let offset = 512u64 + container as u64 * self.sector_size + within;
                self.file.seek(SeekFrom::Start(offset)).ok()?;
                let mut mini = vec![0u8; self.mini_size as usize];
                self.file.read_exact(&mut mini).ok()?;
                bytes.extend_from_slice(&mini);
                if bytes.len() as u64 >= size {
                    break;
                }
            }
        }
        bytes.truncate(size as usize);
        Some(bytes)
    }
}

fn read_sector(file: &mut File, sector: u32, sector_size: u64) -> Option<Vec<u8>> {
    let offset = 512u64.checked_add((sector as u64).checked_mul(sector_size)?)?;
    file.seek(SeekFrom::Start(offset)).ok()?;
    let mut bytes = vec![0u8; sector_size as usize];
    file.read_exact(&mut bytes).ok()?;
    Some(bytes)
}

/// Follow a FAT chain with a cycle guard; `None` on a malformed table.
fn chain_of(start: u32, table: &[u32]) -> Option<Vec<u32>> {
    let mut chain = Vec::new();
    let mut current = start;
    loop {
        if current >= CHAIN_END {
            return Some(chain);
        }
        let next = *table.get(current as usize)?;
        chain.push(current);
        if chain.len() > table.len() {
            // A valid chain visits each sector at most once; longer means a
            // hostile cycle.
            return None;
        }
        current = next;
    }
}

// ---------------------------------------------------------------------------
// BIFF record walk — CODEPAGE, SST (+CONTINUE), BOUNDSHEET, LABELSST.
// ---------------------------------------------------------------------------

/// Re-decode SST-backed strings of one BIFF8 workbook stream. `None` means
/// this stream is not a book the overlay applies to (BIFF5, UTF-8/UTF-16
/// codepage) or the walk cannot vouch for it.
fn repair_stream_strings(stream: &[u8]) -> Option<Vec<SheetStrings>> {
    let first = Record::at(stream, 0)?;
    if first.id != REC_BOF || read_u16(first.body) != Some(BIFF8_BOF_VERSION) {
        return None;
    }

    // Workbook globals substream: ends at the first EOF. CODEPAGE decides
    // the decoder (mirrors calamine: last record wins), SST and BOUNDSHEET
    // bodies are collected here and decoded once the codepage is final.
    let mut codepage = DEFAULT_CODEPAGE;
    let mut sst_chunks: Vec<&[u8]> = Vec::new();
    let mut boundsheet_bodies: Vec<&[u8]> = Vec::new();
    let mut position = 0usize;
    while let Some(record) = Record::at(stream, position) {
        match record.id {
            REC_CODEPAGE => {
                if let Some(value) = read_u16(record.body) {
                    codepage = value;
                }
            }
            REC_SST => sst_chunks = record.continuing(stream),
            REC_BOUNDSHEET => boundsheet_bodies.push(record.body),
            REC_EOF => break,
            _ => {}
        }
        position = record.end;
    }

    let encoding = to_encoding(codepage)?;
    if !encoding.is_single_byte() {
        return None;
    }
    let strings = parse_sst(sst_chunks, encoding);

    let mut sheets = Vec::with_capacity(boundsheet_bodies.len());
    for body in boundsheet_bodies {
        let name = parse_short_xl_unicode_string(body.get(6..)?, encoding)?;
        // lbPlyPos: where the sheet substream starts in the stream.
        let offset = read_u32(body)? as usize;
        let mut cells = HashMap::new();
        let mut position = offset;
        let mut guard = stream.len() / 4 + 1;
        while let Some(record) = Record::at(stream, position) {
            guard -= 1;
            if guard == 0 {
                return None;
            }
            match record.id {
                REC_LABELSST => {
                    let row = record.body.get(0..2).and_then(read_u16);
                    let column = record.body.get(2..4).and_then(read_u16);
                    let index = record.body.get(6..10).and_then(read_u32);
                    if let (Some(row), Some(column), Some(index)) = (row, column, index)
                        && let Some(text) =
                            strings.get(index as usize).filter(|text| !text.is_empty())
                    {
                        cells.insert((row as u32, column as u32), text.clone());
                    }
                }
                REC_EOF => break,
                _ => {}
            }
            position = record.end;
        }
        sheets.push(SheetStrings {
            name: Some(name),
            cells,
        });
    }
    Some(sheets)
}

/// A BIFF record view: 4-byte header (u16 id, u16 length) plus the body
/// slice. `end` is the offset of the next record.
struct Record<'a> {
    id: u16,
    body: &'a [u8],
    end: usize,
}

impl<'a> Record<'a> {
    fn at(stream: &'a [u8], position: usize) -> Option<Self> {
        let header_end = position.checked_add(4)?;
        let header = stream.get(position..header_end)?;
        let id = read_u16(&header[..2])?;
        let length = read_u16(&header[2..])? as usize;
        let end = header_end.checked_add(length)?;
        let body = stream.get(header_end..end)?;
        Some(Self { id, body, end })
    }

    /// Body plus every immediately following CONTINUE body — the way a
    /// spanning record (SST) reads as one byte sequence.
    fn continuing(&self, stream: &'a [u8]) -> Vec<&'a [u8]> {
        let mut chunks = vec![self.body];
        let mut position = self.end;
        while let Some(record) = Record::at(stream, position) {
            if record.id != REC_CONTINUE {
                break;
            }
            chunks.push(record.body);
            position = record.end;
        }
        chunks
    }
}

/// Sequential reader over a record chain (body + CONTINUE bodies).
struct ChunkReader<'a> {
    chunks: Vec<&'a [u8]>,
    index: usize,
    position: usize,
}

impl<'a> ChunkReader<'a> {
    fn current(&self) -> Option<&'a [u8]> {
        self.chunks.get(self.index).copied()
    }

    fn exhausted(&self) -> bool {
        self.index >= self.chunks.len()
    }

    /// Copy the next `len` bytes, crossing chunk boundaries as needed.
    /// Callers keep `len` small or chunk-bounded, so the capacity hint is
    /// capped to keep hostile lengths from sizing an allocation.
    fn take(&mut self, len: usize) -> Option<Vec<u8>> {
        let mut out = Vec::with_capacity(len.min(64));
        let mut remaining = len;
        while remaining > 0 {
            let chunk = self.current()?;
            let available = chunk.len().saturating_sub(self.position);
            if available == 0 {
                self.index += 1;
                self.position = 0;
                continue;
            }
            let step = remaining.min(available);
            let chunk = chunk.get(self.position..self.position + step)?;
            out.extend_from_slice(chunk);
            self.position += step;
            remaining -= step;
        }
        Some(out)
    }

    /// Advance past `len` bytes without copying (format runs, ExtRst).
    fn skip(&mut self, len: u64) -> Option<()> {
        let mut remaining = len;
        while remaining > 0 {
            let chunk = self.current()?;
            let available = (chunk.len().saturating_sub(self.position)) as u64;
            if available == 0 {
                self.index += 1;
                self.position = 0;
                continue;
            }
            let step = remaining.min(available) as usize;
            self.position += step;
            remaining -= step as u64;
        }
        Some(())
    }
}

/// Parse the Shared String Table ([MS-XLS] 2.4.265 rgsxb): an 8-byte
/// header, then XLUnicodeRichExtendedString entries. A malformed entry
/// stops the parse with the strings decoded so far — calamine's own open
/// fails on such a book anyway, so partial repair is pure upside.
fn parse_sst(chunks: Vec<&[u8]>, encoding: &'static Encoding) -> Vec<String> {
    let mut reader = ChunkReader {
        chunks,
        index: 0,
        position: 0,
    };
    if reader.take(8).is_none() {
        return Vec::new();
    }
    let mut strings = Vec::new();
    while !reader.exhausted() {
        let Some(text) = parse_sst_string(&mut reader, encoding) else {
            break;
        };
        strings.push(text);
    }
    strings
}

/// One XLUnicodeRichExtendedString ([MS-XLS] 2.5.293): character count,
/// flags (fHighByte / fRichSt / fExtSt), optional run/extension sizes, the
/// characters, then binary blocks that are skipped, not decoded. When the
/// characters continue into a CONTINUE record, that record's first byte
/// re-specifies the compression flag for the rest of the string.
fn parse_sst_string(reader: &mut ChunkReader<'_>, encoding: &'static Encoding) -> Option<String> {
    let header = reader.take(3)?;
    let char_count = read_u16(&header[..2])? as usize;
    let flags = header[2];
    let mut high_byte = flags & 0x1 != 0;
    let run_bytes = if flags & 0x8 != 0 {
        read_u16(&reader.take(2)?)? as u64 * 4
    } else {
        0
    };
    let extension_bytes = if flags & 0x4 != 0 {
        read_u32(&reader.take(4)?)? as u64
    } else {
        0
    };

    let mut text = String::with_capacity(char_count);
    let mut remaining = char_count;
    while remaining > 0 {
        let chunk = reader.current()?;
        let bytes_per_char: usize = if high_byte { 2 } else { 1 };
        let available = chunk.len().saturating_sub(reader.position);
        if available < bytes_per_char {
            // Character data continues in the next CONTINUE record, whose
            // first byte is a fresh compression flag.
            reader.index += 1;
            reader.position = 0;
            high_byte = reader.take(1)?.first().copied()? & 0x1 != 0;
            continue;
        }
        let run = (available / bytes_per_char).min(remaining);
        let bytes = reader.take(run * bytes_per_char)?;
        if high_byte {
            text.push_str(&decode_utf16le(&bytes));
        } else {
            text.push_str(&encoding.decode(&bytes).0);
        }
        remaining -= run;
    }
    reader.skip(run_bytes)?;
    reader.skip(extension_bytes)?;
    Some(text)
}

/// BoundSheet8 sheet name ([MS-XLS] 2.5.292 ShortXLUnicodeString).
fn parse_short_xl_unicode_string(body: &[u8], encoding: &'static Encoding) -> Option<String> {
    let char_count = *body.first()? as usize;
    let flags = *body.get(1)?;
    let bytes_per_char: usize = if flags & 0x1 != 0 { 2 } else { 1 };
    let bytes = body.get(2..2 + char_count.checked_mul(bytes_per_char)?)?;
    Some(if flags & 0x1 != 0 {
        decode_utf16le(bytes)
    } else {
        encoding.decode(bytes).0.into_owned()
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture() -> Vec<u8> {
        std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/bug-1600-utf16-sst-codepage-1252.xls"
        ))
        .unwrap()
    }

    fn biff_record(id: u16, body: &[u8]) -> Vec<u8> {
        let mut bytes = id.to_le_bytes().to_vec();
        bytes.extend_from_slice(&(body.len() as u16).to_le_bytes());
        bytes.extend_from_slice(body);
        bytes
    }

    fn biff_stream(bof_version: u16, codepage: u16, sst_body: &[u8]) -> Vec<u8> {
        let mut bof = vec![0u8; 16];
        bof[..2].copy_from_slice(&bof_version.to_le_bytes());
        bof[2..4].copy_from_slice(&0x0005u16.to_le_bytes());
        let mut stream = biff_record(REC_BOF, &bof);
        stream.extend(biff_record(REC_CODEPAGE, &codepage.to_le_bytes()));
        stream.extend(biff_record(REC_SST, sst_body));
        stream.extend(biff_record(REC_EOF, &[]));
        stream
    }

    /// SST body with a shared 8-byte header plus one entry per argument.
    fn sst_body(entries: &[Vec<u8>]) -> Vec<u8> {
        let mut body = vec![0u8; 8];
        for entry in entries {
            body.extend_from_slice(entry);
        }
        body
    }

    #[test]
    fn fixture_overlay_carries_repaired_strings() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("book.xls");
        std::fs::write(&path, fixture()).unwrap();

        let overlay = LegacyXlsStrings::extract(&path);
        let sheet = overlay.sheet(0).expect("overlay for the only sheet");
        assert_eq!(sheet.name.as_deref(), Some("Прайс"));
        assert_eq!(
            sheet.cells.get(&(0, 0)).map(String::as_str),
            Some("Наименование")
        );
        assert_eq!(
            sheet.cells.get(&(1, 0)).map(String::as_str),
            Some("Первая строка\nВторая строка")
        );
        assert_eq!(
            sheet.cells.get(&(2, 0)).map(String::as_str),
            Some("Plain ASCII item")
        );
        assert_eq!(
            sheet.cells.get(&(3, 0)).map(String::as_str),
            Some("Кириллица и ASCII mix")
        );
        assert_eq!(
            sheet.cells.get(&(4, 0)).map(String::as_str),
            Some("Вторая книга")
        );
    }

    #[test]
    fn continuation_record_switches_compression_mid_string() {
        // cch=3 fHighByte=1: 'A' fits, one stray byte stays behind; the
        // CONTINUE re-specifies fHighByte=0 for the remaining chars.
        let encoding = to_encoding(1252).unwrap();
        let first = [0u8; 8]
            .into_iter()
            .chain([0x03, 0x00, 0x01, b'A', 0x00, b'B'])
            .collect::<Vec<u8>>();
        let second = vec![0x00, b'B', b'C'];
        let chunks: Vec<&[u8]> = vec![first.as_slice(), second.as_slice()];
        assert_eq!(parse_sst(chunks, encoding), vec!["ABC".to_string()]);
    }

    #[test]
    fn rich_and_ext_blocks_are_skipped_without_decoding() {
        let encoding = to_encoding(1252).unwrap();
        let mut entry = vec![0x02, 0x00, 0x0D, 0x01, 0x00, 0x04, 0x00, 0x00, 0x00];
        entry.extend_from_slice("A\0B\0".as_bytes());
        entry.extend_from_slice(&[0u8; 4]); // one FormatRun
        entry.extend_from_slice(&[0u8; 4]); // ExtRst
        let entry2 = vec![0x01, 0x00, 0x00, b'Z'];
        let body = sst_body(&[entry, entry2]);
        assert_eq!(
            parse_sst(vec![&body], encoding),
            vec!["AB".to_string(), "Z".to_string()]
        );
    }

    #[test]
    fn truncated_sst_keeps_the_entries_parsed_so_far() {
        let encoding = to_encoding(1252).unwrap();
        let entry = vec![0x01, 0x00, 0x00, b'K'];
        let truncated = vec![0x05, 0x00, 0x01, b'A', 0x00, b'B'];
        let body = sst_body(&[entry, truncated]);
        assert_eq!(parse_sst(vec![&body], encoding), vec!["K".to_string()]);
    }

    #[test]
    fn biff5_books_are_left_to_calamine() {
        // BIFF5 has no SST and its BoundSheet names carry no flags byte;
        // the walk must refuse the stream instead of misparsing names.
        let stream = biff_stream(0x0500, 1252, &sst_body(&[vec![0x01, 0x00, 0x00, b'x']]));
        assert!(repair_stream_strings(&stream).is_none());
    }

    #[test]
    fn utf16_codepage_books_need_no_overlay() {
        // calamine decodes codepage-1200 books correctly today; the overlay
        // must stay empty so those conversions stay byte-for-byte identical.
        let stream = biff_stream(
            BIFF8_BOF_VERSION,
            1200,
            &sst_body(&[vec![0x01, 0x00, 0x01, b'x', 0x00]]),
        );
        assert!(repair_stream_strings(&stream).is_none());
    }

    #[test]
    fn cyclic_fat_chain_does_not_hang_the_reader() {
        let mut bytes = vec![0u8; 512 + 3 * 512];
        bytes[..8].copy_from_slice(&CFB_MAGIC);
        bytes[30..32].copy_from_slice(&9u16.to_le_bytes()); // 512-byte sectors
        bytes[32..34].copy_from_slice(&6u16.to_le_bytes()); // 64-byte mini sectors
        bytes[44..48].copy_from_slice(&1u32.to_le_bytes()); // one FAT sector
        bytes[48..52].copy_from_slice(&1u32.to_le_bytes()); // directory at sector 1
        bytes[56..60].copy_from_slice(&4096u32.to_le_bytes()); // mini cutoff
        bytes[60..64].copy_from_slice(&0xFFFFFFFEu32.to_le_bytes()); // no mini FAT
        bytes[64..68].copy_from_slice(&0u32.to_le_bytes());
        bytes[68..72].copy_from_slice(&0xFFFFFFFEu32.to_le_bytes()); // DIFAT in header
        bytes[72..76].copy_from_slice(&0u32.to_le_bytes());
        bytes[76..80].copy_from_slice(&0u32.to_le_bytes()); // FAT at sector 0
        // FAT with a 1 <-> 2 cycle in the directory chain.
        bytes[512..516].copy_from_slice(&0xFFFFFFFDu32.to_le_bytes());
        bytes[516..520].copy_from_slice(&2u32.to_le_bytes());
        bytes[520..524].copy_from_slice(&1u32.to_le_bytes());

        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("cycle.xls");
        std::fs::write(&path, &bytes).unwrap();

        let overlay = LegacyXlsStrings::extract(&path);
        assert!(overlay.sheet(0).is_none());
    }

    #[test]
    fn truncated_fixtures_never_panic_the_reader() {
        let dir = tempfile::tempdir().unwrap();
        let full = fixture();
        for cut in (0..full.len()).step_by(64) {
            let path = dir.path().join("cut.xls");
            std::fs::write(&path, &full[..cut]).unwrap();
            let _ = LegacyXlsStrings::extract(&path);
        }
    }
}

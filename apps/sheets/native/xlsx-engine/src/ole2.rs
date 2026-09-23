//! OLE2 header pre-validation for the legacy convert path (BUG-1606).
//!
//! calamine 0.36 panics — a slice-index failure at `cfb.rs:306` in
//! `Sectors::get`, and the FAT chain walk indexing `fats[sector_id]` right
//! next to it — on some truncated or corrupt OLE2 (compound file) headers,
//! before it can return a normal error: a sector id the header names is
//! checked against the file length while the data backing it was never
//! read. A panic on a user file kills the whole sidecar process, so the
//! convert path refuses the clearly-truncated header shapes here, before
//! calamine opens the file. `convert.rs` additionally wraps its calamine
//! calls in `catch_unwind` as the safety net for malformed structures this
//! header-level fence cannot see (directory chains, mini streams).
//!
//! The fence is deliberately conservative: it only rejects geometry no
//! valid [MS-CFB] file can have — a declared sector (directory start, FAT
//! sectors named by the header DIFAT, mini FAT, DIFAT chain) that lies past
//! the end of the file, FAT/mini-FAT counts larger than the file could
//! hold, no FAT sector named anywhere, and an all-empty first FAT sector
//! (one bounded sector read; every zero entry chains all sectors back to
//! sector 0, which calamine walks into unbounded allocation). Borderline
//! shapes pass through to calamine, whose own open path reports them with
//! normal errors.

use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;

use crate::SidecarError;
use crate::legacy_xls::CFB_MAGIC;

/// Sector ids at or above this are sentinels (ENDOFCHAIN, FREESECT,
/// DIFSECT, FATSECT — [MS-CFB] 2.8), not real sectors.
const RESERVED_SECTORS: u32 = 0xFFFF_FFFA;
/// OLE2 header size; the smallest file a header check applies to.
const HEADER_BYTES: usize = 512;
/// Sector ids are u32 fields read from these fixed header offsets
/// ([MS-CFB] 2.2).
const FAT_SECTOR_COUNT: std::ops::Range<usize> = 44..48;
const FIRST_DIR_SECTOR: std::ops::Range<usize> = 48..52;
const FIRST_MINI_FAT_SECTOR: std::ops::Range<usize> = 60..64;
const MINI_FAT_SECTOR_COUNT: std::ops::Range<usize> = 64..68;
const FIRST_DIFAT_SECTOR: std::ops::Range<usize> = 68..72;
/// First DIFAT entry; the header array holds 109 u32 sector ids. Entries
/// naming real sectors must exist in the file.
const HEADER_DIFAT: std::ops::Range<usize> = 76..512;

/// Refuses OLE2 compound documents whose header promises sectors the file
/// does not contain — the truncated-header class calamine 0.36 panics on.
/// Non-OLE2 sources (zip workbooks, ODS, anything without the compound
/// magic) and files too short to carry a full header pass through: both
/// fail inside calamine with a normal error, exactly as before the fence.
pub(crate) fn validate_ole2_header(source: &Path) -> Result<(), SidecarError> {
    let mut file = File::open(source)?;
    let file_len = file.metadata()?.len();
    if file_len < HEADER_BYTES as u64 {
        // No full header to misread: calamine reports short files with a
        // normal error on its own.
        return Ok(());
    }
    let mut header = [0u8; HEADER_BYTES];
    file.read_exact(&mut header)?;
    if header[..CFB_MAGIC.len()] != CFB_MAGIC {
        // The fence only understands compound documents; every other
        // format stays with calamine's own open path.
        return Ok(());
    }
    let sector_size = match le_u16(&header[30..32]) {
        0x0009 => 512usize,
        0x000C => 4096,
        // Any other shift is refused by calamine itself with a normal
        // error; the fence has no opinion.
        _ => return Ok(()),
    };
    // Version 4 files carry a 4096-byte header region; sectors start past
    // it in both versions ([MS-CFB] 2.2, 3.2).
    let prefix = if sector_size == 4096 {
        4096
    } else {
        HEADER_BYTES
    };
    let Some(sector_count) = (file_len as usize)
        .checked_sub(prefix)
        .map(|data_bytes| data_bytes / sector_size)
    else {
        return refuse("leaves no room for a single sector".into());
    };

    let fat_sectors = le_u32(&header[FAT_SECTOR_COUNT]);
    let dir_start = le_u32(&header[FIRST_DIR_SECTOR]);
    let mini_fat_start = le_u32(&header[FIRST_MINI_FAT_SECTOR]);
    let mini_fat_sectors = le_u32(&header[MINI_FAT_SECTOR_COUNT]);
    let difat_start = le_u32(&header[FIRST_DIFAT_SECTOR]);

    if fat_sectors == 0 {
        // Every valid compound file chains its sectors through a FAT; with
        // none declared calamine builds an empty table and indexes it.
        return refuse("declares no FAT sectors".into());
    }
    if fat_sectors as usize > sector_count {
        // FAT sectors live in the data area, so there cannot be more of
        // them than the file holds — a count past the file end is the
        // truncation signature (calamine would short-read the FATs and
        // later index the short table by a sector id it still considers
        // valid).
        return refuse(format!(
            "declares {fat_sectors} FAT sectors but the file holds only {sector_count}"
        ));
    }
    if mini_fat_sectors as usize > sector_count {
        return refuse(format!(
            "declares {mini_fat_sectors} mini FAT sectors but the file holds only {sector_count}"
        ));
    }
    for entry in header[HEADER_DIFAT].chunks_exact(4) {
        let id = le_u32(entry);
        if id < RESERVED_SECTORS && (id as usize) >= sector_count {
            return refuse(format!(
                "names FAT sector {id} beyond the end of the file ({sector_count} sectors present)"
            ));
        }
    }
    if dir_start < RESERVED_SECTORS && (dir_start as usize) >= sector_count {
        return refuse(format!(
            "starts the directory at sector {dir_start}, beyond the end of the file ({sector_count} sectors present)"
        ));
    }
    if mini_fat_sectors > 0
        && mini_fat_start < RESERVED_SECTORS
        && (mini_fat_start as usize) >= sector_count
    {
        return refuse(format!(
            "starts the mini FAT at sector {mini_fat_start}, beyond the end of the file ({sector_count} sectors present)"
        ));
    }
    if difat_start < RESERVED_SECTORS && (difat_start as usize) >= sector_count {
        return refuse(format!(
            "chains the DIFAT at sector {difat_start}, beyond the end of the file ({sector_count} sectors present)"
        ));
    }
    // The header DIFAT is the only place a single-DIFAT file can name its
    // FAT sectors. A FAT count above zero with nothing naming any FAT
    // sector leaves calamine an empty table to index — the panic class
    // behind BUG-1606 — so that combination is refused outright.
    let first_fat_sector = header[HEADER_DIFAT]
        .chunks_exact(4)
        .map(le_u32)
        .find(|id| *id < RESERVED_SECTORS);
    if first_fat_sector.is_none() && difat_start >= RESERVED_SECTORS {
        return refuse("declares FAT sectors but no DIFAT entry names one".into());
    }
    // calamine never bounds its FAT chain walk: a FAT whose every entry is
    // zero chains all sectors back to sector 0, and it allocates without
    // limit walking that cycle (found while testing this fence; an abort
    // the catch_unwind net in convert.rs cannot contain). Reading the one
    // sector the header DIFAT names is bounded and settles the case.
    if let Some(fat_sector) = first_fat_sector {
        file.seek(SeekFrom::Start(
            prefix as u64 + u64::from(fat_sector) * sector_size as u64,
        ))?;
        let mut fat = vec![0u8; sector_size];
        file.read_exact(&mut fat)?;
        if fat.chunks_exact(4).all(|entry| le_u32(entry) == 0) {
            return refuse("carries an all-empty FAT: no sector chain can be walked".into());
        }
    }
    Ok(())
}

/// One refusal text for every truncated-header shape, in the same
/// "Unable to read the workbook" form the rest of the convert path uses.
fn refuse(detail: String) -> Result<(), SidecarError> {
    Err(SidecarError::Workbook(format!(
        "Unable to read the workbook: the OLE2 header {detail}. The file is truncated or corrupt."
    )))
}

fn le_u16(bytes: &[u8]) -> u16 {
    u16::from_le_bytes(bytes.try_into().expect("fixed-width header field"))
}

fn le_u32(bytes: &[u8]) -> u32 {
    u32::from_le_bytes(bytes.try_into().expect("fixed-width header field"))
}

#[cfg(test)]
pub(crate) mod test_fixtures {
    //! Programmatic compound-file fixtures (no committed binaries): a
    //! spec-valid minimal OLE2 header over a chosen sector layout, plus the
    //! mutators the regression tests use to forge each truncated shape.

    use super::{HEADER_BYTES, RESERVED_SECTORS};
    use crate::legacy_xls::CFB_MAGIC;

    /// Header for a file holding `data_sectors` 512-byte sectors past the
    /// header, one FAT sector (sector 0), directory at sector 1, no mini
    /// FAT, no DIFAT chain.
    pub(crate) fn minimal_header(data_sectors: usize) -> Vec<u8> {
        let mut header = vec![0u8; HEADER_BYTES];
        header[..8].copy_from_slice(&CFB_MAGIC);
        header[26..28].copy_from_slice(&3u16.to_le_bytes()); // major version
        header[30..32].copy_from_slice(&9u16.to_le_bytes()); // sector shift
        header[32..34].copy_from_slice(&6u16.to_le_bytes()); // mini sector shift
        header[44..48].copy_from_slice(&1u32.to_le_bytes()); // one FAT sector
        header[48..52].copy_from_slice(&1u32.to_le_bytes()); // directory at 1
        header[60..64].copy_from_slice(&0xFFFF_FFFEu32.to_le_bytes()); // no mini FAT
        header[64..68].copy_from_slice(&0u32.to_le_bytes());
        header[68..72].copy_from_slice(&0xFFFF_FFFEu32.to_le_bytes()); // no DIFAT chain
        header[72..76].copy_from_slice(&0u32.to_le_bytes());
        header[76..80].copy_from_slice(&0u32.to_le_bytes()); // FAT lives at 0
        for entry in header[80..HEADER_BYTES].chunks_exact_mut(4) {
            entry.copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // FREESECT
        }
        debug_assert!(data_sectors >= 2, "FAT + directory must fit");
        header
    }

    /// Header over `data_sectors` zeroed sectors. Its all-zero FAT content
    /// is a corrupt shape the fence refuses, so this fixture only feeds
    /// refusal-shape tests that end before calamine or the FAT read —
    /// anything meant to pass the fence must use `valid_cfb` instead.
    pub(crate) fn zeros(data_sectors: usize) -> Vec<u8> {
        let mut bytes = minimal_header(data_sectors);
        bytes.extend(std::iter::repeat_n(0u8, data_sectors * 512));
        bytes
    }

    /// Free-sector sentinel for forging DIFAT tails.
    pub(crate) const FREESECT: u32 = 0xFFFF_FFFF;

    /// The sentinel ceiling the fence treats as "not a sector id".
    pub(crate) const RESERVED: u32 = RESERVED_SECTORS;

    /// One 128-byte directory entry: name (UTF-16LE) + type/color + start
    /// sector + stream length. calamine scans entries flat, so sibling and
    /// child pointers stay empty.
    fn directory_entry(name: &str, kind: u8, start: u32, len: u64) -> Vec<u8> {
        let mut block = vec![0u8; 128];
        let encoded: Vec<u8> = name
            .encode_utf16()
            .flat_map(|unit| unit.to_le_bytes())
            .collect();
        block[..encoded.len()].copy_from_slice(&encoded);
        block[64..66].copy_from_slice(&((encoded.len() + 2) as u16).to_le_bytes());
        block[66] = kind; // 5 = root, 2 = stream
        block[67] = 1; // color: black
        block[116..120].copy_from_slice(&start.to_le_bytes());
        block[120..128].copy_from_slice(&len.to_le_bytes());
        block
    }

    /// A structurally valid, content-free compound file: the minimal
    /// header, one FAT sector marking itself and ending the directory
    /// chain, and one directory sector holding only a Root Entry (plus
    /// `extra_sectors` free sectors at the end). The fence accepts it and
    /// calamine opens it, reporting the missing Workbook stream as a
    /// normal error.
    pub(crate) fn valid_cfb(extra_sectors: usize) -> Vec<u8> {
        const END: u32 = 0xFFFF_FFFE; // ENDOFCHAIN
        const FATSECT: u32 = 0xFFFF_FFFD;
        let mut fat_sector = vec![0u8; 512];
        for (index, entry) in fat_sector.chunks_exact_mut(4).enumerate() {
            let value = match index {
                0 => FATSECT,
                1 => END, // directory: single sector
                _ => FREESECT,
            };
            entry.copy_from_slice(&value.to_le_bytes());
        }
        let mut bytes = minimal_header(2 + extra_sectors);
        bytes.extend_from_slice(&fat_sector);
        // Entries live inside one full 512-byte directory sector.
        let mut directory = directory_entry("Root Entry", 5, END, 0);
        directory.resize(512, 0);
        bytes.extend_from_slice(&directory);
        bytes.extend(std::iter::repeat_n(0u8, extra_sectors * 512));
        bytes
    }

    /// A complete, valid compound file carrying a minimal BIFF8 workbook —
    /// header, one FAT sector, one directory sector, and a "Workbook"
    /// stream of 4096 bytes (large enough that calamine reads it through
    /// the regular sector chain, not the mini stream) holding the
    /// smallest record set calamine accepts: globals BOF, one BOUNDSHEET
    /// naming sheet "S", the globals EOF, and that sheet's BOF + EOF.
    /// Enough for a full conversion with zero cells.
    pub(crate) fn minimal_biff8_cfb() -> Vec<u8> {
        const END: u32 = 0xFFFF_FFFE; // ENDOFCHAIN
        // Sectors: 0 = FAT, 1 = directory, 2..=9 = 4096-byte Workbook stream.
        const STREAM_SECTORS: usize = 8;
        const LAST: u32 = (1 + STREAM_SECTORS) as u32; // = 9

        let mut fat = vec![0u32; 128];
        fat[0] = 0xFFFF_FFFD; // FATSECT: this sector is the FAT
        fat[1] = END; // directory: single sector
        for sector in 2..LAST {
            fat[sector as usize] = sector + 1; // stream chain
        }
        fat[LAST as usize] = END;

        // Directory entries are 128 bytes; calamine scans them flat.
        let mut directory = directory_entry("Root Entry", 5, END, 0);
        directory.extend(directory_entry(
            "Workbook",
            2,
            2,
            (STREAM_SECTORS * 512) as u64,
        ));

        // BIFF8 records: [id: u16][len: u16][payload].
        fn record(id: u16, payload: &[u8]) -> Vec<u8> {
            let mut bytes = id.to_le_bytes().to_vec();
            bytes.extend((payload.len() as u16).to_le_bytes());
            bytes.extend_from_slice(payload);
            bytes
        }
        // BOF payload: version (BIFF8), substream kind, then build/year and
        // history fields the spec sizes at 16 bytes total.
        fn bof(kind: u16) -> Vec<u8> {
            let mut payload = 0x0600u16.to_le_bytes().to_vec();
            payload.extend(kind.to_le_bytes());
            payload.extend_from_slice(&[0u8; 12]);
            record(0x0809, &payload)
        }
        let globals_bof = bof(0x0005); // workbook globals
        let sheet_bof = bof(0x0010); // worksheet
        let eof = record(0x000A, &[]);
        // BOUNDSHEET: sheet position (offset of the sheet BOF within the
        // Workbook stream), visible, one 8-bit character, name "S".
        let mut payload = 0u32.to_le_bytes().to_vec(); // lbPlyPos, patched below
        payload.extend(0u16.to_le_bytes()); // grbit: visible
        payload.push(1); // cch
        payload.push(0); // fHighByte=0: compressed 8-bit name
        payload.push(b'S');
        let mut boundsheet = record(0x0085, &payload);
        let sheet_offset = (globals_bof.len() + boundsheet.len() + eof.len()) as u32;
        boundsheet[4..8].copy_from_slice(&sheet_offset.to_le_bytes());

        // Globals substream: BOF, the sheet list, EOF. The BOUNDSHEET must
        // sit before the globals EOF — the reader stops there — and its
        // lbPlyPos points at the sheet substream's BOF right after it.
        let mut stream = globals_bof;
        stream.extend(&boundsheet);
        stream.extend(&eof);
        stream.extend(&sheet_bof);
        stream.extend(&eof);
        stream.extend(std::iter::repeat_n(
            0u8,
            STREAM_SECTORS * 512 - stream.len(),
        ));

        let mut bytes = minimal_header(1 + STREAM_SECTORS);
        for word in fat {
            bytes.extend_from_slice(&word.to_le_bytes());
        }
        // The directory occupies one full sector so the Workbook stream
        // starts exactly at its named sector boundary.
        directory.resize(512, 0);
        bytes.extend_from_slice(&directory);
        bytes.extend_from_slice(&stream);
        bytes
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use test_fixtures::{valid_cfb, zeros};

    fn validate_bytes(bytes: &[u8]) -> Result<(), SidecarError> {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("book.xls");
        std::fs::write(&path, bytes).unwrap();
        validate_ole2_header(&path)
    }

    #[test]
    fn accepts_a_minimal_valid_header() {
        // Two data sectors — FAT + directory — the smallest valid layout.
        assert_eq!(valid_cfb(0).len(), 3 * 512);
        validate_bytes(&valid_cfb(0)).unwrap();
    }

    #[test]
    fn accepts_larger_valid_headers() {
        validate_bytes(&valid_cfb(62)).unwrap();
    }

    #[test]
    fn passes_through_files_without_the_compound_magic() {
        let mut bytes = b"PK\x03\x04".to_vec();
        bytes.extend(std::iter::repeat_n(0u8, 1024));
        validate_bytes(&bytes).unwrap();
    }

    #[test]
    fn passes_through_files_shorter_than_a_header() {
        // The existing convert-path behavior for short junk: calamine
        // reports it, the fence stays out of the way.
        let mut bytes = CFB_MAGIC.to_vec();
        bytes.extend(std::iter::repeat_n(0u8, 64));
        validate_bytes(&bytes).unwrap();
    }

    #[test]
    fn passes_through_an_unknown_sector_shift() {
        let mut bytes = valid_cfb(0);
        bytes[30..32].copy_from_slice(&7u16.to_le_bytes());
        validate_bytes(&bytes).unwrap();
    }

    #[test]
    fn refuses_a_fat_sector_named_beyond_eof() {
        // The reported panic shape: the only FAT sector the header DIFAT
        // names does not exist in the truncated file.
        let mut bytes = zeros(2);
        bytes[76..80].copy_from_slice(&9u32.to_le_bytes());
        let error = validate_bytes(&bytes).unwrap_err().to_string();
        assert!(
            error.contains("names FAT sector 9 beyond the end"),
            "{error}"
        );
        assert!(error.contains("truncated or corrupt"), "{error}");
    }

    #[test]
    fn refuses_a_directory_sector_beyond_eof() {
        let mut bytes = zeros(2);
        bytes[48..52].copy_from_slice(&5u32.to_le_bytes());
        let error = validate_bytes(&bytes).unwrap_err().to_string();
        assert!(
            error.contains("starts the directory at sector 5"),
            "{error}"
        );
    }

    #[test]
    fn refuses_a_fat_count_larger_than_the_file() {
        let mut bytes = zeros(2);
        bytes[44..48].copy_from_slice(&7u32.to_le_bytes());
        let error = validate_bytes(&bytes).unwrap_err().to_string();
        assert!(
            error.contains("declares 7 FAT sectors but the file holds only 2"),
            "{error}"
        );
    }

    #[test]
    fn refuses_zero_fat_sectors() {
        let mut bytes = zeros(2);
        bytes[44..48].copy_from_slice(&0u32.to_le_bytes());
        let error = validate_bytes(&bytes).unwrap_err().to_string();
        assert!(error.contains("declares no FAT sectors"), "{error}");
    }

    #[test]
    fn refuses_a_mini_fat_sector_beyond_eof() {
        let mut bytes = zeros(2);
        bytes[60..64].copy_from_slice(&8u32.to_le_bytes()); // start
        bytes[64..68].copy_from_slice(&1u32.to_le_bytes()); // count > 0
        let error = validate_bytes(&bytes).unwrap_err().to_string();
        assert!(error.contains("starts the mini FAT at sector 8"), "{error}");
    }

    #[test]
    fn refuses_a_mini_fat_count_larger_than_the_file() {
        let mut bytes = zeros(2);
        bytes[64..68].copy_from_slice(&50u32.to_le_bytes());
        let error = validate_bytes(&bytes).unwrap_err().to_string();
        assert!(
            error.contains("declares 50 mini FAT sectors but the file holds only 2"),
            "{error}"
        );
    }

    #[test]
    fn refuses_a_difat_chain_sector_beyond_eof() {
        let mut bytes = zeros(2);
        bytes[68..72].copy_from_slice(&6u32.to_le_bytes());
        let error = validate_bytes(&bytes).unwrap_err().to_string();
        assert!(error.contains("chains the DIFAT at sector 6"), "{error}");
    }

    #[test]
    fn refuses_a_header_that_names_no_fat_anywhere() {
        // A FAT count above zero with every header DIFAT entry a sentinel
        // and no DIFAT chain leaves no FAT sector anywhere — calamine
        // builds an empty table and indexes it (the BUG-1606 class).
        let mut bytes = valid_cfb(0);
        bytes[76..80].copy_from_slice(&test_fixtures::FREESECT.to_le_bytes());
        let error = validate_bytes(&bytes).unwrap_err().to_string();
        assert!(
            error.contains("declares FAT sectors but no DIFAT entry names one"),
            "{error}"
        );
    }

    #[test]
    fn refuses_an_all_empty_fat_sector() {
        // A FAT whose every entry is zero chains all sectors back to
        // sector 0; calamine walks that cycle into unbounded allocation.
        // The fence reads the one named FAT sector and refuses first.
        let mut bytes = valid_cfb(0);
        bytes[HEADER_BYTES..2 * HEADER_BYTES].fill(0);
        let error = validate_bytes(&bytes).unwrap_err().to_string();
        assert!(error.contains("all-empty FAT"), "{error}");
    }

    #[test]
    fn tolerates_sentinel_sector_ids() {
        // ENDOFCHAIN/FREESECT-class ids mean "absent", not out of bounds:
        // the forged all-FREESECT DIFAT tail below must pass like the real one.
        let mut bytes = valid_cfb(0);
        for entry in bytes[80..HEADER_BYTES].chunks_exact_mut(4) {
            entry.copy_from_slice(&test_fixtures::FREESECT.to_le_bytes());
        }
        validate_bytes(&bytes).unwrap();
        // A sentinel in the directory-start field passes through too —
        // calamine answers it with its own empty-root error.
        let mut absent_dir = valid_cfb(0);
        absent_dir[48..52].copy_from_slice(&test_fixtures::RESERVED.to_le_bytes());
        validate_bytes(&absent_dir).unwrap();
    }
}

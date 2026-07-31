//! Bounded ZIP access for OOXML packages.
//!
//! This module owns archive-specific reading and metadata inspection. Persistent
//! policy, accounting, usage snapshots, and poison state live in `resource`.

use crate::resource::{
    self, OoxmlFormat, ResourceGovernor, ResourceScope, HARD_MAX_ARCHIVE_ENTRY_BYTES,
};

/// Cap eager allocation based on attacker-controlled ZIP declarations. Large
/// legitimate entries grow incrementally while reads remain resource-bounded.
const INITIAL_RESERVE_CAP: usize = 1024 * 1024;
const READ_CHUNK_BYTES: usize = 32 * 1024;
const EOCD_SIGNATURE: u32 = 0x0605_4b50;
const ZIP64_LOCATOR_SIGNATURE: u32 = 0x0706_4b50;
const ZIP64_EOCD_SIGNATURE: u32 = 0x0606_4b50;

fn initial_reserve(declared_size: u64, read_limit: u64) -> usize {
    declared_size
        .min(read_limit)
        .min(INITIAL_RESERVE_CAP as u64) as usize
}

fn le_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(
        bytes.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

fn le_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        bytes.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

fn le_u64(bytes: &[u8], offset: usize) -> Option<u64> {
    Some(u64::from_le_bytes(
        bytes.get(offset..offset + 8)?.try_into().ok()?,
    ))
}

fn raw_entry_count_from_eocd(data: &[u8], eocd: usize) -> Option<u64> {
    let comment_len = le_u16(data, eocd + 20)?;
    if eocd.checked_add(22 + comment_len as usize)? != data.len() {
        return None;
    }
    let (disk, cd_disk, entries_disk, entries_total) = (
        le_u16(data, eocd + 4)?,
        le_u16(data, eocd + 6)?,
        le_u16(data, eocd + 8)?,
        le_u16(data, eocd + 10)?,
    );
    if disk != 0 || cd_disk != 0 || entries_disk != entries_total {
        return None;
    }

    let (raw_count, central_size, central_offset, directory_boundary) = if entries_total != u16::MAX
    {
        (
            entries_total as u64,
            le_u32(data, eocd + 12)? as u64,
            le_u32(data, eocd + 16)? as u64,
            eocd as u64,
        )
    } else {
        if eocd < 20 || le_u32(data, eocd - 20)? != ZIP64_LOCATOR_SIGNATURE {
            return None;
        }
        let zip64 = usize::try_from(le_u64(data, eocd - 12)?).ok()?;
        if le_u32(data, zip64)? != ZIP64_EOCD_SIGNATURE {
            return None;
        }
        let record_size = le_u64(data, zip64 + 4)?;
        let zip64_end = (zip64 as u64).checked_add(12)?.checked_add(record_size)?;
        if record_size < 44
            || zip64_end != (eocd - 20) as u64
            || le_u32(data, zip64 + 16)? != 0
            || le_u32(data, zip64 + 20)? != 0
        {
            return None;
        }
        let entries_disk = le_u64(data, zip64 + 24)?;
        let entries_total = le_u64(data, zip64 + 32)?;
        if entries_disk != entries_total {
            return None;
        }
        (
            entries_total,
            le_u64(data, zip64 + 40)?,
            le_u64(data, zip64 + 48)?,
            zip64 as u64,
        )
    };

    // Do not treat an EOCD-looking byte sequence inside an arbitrary ZIP
    // comment as proven metadata. Every central record has a 46-byte fixed
    // header, and a non-empty directory starts with its own signature.
    let minimum_central_size = raw_count.checked_mul(46)?;
    let central_end = central_offset.checked_add(central_size)?;
    if central_size < minimum_central_size
        || central_end > directory_boundary
        || (raw_count > 0
            && usize::try_from(central_offset)
                .ok()
                .and_then(|offset| le_u32(data, offset))
                != Some(0x0201_4b50))
    {
        return None;
    }
    Some(raw_count)
}

/// Inspect the legal EOCD tail before `zip::ZipArchive::new` allocates its
/// filename index. A well-formed raw record count is a proven hard-quota signal;
/// malformed/truncated metadata is left to the ZIP crate's ordinary corruption
/// path rather than being mislabeled as a resource violation.
pub fn preflight_archive_limits(data: &[u8]) -> Result<(), String> {
    const MAX_EOCD_SEARCH: usize = 22 + u16::MAX as usize;

    if data.len() < 22 {
        return Ok(());
    }
    let start = data.len().saturating_sub(MAX_EOCD_SEARCH);
    for eocd in (start..=data.len() - 4).rev() {
        if le_u32(data, eocd) != Some(EOCD_SIGNATURE) {
            continue;
        }
        if let Some(raw_count) = raw_entry_count_from_eocd(data, eocd) {
            return resource::observe_archive_metadata(raw_count, 0);
        }
    }
    Ok(())
}

/// Create an ephemeral governor for one free-function operation.
pub fn scoped_limits(
    format: OoxmlFormat,
    operation: &str,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> ResourceScope {
    ResourceGovernor::from_wasm(format, max_archive_entry_bytes, max_total_inflated_bytes)
        .scope(operation)
}

/// Install a retained archive's persistent governor for one synchronous method.
pub fn scoped_governor(governor: &ResourceGovernor, operation: &str) -> ResourceScope {
    governor.scope(operation)
}

/// Record accessible central-directory metadata and enforce hard entry count.
/// Declared bytes are diagnostic/early-entry evidence, never treated as proof of
/// actual total inflation.
pub fn validate_archive_limits<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Result<(), String> {
    resource::assert_healthy()?;
    let mut declared_total = 0u64;
    for index in 0..archive.len() {
        let entry = archive
            .by_index(index)
            .map_err(|e| format!("ZIP archive entry metadata error: {e}"))?;
        declared_total = declared_total.saturating_add(entry.size());
    }
    resource::observe_archive_metadata(archive.len() as u64, declared_total)
}

fn read_entry<R: std::io::Read>(
    entry: &mut R,
    part_id: usize,
    path: &str,
    declared_size: u64,
    requested_prefix: Option<u64>,
) -> Result<Vec<u8>, String> {
    resource::assert_healthy()?;
    let resource_allowance = resource::read_allowance(part_id, path, declared_size)?;
    let caller_limit = requested_prefix.unwrap_or(u64::MAX);
    let ordinary_limit = resource_allowance.min(caller_limit);
    let resource_bound_is_tighter = resource_allowance < caller_limit;
    let mut bytes = Vec::with_capacity(initial_reserve(declared_size, ordinary_limit));
    let mut observed = 0u64;
    let mut chunk = [0u8; READ_CHUNK_BYTES];

    loop {
        if observed == ordinary_limit {
            if !resource_bound_is_tighter {
                break;
            }
            // Read exactly one more byte to distinguish EOF-at-limit from a
            // proven policy crossing. Recording it latches limit+1.
            let count = entry
                .read(&mut chunk[..1])
                .map_err(|e| format!("read error: {e}"))?;
            if count == 0 {
                break;
            }
            observed = observed.saturating_add(count as u64);
            resource::observe_inflated(part_id, path, observed, count as u64)?;
            unreachable!("resource allowance + 1 must reject");
        }

        let remaining = ordinary_limit - observed;
        let count = entry
            .read(&mut chunk[..remaining.min(READ_CHUNK_BYTES as u64) as usize])
            .map_err(|e| format!("read error: {e}"))?;
        if count == 0 {
            break;
        }
        observed = observed.saturating_add(count as u64);
        // Charge each successful decompressor delivery immediately so bytes
        // emitted before a later CRC/read failure remain accounted.
        resource::observe_inflated(part_id, path, observed, count as u64)?;
        bytes.extend_from_slice(&chunk[..count]);
    }
    Ok(bytes)
}

/// Open one package and extract a single entry under an ephemeral resource
/// session. Stateful browser paths use their retained archive governor instead.
pub fn extract_zip_entry(
    data: &[u8],
    path: &str,
    format: OoxmlFormat,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<Vec<u8>, String> {
    use std::io::Cursor;
    let _scope = scoped_limits(
        format,
        "extract",
        max_archive_entry_bytes,
        max_total_inflated_bytes,
    );
    preflight_archive_limits(data)?;
    let mut archive =
        zip::ZipArchive::new(Cursor::new(data)).map_err(|e| format!("zip open error: {e}"))?;
    validate_archive_limits(&mut archive)?;
    read_zip_bytes(&mut archive, path)
}

pub fn read_zip_bytes<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<Vec<u8>, String> {
    resource::assert_healthy()?;
    let part_id = archive
        .index_for_name(path)
        .ok_or_else(|| format!("entry not found: {path}"))?;
    let mut entry = archive
        .by_index(part_id)
        .map_err(|e| format!("entry not found: {path}: {e}"))?;
    let declared_size = entry.size();
    read_entry(&mut entry, part_id, path, declared_size, None)
}

pub fn read_zip_string<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<String, String> {
    String::from_utf8(read_zip_bytes(archive, path)?).map_err(|e| format!("read error: {e}"))
}

/// Read a deliberate UTF-8 prefix. A caller prefix does not inspect the next
/// byte; a tighter resource bound does, so it cannot silently truncate.
pub fn read_zip_string_head<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
    max_bytes: usize,
) -> Result<String, String> {
    resource::assert_healthy()?;
    let part_id = archive
        .index_for_name(path)
        .ok_or_else(|| format!("entry not found: {path}"))?;
    let mut entry = archive
        .by_index(part_id)
        .map_err(|e| format!("entry not found: {path}: {e}"))?;
    let declared_size = entry.size();
    let mut bytes = read_entry(
        &mut entry,
        part_id,
        path,
        declared_size,
        Some(max_bytes as u64),
    )?;
    match std::str::from_utf8(&bytes) {
        Ok(text) => Ok(text.to_owned()),
        Err(error) if error.error_len().is_none() => {
            bytes.truncate(error.valid_up_to());
            Ok(String::from_utf8(bytes).expect("validated UTF-8 prefix"))
        }
        Err(error) => Err(format!("read error: {error}")),
    }
}

/// Legacy fallback used only when a parser test/helper has not installed a
/// governor. Public browser paths always install a normalized session policy.
pub fn fallback_max_archive_entry_bytes() -> u64 {
    HARD_MAX_ARCHIVE_ENTRY_BYTES
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};

    fn archive_with(name: &str, body: &[u8]) -> zip::ZipArchive<Cursor<Vec<u8>>> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(name, zip::write::SimpleFileOptions::default())
                .unwrap();
            writer.write_all(body).unwrap();
            writer.finish().unwrap();
        }
        zip::ZipArchive::new(Cursor::new(bytes)).unwrap()
    }

    fn archive_with_two() -> zip::ZipArchive<Cursor<Vec<u8>>> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            writer.start_file("word/a.xml", options).unwrap();
            writer.write_all(b"1234").unwrap();
            writer.start_file("word/b.xml", options).unwrap();
            writer.write_all(b"5678").unwrap();
            writer.finish().unwrap();
        }
        zip::ZipArchive::new(Cursor::new(bytes)).unwrap()
    }

    fn details(error: &str) -> serde_json::Value {
        serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("typed resource envelope"),
        )
        .unwrap()
    }

    #[test]
    fn extracts_by_path_and_reports_missing() {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(
                    "ppt/media/image1.png",
                    zip::write::SimpleFileOptions::default(),
                )
                .unwrap();
            writer.write_all(b"\x89PNGdata").unwrap();
            writer.finish().unwrap();
        }
        assert_eq!(
            extract_zip_entry(
                &bytes,
                "ppt/media/image1.png",
                OoxmlFormat::Pptx,
                Some(64),
                Some(64)
            )
            .unwrap(),
            b"\x89PNGdata"
        );
        assert!(
            extract_zip_entry(&bytes, "missing", OoxmlFormat::Pptx, Some(64), Some(64))
                .unwrap_err()
                .contains("not found")
        );
    }

    #[test]
    fn declared_entry_limit_is_typed_and_poisoned() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(4), Some(64));
        let _scope = governor.scope("parse");
        let mut archive = archive_with("word/document.xml", b"12345678");
        validate_archive_limits(&mut archive).unwrap();
        let first = read_zip_bytes(&mut archive, "word/document.xml").unwrap_err();
        let json = details(&first);
        assert_eq!(json["code"], "ooxml-resource-limit");
        assert_eq!(json["details"]["stage"], "container");
        assert_eq!(
            json["details"]["violation"]["metric"],
            "declared-inflated-bytes"
        );
        assert_eq!(
            read_zip_bytes(&mut archive, "word/document.xml").unwrap_err(),
            first
        );
    }

    #[test]
    fn forged_small_declaration_is_stopped_by_actual_output() {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(
                    "word/document.xml",
                    zip::write::SimpleFileOptions::default(),
                )
                .unwrap();
            writer.write_all(b"12345678").unwrap();
            writer.finish().unwrap();
        }
        // Forge both local and central uncompressed-size fields down to one.
        // The stored payload still delivers eight bytes, so metadata alone must
        // not authorize the read.
        bytes[22..26].copy_from_slice(&1u32.to_le_bytes());
        let central = bytes
            .windows(4)
            .position(|window| window == 0x0201_4b50u32.to_le_bytes())
            .expect("central-directory header");
        bytes[central + 24..central + 28].copy_from_slice(&1u32.to_le_bytes());

        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(4), Some(64));
        let _scope = governor.scope("parse");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        validate_archive_limits(&mut archive).unwrap();
        let error = read_zip_bytes(&mut archive, "word/document.xml").unwrap_err();
        let json = details(&error);
        let violation = &json["details"]["violation"];
        assert_eq!(violation["metric"], "actual-inflated-bytes");
        assert_eq!(violation["limit"], 4);
        assert_eq!(violation["observed"], 5);
    }

    #[test]
    fn distinct_total_counts_two_entries_and_stops_at_limit_plus_one() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(16), Some(6));
        let _scope = governor.scope("parse");
        let mut archive = archive_with_two();
        validate_archive_limits(&mut archive).unwrap();
        assert_eq!(read_zip_bytes(&mut archive, "word/a.xml").unwrap(), b"1234");
        let error = read_zip_bytes(&mut archive, "word/b.xml").unwrap_err();
        let json = details(&error);
        assert_eq!(
            json["details"]["violation"]["metric"],
            "distinct-inflated-bytes"
        );
        assert_eq!(json["details"]["violation"]["limit"], 6);
        assert_eq!(json["details"]["violation"]["observed"], 7);
    }

    #[test]
    fn reread_does_not_double_charge_distinct_total() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(16), Some(8));
        let mut archive = archive_with_two();
        {
            let _scope = governor.scope("parse");
            validate_archive_limits(&mut archive).unwrap();
            assert_eq!(read_zip_bytes(&mut archive, "word/a.xml").unwrap(), b"1234");
        }
        {
            let _scope = governor.scope("markdown");
            assert_eq!(read_zip_bytes(&mut archive, "word/a.xml").unwrap(), b"1234");
            assert_eq!(read_zip_bytes(&mut archive, "word/b.xml").unwrap(), b"5678");
        }
        assert_eq!(governor.usage().distinct_inflated_bytes, 8);
    }

    #[test]
    fn prefix_then_full_charges_only_the_larger_observation() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(16), Some(8));
        let _scope = governor.scope("parse-sheet");
        let mut archive = archive_with("xl/sheet.xml", b"12345678");
        validate_archive_limits(&mut archive).unwrap();
        assert_eq!(
            read_zip_string_head(&mut archive, "xl/sheet.xml", 3).unwrap(),
            "123"
        );
        assert_eq!(
            read_zip_string(&mut archive, "xl/sheet.xml").unwrap(),
            "12345678"
        );
        assert_eq!(governor.usage().distinct_inflated_bytes, 8);
        assert_eq!(governor.usage().operation_inflated_bytes, 11);
    }

    #[test]
    fn full_then_prefix_adds_no_distinct_bytes() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(16), Some(8));
        let _scope = governor.scope("parse-sheet");
        let mut archive = archive_with("xl/sheet.xml", b"12345678");
        validate_archive_limits(&mut archive).unwrap();
        read_zip_bytes(&mut archive, "xl/sheet.xml").unwrap();
        read_zip_string_head(&mut archive, "xl/sheet.xml", 2).unwrap();
        assert_eq!(governor.usage().distinct_inflated_bytes, 8);
        assert_eq!(governor.usage().operation_inflated_bytes, 10);
    }

    #[test]
    fn prefix_preserves_utf8_boundary() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(16), Some(16));
        let _scope = governor.scope("probe");
        let mut archive = archive_with("xl/sheet.xml", "ab€cd".as_bytes());
        validate_archive_limits(&mut archive).unwrap();
        assert_eq!(
            read_zip_string_head(&mut archive, "xl/sheet.xml", 3).unwrap(),
            "ab"
        );
        assert_eq!(governor.usage().distinct_inflated_bytes, 3);
    }

    #[test]
    fn hard_entry_count_has_no_dummy_part() {
        let mut archive = archive_with_two();
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(16), Some(16));
        let _scope = governor.scope("open");
        // Exercise the resource layer directly with a proven raw count beyond
        // the hard quota; ZIP preflight wiring is covered separately.
        let error = resource::observe_archive_metadata(20_001, 8).unwrap_err();
        let json = details(&error);
        let violation = &json["details"]["violation"];
        assert_eq!(violation["metric"], "entry-count");
        assert!(violation.get("part").is_none());
        assert!(validate_archive_limits(&mut archive).is_err());
    }

    #[test]
    fn eager_reserve_is_bounded() {
        assert_eq!(initial_reserve(8, 64), 8);
        assert_eq!(
            initial_reserve(512 * 1024 * 1024, 512 * 1024 * 1024),
            INITIAL_RESERVE_CAP
        );
    }

    #[test]
    fn raw_eocd_entry_count_is_rejected_before_zip_open() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(64), Some(64));
        let _scope = governor.scope("open");
        let central_size = 46usize * 20_001;
        let mut eocd = vec![0; central_size];
        eocd[..4].copy_from_slice(&0x0201_4b50u32.to_le_bytes());
        eocd.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
        eocd.extend_from_slice(&0u16.to_le_bytes());
        eocd.extend_from_slice(&0u16.to_le_bytes());
        eocd.extend_from_slice(&20_001u16.to_le_bytes());
        eocd.extend_from_slice(&20_001u16.to_le_bytes());
        eocd.extend_from_slice(&(central_size as u32).to_le_bytes());
        eocd.extend_from_slice(&0u32.to_le_bytes());
        eocd.extend_from_slice(&22u16.to_le_bytes());

        // A later EOCD-looking sequence inside the legal comment is malformed.
        // Preflight must continue scanning and use the real record above.
        eocd.extend_from_slice(&EOCD_SIGNATURE.to_le_bytes());
        eocd.extend_from_slice(&1u16.to_le_bytes());
        eocd.resize(eocd.len() + 16, 0);

        let error = preflight_archive_limits(&eocd).unwrap_err();
        let json = details(&error);
        assert_eq!(json["details"]["violation"]["metric"], "entry-count");
        assert_eq!(json["details"]["violation"]["observed"], 20_001);
    }

    #[test]
    fn raw_zip64_eocd_entry_count_is_rejected_before_zip_open() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Pptx, Some(64), Some(64));
        let _scope = governor.scope("open");
        let central_size = 46usize * 20_001;
        let mut bytes = vec![0; central_size];
        bytes[..4].copy_from_slice(&0x0201_4b50u32.to_le_bytes());

        // ZIP64 end of central directory record follows the synthetic central
        // directory. Only its bounds and first signature are needed pre-open.
        let zip64_offset = bytes.len() as u64;
        bytes.extend_from_slice(&0x0606_4b50u32.to_le_bytes());
        bytes.extend_from_slice(&44u64.to_le_bytes());
        bytes.extend_from_slice(&45u16.to_le_bytes());
        bytes.extend_from_slice(&45u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&20_001u64.to_le_bytes());
        bytes.extend_from_slice(&20_001u64.to_le_bytes());
        bytes.extend_from_slice(&(central_size as u64).to_le_bytes());
        bytes.extend_from_slice(&0u64.to_le_bytes());

        // Locator points back to the ZIP64 record above.
        bytes.extend_from_slice(&0x0706_4b50u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&zip64_offset.to_le_bytes());
        bytes.extend_from_slice(&1u32.to_le_bytes());

        // Saturated classic EOCD fields require ZIP64 metadata.
        bytes.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&u16::MAX.to_le_bytes());
        bytes.extend_from_slice(&u16::MAX.to_le_bytes());
        bytes.extend_from_slice(&u32::MAX.to_le_bytes());
        bytes.extend_from_slice(&u32::MAX.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());

        let error = preflight_archive_limits(&bytes).unwrap_err();
        let json = details(&error);
        assert_eq!(json["details"]["violation"]["metric"], "entry-count");
        assert_eq!(json["details"]["violation"]["observed"], 20_001);
    }

    #[test]
    fn malformed_eocd_remains_an_ordinary_container_error() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(64), Some(64));
        let _scope = governor.scope("open");
        let mut bytes = 0x0605_4b50u32.to_le_bytes().to_vec();
        bytes.resize(22, 0);
        bytes[20..22].copy_from_slice(&10u16.to_le_bytes());
        assert!(preflight_archive_limits(&bytes).is_ok());
        assert!(governor.first_error().is_none());
    }
}

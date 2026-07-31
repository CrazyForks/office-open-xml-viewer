use wasm_bindgen::prelude::*;

mod drawing_compatibility;
mod markdown;
mod math;
mod numbering;
mod parser;
mod styles;
mod types;
mod xml_util;

/// Parse a docx archive and return the model as UTF-8 JSON **bytes**.
///
/// Returning `Vec<u8>` (a fresh copy on the JS side) instead of `String` keeps
/// the model out of the JsString/UTF-16 representation: the worker forwards the
/// resulting `ArrayBuffer` to the main thread as a transferable and the main
/// thread does a single `TextDecoder.decode` + `JSON.parse`, collapsing three
/// serializations (Rust String → JsString → structured clone) into one decode.
#[wasm_bindgen]
pub fn parse_docx(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    console_error_panic_hook::set_once();
    let doc = parser::parse_from_bytes_with_limits(
        data,
        max_archive_entry_bytes,
        max_total_inflated_bytes,
        "parse",
    )
    .map_err(docx_parser_js_error)?;
    serde_json::to_vec(&doc).map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
}

/// WASM-callable markdown projection (mirrors `to_markdown_native`). Returns
/// GitHub-flavoured markdown of headings / paragraphs / tables / footnotes,
/// discarding positioning, section properties, fonts, and drawing shapes.
#[wasm_bindgen]
pub fn docx_to_markdown(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    let doc = parser::parse_from_bytes_with_limits(
        data,
        max_archive_entry_bytes,
        max_total_inflated_bytes,
        "markdown",
    )
    .map_err(docx_markdown_js_error)?;
    Ok(markdown::render_document(&doc))
}

/// Extract raw bytes for a single embedded image entry (e.g.
/// "word/media/image1.png") from a docx zip archive. Used by the main thread to
/// lazily materialize image blobs through one bounded package operation.
#[wasm_bindgen]
pub fn extract_image(
    data: &[u8],
    path: &str,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    extract_entry_with_limits(
        data,
        path,
        max_archive_entry_bytes,
        max_total_inflated_bytes,
        "extract-image",
    )
    .map_err(|e| JsValue::from_str(&e))
}

/// A stateful handle over an opened docx archive.
///
/// The free functions above (`parse_docx` / `docx_to_markdown` / `extract_image`)
/// each re-copy the whole file into WASM and re-scan the ZIP central directory on
/// every call. A `DocxArchive` copies the bytes into WASM **once** (in `new`) and
/// keeps the opened [`parser::Zip`] session alive, so a `parse` followed by any number of
/// `extract_image` calls (the viewer's parse-then-lazily-load-media pattern)
/// pays the copy + open cost a single time. The session owns the source bytes,
/// validated central-directory index, resource governor, and first package-wide
/// poison error.
#[wasm_bindgen]
pub struct DocxArchive {
    /// The opened archive, or the container-open error string when the ZIP itself
    /// was truncated / corrupt (RB7 MAJOR). Deferring the failure here — instead of
    /// erroring out of `new` — lets `parse()` return a degraded placeholder
    /// document (symmetric with a corrupt inner part) rather than the constructor
    /// throwing an opaque error the viewer can't turn into a placeholder page.
    archive: Result<parser::Zip, String>,
}

#[wasm_bindgen]
impl DocxArchive {
    /// Copy `data` into WASM once and open the ZIP central directory once.
    /// Resource limits are retained by the package session and applied to every
    /// subsequent logical operation.
    ///
    /// `data` is taken by value (`Vec<u8>`): wasm-bindgen copies the JS `Uint8Array`
    /// once into a WASM-owned buffer and hands that allocation to Rust as this
    /// `Vec`, which `Cursor` then takes by value — a single copy across the
    /// JS→WASM boundary. Taking `&[u8]` would force a second `to_vec()` copy so
    /// the `Cursor` could own its backing store, transiently doubling WASM
    /// linear memory to ~2x the file size during construction.
    #[wasm_bindgen(constructor)]
    pub fn new(
        data: Vec<u8>,
        max_archive_entry_bytes: Option<u64>,
        max_total_inflated_bytes: Option<u64>,
    ) -> Result<DocxArchive, JsValue> {
        console_error_panic_hook::set_once();
        // RB7 (MAJOR): a truncated / corrupt CONTAINER is deferred, not thrown, so
        // `parse()` can degrade it to a placeholder document instead of the
        // constructor failing with an opaque error.
        let archive =
            parser::open_zip_with_limits(data, max_archive_entry_bytes, max_total_inflated_bytes);
        if let Err(error) = &archive {
            if error.starts_with("OOXML_RESOURCE_LIMIT:") {
                return Err(JsValue::from_str(error));
            }
        }
        Ok(DocxArchive { archive })
    }

    /// Parse the retained archive and return the model as UTF-8 JSON bytes.
    /// Byte-for-byte identical to `parse_docx` on the same file — same parser,
    /// same serializer, same error strings. When the CONTAINER failed to open
    /// (RB7 MAJOR) the model is a degraded placeholder tagged with the container.
    pub fn parse(&mut self) -> Result<Vec<u8>, JsValue> {
        let doc = match self.archive.as_mut() {
            Ok(zip) => zip.run_operation("parse", parser::parse),
            Err(e) => Ok(parser::degraded_container_document(e.clone())),
        };
        let doc = doc.map_err(docx_parser_js_error)?;
        serde_json::to_vec(&doc).map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
    }

    /// Fail cached worker operations after this package session was poisoned.
    pub fn assert_healthy(&self) -> Result<(), JsValue> {
        match &self.archive {
            Ok(zip) => zip.assert_healthy().map_err(|e| JsValue::from_str(&e)),
            Err(_) => Ok(()),
        }
    }

    /// Extract raw bytes for one embedded entry (e.g. "word/media/image1.png")
    /// from the retained archive. Twin of the free `extract_image`, but reads
    /// through the already-open archive instead of re-opening it. A corrupt
    /// container has no entries, so this surfaces the container-open error.
    pub fn extract_image(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        let zip = self
            .archive
            .as_mut()
            .map_err(|e| JsValue::from_str(&format!("docx-parser error: {e}")))?;
        zip.run_operation("extract-image", |zip| parser::read_zip_bytes(zip, path))
            .map_err(|e| JsValue::from_str(&e))
    }

    /// GitHub-flavoured markdown projection of the retained archive. Mirrors the
    /// free `docx_to_markdown`. A corrupt container degrades to an empty document.
    pub fn to_markdown(&mut self) -> Result<String, JsValue> {
        let doc = match self.archive.as_mut() {
            Ok(zip) => zip.run_operation("markdown", parser::parse),
            Err(e) => Ok(parser::degraded_container_document(e.clone())),
        };
        let doc = doc.map_err(docx_markdown_js_error)?;
        Ok(markdown::render_document(&doc))
    }
}

/// Native equivalent of `parse_docx` for use from the MCP server.
#[cfg(not(target_arch = "wasm32"))]
pub fn parse_docx_native(data: &[u8]) -> Result<String, String> {
    parser::parse_from_bytes(data)
        .and_then(|doc| serde_json::to_string(&doc).map_err(|e| e.to_string()))
}

/// Parse a docx and project the result to GitHub-flavoured markdown:
/// headings (from outlineLevel), paragraphs with bullet/numbered lists,
/// tables, footnote references collated at the end, and rich-text
/// formatting (bold / italic / strikethrough / hyperlink). Designed for AI
/// agents that need to read content efficiently — discards positioning,
/// section properties, font metrics, drawing shapes.
#[cfg(not(target_arch = "wasm32"))]
pub fn to_markdown_native(data: &[u8]) -> Result<String, String> {
    let doc = parser::parse_from_bytes_with_limits(data, None, None, "markdown")?;
    Ok(markdown::render_document(&doc))
}

fn docx_parser_js_error(error: String) -> JsValue {
    if error.starts_with("OOXML_RESOURCE_LIMIT:") {
        JsValue::from_str(&error)
    } else {
        JsValue::from_str(&format!("docx-parser error: {error}"))
    }
}

fn docx_markdown_js_error(error: String) -> JsValue {
    JsValue::from_str(&error)
}

fn extract_entry_with_limits(
    data: &[u8],
    path: &str,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
    operation: &str,
) -> Result<Vec<u8>, String> {
    let mut zip = parser::open_zip_with_limits(
        data.to_vec(),
        max_archive_entry_bytes,
        max_total_inflated_bytes,
    )?;
    zip.run_operation(operation, |zip| parser::read_zip_bytes(zip, path))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn zip_parts(parts: &[(&str, &[u8])]) -> Vec<u8> {
        use std::io::{Cursor, Write};
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            for (path, body) in parts {
                writer.start_file(path, options).unwrap();
                writer.write_all(body).unwrap();
            }
            writer.finish().unwrap();
        }
        bytes
    }

    fn forge_declared_size(bytes: &mut [u8], target: &str, declared_size: u32) {
        let target = target.as_bytes();
        let mut cursor = 0;
        let mut local_found = false;
        while cursor + 30 <= bytes.len() {
            if bytes[cursor..cursor + 4] == 0x0403_4b50u32.to_le_bytes() {
                let name_len =
                    u16::from_le_bytes([bytes[cursor + 26], bytes[cursor + 27]]) as usize;
                let extra_len =
                    u16::from_le_bytes([bytes[cursor + 28], bytes[cursor + 29]]) as usize;
                let name_start = cursor + 30;
                let name_end = name_start + name_len;
                if name_end <= bytes.len() && &bytes[name_start..name_end] == target {
                    bytes[cursor + 22..cursor + 26].copy_from_slice(&declared_size.to_le_bytes());
                    local_found = true;
                    break;
                }
                cursor = name_end.saturating_add(extra_len);
            } else {
                cursor += 1;
            }
        }

        cursor = 0;
        let mut central_found = false;
        while cursor + 46 <= bytes.len() {
            if bytes[cursor..cursor + 4] == 0x0201_4b50u32.to_le_bytes() {
                let name_len =
                    u16::from_le_bytes([bytes[cursor + 28], bytes[cursor + 29]]) as usize;
                let extra_len =
                    u16::from_le_bytes([bytes[cursor + 30], bytes[cursor + 31]]) as usize;
                let comment_len =
                    u16::from_le_bytes([bytes[cursor + 32], bytes[cursor + 33]]) as usize;
                let name_start = cursor + 46;
                let name_end = name_start + name_len;
                if name_end <= bytes.len() && &bytes[name_start..name_end] == target {
                    bytes[cursor + 24..cursor + 28].copy_from_slice(&declared_size.to_le_bytes());
                    central_found = true;
                    break;
                }
                cursor = name_end
                    .saturating_add(extra_len)
                    .saturating_add(comment_len);
            } else {
                cursor += 1;
            }
        }
        assert!(local_found && central_found, "target part must be forged");
    }

    fn forged_document_package() -> Vec<u8> {
        let padding = "x".repeat(4096);
        let document = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{padding}</w:t></w:r></w:p></w:body></w:document>"#
        );
        let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
        let mut bytes = zip_parts(&[
            ("word/document.xml", document.as_bytes()),
            ("word/_rels/document.xml.rels", rels),
        ]);
        forge_declared_size(&mut bytes, "word/document.xml", 1);
        bytes
    }

    fn forged_optional_styles_package() -> Vec<u8> {
        let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#;
        let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
        let padding = "x".repeat(4096);
        let styles = format!(
            r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><!--{padding}--></w:styles>"#
        );
        let mut bytes = zip_parts(&[
            ("word/document.xml", document),
            ("word/_rels/document.xml.rels", rels),
            ("word/styles.xml", styles.as_bytes()),
        ]);
        forge_declared_size(&mut bytes, "word/styles.xml", 1);
        bytes
    }

    #[test]
    fn extract_image_reads_entry() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            w.start_file("word/media/i.png", o).unwrap();
            w.write_all(b"X").unwrap();
            w.finish().unwrap();
        }
        assert_eq!(
            extract_image(&buf, "word/media/i.png", None, None).unwrap(),
            b"X"
        );
    }

    /// The docx JSON path never hand-assembles `{"error":"…"}` — a message with a
    /// `"` used to produce invalid JSON that made the TS-side `JSON.parse` throw a
    /// confusing SyntaxError. Since RB7 (MAJOR) a non-zip / corrupt CONTAINER no
    /// longer errors at all: `parse_from_bytes` degrades to a placeholder Document
    /// whose `parse_error` field is serialized by serde, so any quotes in the
    /// message are escaped by construction. This pins both facts: the input
    /// degrades (does not panic / error out) AND the placeholder serializes to
    /// valid JSON with the message intact.
    #[test]
    fn parse_non_zip_bytes_degrades_without_json_escaping_hazard() {
        // Not a zip archive — degrades to a placeholder, does not error or panic.
        let doc = parser::parse_from_bytes(&[1, 2, 3])
            .expect("non-zip bytes degrade to a placeholder, not an error");
        let err = doc
            .parse_error
            .as_deref()
            .expect("placeholder carries a container-tagged parse_error");
        assert!(
            err.contains("zip container"),
            "names the container; got {err:?}"
        );
        // serde escapes any quotes: the serialized model is valid JSON and the
        // message round-trips through it unharmed (the old hand-built JSON hazard).
        let json = serde_json::to_string(&doc).expect("serializes to valid JSON");
        let round: serde_json::Value = serde_json::from_str(&json).expect("valid JSON");
        assert_eq!(
            round["parseError"],
            serde_json::Value::String(err.to_string()),
            "parse_error round-trips through serde JSON intact"
        );
    }

    #[test]
    fn document_resource_overrun_is_typed_and_poisons_the_stateful_session() {
        const LIMIT: u64 = 2048;
        let mut zip =
            parser::open_zip_with_limits(forged_document_package(), Some(LIMIT), Some(64 * 1024))
                .expect("forged declaration passes package preflight");
        let first = zip
            .run_operation("parse", parser::parse)
            .expect_err("degraded document must not hide a resource violation");
        let details: serde_json::Value = serde_json::from_str(
            first
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("canonical typed resource-limit envelope"),
        )
        .unwrap();
        let violation = &details["details"]["violation"];
        assert_eq!(violation["operation"], "parse");
        assert_eq!(violation["part"], "word/document.xml");
        assert_eq!(violation["metric"], "actual-inflated-bytes");
        assert_eq!(violation["limit"], LIMIT);
        assert_eq!(violation["observed"], LIMIT + 1);

        let later = zip
            .run_operation("markdown", |_| Ok(()))
            .expect_err("poisoned session rejects later operations");
        assert_eq!(later, first);
    }

    #[test]
    fn free_parse_helper_attributes_resource_failure_to_markdown() {
        let error = parser::parse_from_bytes_with_limits(
            &forged_document_package(),
            Some(2048),
            Some(64 * 1024),
            "markdown",
        )
        .expect_err("free parse helper surfaces the package poison");
        let details: serde_json::Value = serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("canonical typed resource-limit envelope"),
        )
        .unwrap();
        assert_eq!(details["details"]["violation"]["operation"], "markdown");
        assert_eq!(details["details"]["violation"]["part"], "word/document.xml");
    }

    #[test]
    fn optional_styles_resource_overrun_survives_fallback_and_settles_typed() {
        const LIMIT: u64 = 2048;
        let error = parser::parse_from_bytes_with_limits(
            &forged_optional_styles_package(),
            Some(LIMIT),
            Some(64 * 1024),
            "parse",
        )
        .expect_err("optional styles fallback must not hide package poison");
        let details: serde_json::Value = serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("canonical typed resource-limit envelope"),
        )
        .unwrap();
        let violation = &details["details"]["violation"];
        assert_eq!(violation["operation"], "parse");
        assert_eq!(violation["part"], "word/styles.xml");
        assert_eq!(violation["metric"], "actual-inflated-bytes");
        assert_eq!(violation["limit"], LIMIT);
        assert_eq!(violation["observed"], LIMIT + 1);
    }

    #[test]
    fn large_image_stays_lazy_during_parse_but_extract_is_bounded() {
        const TOTAL_LIMIT: u64 = 4096;
        let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#;
        let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/large.png"/></Relationships>"#;
        let image = vec![0x5a; 64 * 1024];
        let data = zip_parts(&[
            ("word/document.xml", document),
            ("word/_rels/document.xml.rels", rels),
            ("word/media/large.png", &image),
        ]);

        let doc = parser::parse_from_bytes_with_limits(&data, None, Some(TOTAL_LIMIT), "parse")
            .expect("media existence check must not inflate the image");
        assert!(doc.parse_error.is_none());

        let error = extract_entry_with_limits(
            &data,
            "word/media/large.png",
            None,
            Some(TOTAL_LIMIT),
            "extract-image",
        )
        .expect_err("materializing the large image exceeds the total budget");
        let details: serde_json::Value = serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("canonical typed resource-limit envelope"),
        )
        .unwrap();
        let violation = &details["details"]["violation"];
        assert_eq!(violation["operation"], "extract-image");
        assert_eq!(violation["part"], "word/media/large.png");
        assert_eq!(violation["metric"], "distinct-inflated-bytes");
        assert_eq!(violation["limit"], TOTAL_LIMIT);
        assert_eq!(violation["observed"], TOTAL_LIMIT + 1);
    }

    #[test]
    fn ordinary_rb7_failures_still_degrade() {
        let malformed = zip_parts(&[
            ("word/document.xml", br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#),
            ("word/_rels/document.xml.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#),
        ]);
        let malformed_doc = parser::parse_from_bytes_with_limits(&malformed, None, None, "parse")
            .expect("malformed document.xml degrades");
        assert!(malformed_doc
            .parse_error
            .as_deref()
            .is_some_and(|error| error.starts_with("word/document.xml:")));

        let missing = zip_parts(&[(
            "word/_rels/document.xml.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
        )]);
        let missing_doc = parser::parse_from_bytes_with_limits(&missing, None, None, "parse")
            .expect("missing document.xml degrades");
        assert!(missing_doc
            .parse_error
            .as_deref()
            .is_some_and(|error| error.starts_with("word/document.xml:")));

        let container = parser::parse_from_bytes_with_limits(b"not a zip", None, None, "parse")
            .expect("ordinary corrupt container degrades");
        assert!(container
            .parse_error
            .as_deref()
            .is_some_and(|error| error.starts_with("(zip container): ")));
    }

    #[test]
    #[allow(clippy::type_complexity)] // Exact exported ABI shapes are the assertion.
    fn public_signatures_remain_stable() {
        let _: fn(&[u8], Option<u64>, Option<u64>) -> Result<Vec<u8>, JsValue> = parse_docx;
        let _: fn(&[u8], Option<u64>, Option<u64>) -> Result<String, JsValue> = docx_to_markdown;
        let _: fn(&[u8], &str, Option<u64>, Option<u64>) -> Result<Vec<u8>, JsValue> =
            extract_image;
        let _: fn(Vec<u8>, Option<u64>, Option<u64>) -> Result<DocxArchive, JsValue> =
            DocxArchive::new;
        let _: fn(&mut DocxArchive) -> Result<Vec<u8>, JsValue> = DocxArchive::parse;
        let _: fn(&mut DocxArchive, &str) -> Result<Vec<u8>, JsValue> = DocxArchive::extract_image;
        let _: fn(&mut DocxArchive) -> Result<String, JsValue> = DocxArchive::to_markdown;
        let _: fn(&DocxArchive) -> Result<(), JsValue> = DocxArchive::assert_healthy;
        let _: fn(&[u8]) -> Result<String, String> = parse_docx_native;
        let _: fn(&[u8]) -> Result<String, String> = to_markdown_native;
    }
}

use docx_parser::parse_docx_native;
use serde_json::{json, Value};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const STRICT_W_NS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_WP_NS: &str = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";

fn build_docx(body: &str) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="{W_NS}" xmlns:wp="{WP_NS}">
             <w:body>{body}<w:sectPr/></w:body>
           </w:document>"#
    );
    let mut bytes = Vec::new();
    {
        let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
        zip.start_file("word/document.xml", SimpleFileOptions::default())
            .expect("document entry");
        zip.write_all(document.as_bytes()).expect("document XML");
        zip.finish().expect("finish zip");
    }
    bytes
}

fn parse(body: &str) -> Value {
    let json = parse_docx_native(&build_docx(body)).expect("minimal DOCX parses");
    serde_json::from_str(&json).expect("parser output JSON")
}

#[test]
fn reports_stable_private_diagnostics_without_raw_values_or_document_text() {
    let document = parse(
        r#"
        <w:p>
          <w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr><w:t>PRIVATE_SENTINEL</w:t></w:r>
          <w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr><w:t>duplicate</w:t></w:r>
        </w:p>
        <w:p><w:r><w:rPr><w:effect w:val="producer-private-value"/></w:rPr></w:r></w:p>
        <w:p><w:r><w:drawing><wp:inline/></w:drawing></w:r></w:p>
        <w:p><w:r><w:drawing>
          <wp:anchor><wp:extent cx="not-a-coordinate" cy="-1"/></wp:anchor>
        </w:drawing></w:r></w:p>
        <w:p><w:r><w:drawing>
          <wp:inline><wp:extent cx="0" cy="12700"/></wp:inline>
        </w:drawing></w:r></w:p>
        "#,
    );

    assert_eq!(
        document["diagnostics"],
        json!([
          {
            "code": "UNSUPPORTED_TEXT_EFFECT",
            "severity": "warning",
            "part": "word/document.xml",
            "path": [0]
          },
          {
            "code": "INVALID_TEXT_EFFECT_VALUE",
            "severity": "warning",
            "part": "word/document.xml",
            "path": [1]
          },
          {
            "code": "MISSING_DRAWING_EXTENT",
            "severity": "error",
            "part": "word/document.xml",
            "path": [2]
          },
          {
            "code": "INVALID_DRAWING_EXTENT",
            "severity": "error",
            "part": "word/document.xml",
            "path": [3]
          },
          {
            "code": "DEGENERATE_DRAWING_EXTENT",
            "severity": "warning",
            "part": "word/document.xml",
            "path": [4]
          }
        ])
    );

    let diagnostics = serde_json::to_string(&document["diagnostics"]).unwrap();
    assert!(!diagnostics.contains("PRIVATE_SENTINEL"));
    assert!(!diagnostics.contains("producer-private-value"));
    assert!(!diagnostics.contains("not-a-coordinate"));
}

#[test]
fn uses_the_emitted_body_cursor_across_wrappers_and_split_paragraphs() {
    let wrapped = parse(
        r#"
        <w:sdt><w:sdtContent>
          <w:p><w:r><w:t>first</w:t></w:r></w:p>
          <w:p><w:r><w:t>second</w:t></w:r></w:p>
        </w:sdtContent></w:sdt>
        <w:p><w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr></w:r></w:p>
        "#,
    );
    assert_eq!(wrapped["diagnostics"][0]["path"], json!([2]));

    let split = parse(
        r#"
        <w:p>
          <w:r><w:t>before</w:t></w:r>
          <w:r><w:br w:type="page"/></w:r>
          <w:r><w:t>after</w:t></w:r>
        </w:p>
        <w:p><w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr></w:r></w:p>
        "#,
    );
    assert_eq!(split["diagnostics"][0]["path"], json!([3]));
}

#[test]
fn supported_and_schema_valid_control_values_do_not_report_diagnostics() {
    let document = parse(
        r#"
        <w:p>
          <w:r><w:rPr><w:effect w:val="none"/></w:rPr><w:t>control</w:t></w:r>
          <w:r><w:drawing>
            <wp:inline><wp:extent cx="12700" cy="25400"/></wp:inline>
          </w:drawing></w:r>
          <w:r><w:rPr><w:vanish/><w:effect w:val="sparkle"/></w:rPr></w:r>
          <w:r><w:drawing><wp:inline><wp:docPr hidden="1"/></wp:inline></w:drawing></w:r>
        </w:p>
        "#,
    );

    assert!(
        document.get("diagnostics").is_none(),
        "the empty private wire stays omitted"
    );
}

#[test]
fn follows_selected_mce_branch_and_remaps_removed_cover_breaks() {
    let selected_fallback = parse(&format!(
        r#"
        <w:p xmlns:mc="{MC_NS}" xmlns:future="urn:unsupported:drawing">
          <w:r><mc:AlternateContent>
            <mc:Choice Requires="future">
              <w:drawing><wp:inline/></w:drawing>
            </mc:Choice>
            <mc:Fallback>
              <w:drawing><wp:inline><wp:extent cx="12700" cy="12700"/></wp:inline></w:drawing>
            </mc:Fallback>
          </mc:AlternateContent></w:r>
        </w:p>
        "#
    ));
    assert!(
        selected_fallback.get("diagnostics").is_none(),
        "unselected MCE markup must not produce diagnostics"
    );

    let cover = parse(
        r#"
        <w:sdt>
          <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Cover Pages"/></w:docPartObj></w:sdtPr>
          <w:sdtContent><w:p><w:r><w:br w:type="page"/></w:r></w:p></w:sdtContent>
        </w:sdt>
        <w:p><w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr></w:r></w:p>
        "#,
    );
    assert_eq!(
        cover["diagnostics"][0]["path"],
        json!([1]),
        "removing the redundant synthetic cover break remaps later source paths"
    );
}

#[test]
fn accepts_strict_namespaces_and_positive_coordinate_boundaries() {
    let strict = parse(&format!(
        r#"
        <w:p xmlns:w="{STRICT_W_NS}" xmlns:wp="{STRICT_WP_NS}">
          <w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr></w:r>
          <w:r><w:drawing><wp:inline>
            <wp:extent cx=" 27273042316900 " cy="+1"/>
          </wp:inline></w:drawing></w:r>
        </w:p>
        <w:p xmlns:wp="{STRICT_WP_NS}"><w:r><w:drawing><wp:inline>
          <wp:extent cx="27273042316901" cy="1"/>
        </wp:inline></w:drawing></w:r></w:p>
        "#
    ));

    assert_eq!(
        strict["diagnostics"],
        json!([
          {
            "code": "UNSUPPORTED_TEXT_EFFECT",
            "severity": "warning",
            "part": "word/document.xml",
            "path": [0]
          },
          {
            "code": "INVALID_DRAWING_EXTENT",
            "severity": "error",
            "part": "word/document.xml",
            "path": [1]
          }
        ])
    );
}

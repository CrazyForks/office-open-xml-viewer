//! Bounded, namespace-preserving projection of `word/document.xml` into
//! complete logical body blocks.
//!
//! WordprocessingML is not page-random-access: sections, fields, notes, and
//! pagination depend on the complete preceding story. This projector does not
//! pretend otherwise. It removes only the whole-part XML arena: a first pass
//! validates the part and records compact cross-block facts, while a second
//! pass can hand one complete paragraph/table/section block at a time to the
//! existing semantic parser.

use std::io::BufRead;
use std::rc::Rc;

use ooxml_common::bounded_xml::{
    self, BoundedXmlError, BoundedXmlReadError, BoundedXmlReader, NamespaceContext,
};
use ooxml_common::depth::MAX_XML_DEPTH;
use ooxml_common::ns::is_w_ns;
use ooxml_common::package_session::PackageLimitReporter;
use ooxml_common::resource::{HardResourceLimitKind, HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES};
use quick_xml::events::{BytesStart, Event};
use quick_xml::XmlVersion;

const DOCUMENT_PART: &str = "word/document.xml";
const STREAMED_XML_EVENT_BYTES: usize = 1024 * 1024;
const STREAMED_XML_CONTEXT_BYTES: usize = 4 * 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ProjectedBodyBlock {
    pub(crate) ordinal: usize,
    /// Byte offset of the source block start in `word/document.xml`. This is a
    /// stable parser-owned identity used by logical-table acquisition facts.
    pub(crate) source_offset: usize,
    pub(crate) local_name: String,
    pub(crate) xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DocumentBodyPlan {
    pub(crate) cover_break_after: Vec<bool>,
}

struct ElementFrame {
    context: Rc<NamespaceContext>,
}

#[derive(Default)]
struct SdtFrame {
    open_depth: usize,
    properties_depth: Option<usize>,
    content_depth: Option<usize>,
    cover_pages: bool,
    last_block_ordinal: Option<usize>,
}

struct Capture {
    root_depth: usize,
    source_offset: usize,
    local_name: String,
    xml: Vec<u8>,
}

/// Streaming body projector. `next_block` retains at most one complete logical
/// block and bounded active namespace state. Block-level `w:sdt` wrappers are
/// transparent, matching `element_children_flat`; their cached `w:sdtContent`
/// children become independent blocks.
pub(crate) struct DocumentBodyProjector<R: BufRead> {
    reader: BoundedXmlReader<R>,
    reporter: Option<PackageLimitReporter>,
    frames: Vec<ElementFrame>,
    depth: usize,
    body_content_depth: Option<usize>,
    body_seen: bool,
    body_closed: bool,
    sdts: Vec<SdtFrame>,
    capture: Option<Capture>,
    cover_break_after: Vec<bool>,
    finished: bool,
}

impl<R: BufRead> DocumentBodyProjector<R> {
    pub(crate) fn new(source: R, reporter: Option<PackageLimitReporter>) -> Self {
        Self {
            reader: BoundedXmlReader::new(source, STREAMED_XML_EVENT_BYTES, "document body"),
            reporter,
            frames: Vec::new(),
            depth: 0,
            body_content_depth: None,
            body_seen: false,
            body_closed: false,
            sdts: Vec::new(),
            capture: None,
            cover_break_after: Vec::new(),
            finished: false,
        }
    }

    pub(crate) fn next_block(&mut self) -> Result<Option<ProjectedBodyBlock>, String> {
        if self.finished {
            return Ok(None);
        }
        loop {
            let bounded = self
                .reader
                .read_event()
                .map_err(|error| self.map_read_error(error))?;
            let source_offset = usize::try_from(bounded.span.start)
                .map_err(|_| "document body source offset exceeds this platform".to_string())?;
            let namespace = bounded.namespace.as_deref();
            match bounded.event {
                Event::Start(start) => {
                    let inherited = self
                        .frames
                        .last()
                        .map(|frame| Rc::clone(&frame.context))
                        .unwrap_or_else(NamespaceContext::root);
                    let context = NamespaceContext::derive(
                        &start,
                        &inherited,
                        STREAMED_XML_CONTEXT_BYTES,
                        "document body",
                    )
                    .map_err(|error| {
                        self.map_bounded_error(error, HardResourceLimitKind::XmlContextBytes)
                    })?;
                    let next_depth = self.depth.saturating_add(1);
                    if next_depth > MAX_XML_DEPTH as usize {
                        return Err(self.report_limit(
                            HardResourceLimitKind::XmlNestingDepth,
                            MAX_XML_DEPTH as usize,
                            next_depth,
                        ));
                    }

                    if self.capture.is_some() {
                        self.append_capture(Event::Start(start))?;
                    } else {
                        self.handle_start(namespace, start, &context, next_depth, source_offset)?;
                    }
                    self.frames.push(ElementFrame { context });
                    self.depth = next_depth;
                }
                Event::Empty(empty) => {
                    let inherited = self
                        .frames
                        .last()
                        .map(|frame| Rc::clone(&frame.context))
                        .unwrap_or_else(NamespaceContext::root);
                    let context = NamespaceContext::derive(
                        &empty,
                        &inherited,
                        STREAMED_XML_CONTEXT_BYTES,
                        "document body",
                    )
                    .map_err(|error| {
                        self.map_bounded_error(error, HardResourceLimitKind::XmlContextBytes)
                    })?;
                    if self.capture.is_some() {
                        self.append_capture(Event::Empty(empty))?;
                    } else if let Some(block) =
                        self.handle_empty(namespace, empty, &context, source_offset)?
                    {
                        return Ok(Some(block));
                    }
                }
                Event::End(end) => {
                    if let Some(capture) = self.capture.as_ref() {
                        let closes_capture = self.depth == capture.root_depth;
                        self.append_capture(Event::End(end))?;
                        self.depth = self.depth.saturating_sub(1);
                        self.frames.pop();
                        if closes_capture {
                            return self.finish_capture().map(Some);
                        }
                    } else {
                        self.handle_end(namespace, end.local_name().as_ref())?;
                        self.depth = self.depth.saturating_sub(1);
                        self.frames.pop();
                    }
                }
                Event::Eof => {
                    self.finished = true;
                    if self.capture.is_some() {
                        return Err(
                            "word/document.xml ended inside a logical body block".to_string()
                        );
                    }
                    if !self.body_seen {
                        return Err("word/document.xml: no <w:body> element".to_string());
                    }
                    if !self.body_closed {
                        return Err("word/document.xml ended before </w:body>".to_string());
                    }
                    return Ok(None);
                }
                event => {
                    if self.capture.is_some() {
                        self.append_capture(event)?;
                    }
                }
            }
        }
    }

    pub(crate) fn plan(&self) -> Result<DocumentBodyPlan, String> {
        if !self.finished {
            return Err("document body projection plan is not complete".to_string());
        }
        Ok(DocumentBodyPlan {
            cover_break_after: self.cover_break_after.clone(),
        })
    }

    fn handle_start(
        &mut self,
        namespace: Option<&str>,
        start: BytesStart<'static>,
        context: &Rc<NamespaceContext>,
        next_depth: usize,
        source_offset: usize,
    ) -> Result<(), String> {
        let local_name = local_name(&start)?;
        let is_word = is_w_ns(namespace);
        if is_word && local_name == "body" && self.body_content_depth.is_none() {
            self.body_seen = true;
            self.body_content_depth = Some(next_depth);
            return Ok(());
        }

        self.observe_sdt_metadata(is_word, &local_name, &start, next_depth)?;
        if self.at_logical_level() && is_word && local_name == "sdt" {
            self.sdts.push(SdtFrame {
                open_depth: next_depth,
                ..SdtFrame::default()
            });
            return Ok(());
        }
        if self.at_logical_level() && is_word && is_body_block(&local_name) {
            let mut root = start;
            bounded_xml::inject_missing_namespaces(
                &mut root,
                context,
                None,
                STREAMED_XML_CONTEXT_BYTES,
                "document body block",
            )
            .map_err(|error| {
                self.map_bounded_error(error, HardResourceLimitKind::XmlContextBytes)
            })?;
            self.capture = Some(Capture {
                root_depth: next_depth,
                source_offset,
                local_name,
                xml: Vec::new(),
            });
            self.append_capture(Event::Start(root))?;
        }
        Ok(())
    }

    fn handle_empty(
        &mut self,
        namespace: Option<&str>,
        empty: BytesStart<'static>,
        context: &Rc<NamespaceContext>,
        source_offset: usize,
    ) -> Result<Option<ProjectedBodyBlock>, String> {
        let local_name = local_name(&empty)?;
        let is_word = is_w_ns(namespace);
        self.observe_sdt_metadata(is_word, &local_name, &empty, self.depth)?;
        if !(self.at_logical_level() && is_word && is_body_block(&local_name)) {
            return Ok(None);
        }
        let mut root = empty;
        bounded_xml::inject_missing_namespaces(
            &mut root,
            context,
            None,
            STREAMED_XML_CONTEXT_BYTES,
            "document body block",
        )
        .map_err(|error| self.map_bounded_error(error, HardResourceLimitKind::XmlContextBytes))?;
        self.capture = Some(Capture {
            root_depth: self.depth,
            source_offset,
            local_name,
            xml: Vec::new(),
        });
        self.append_capture(Event::Empty(root))?;
        self.finish_capture().map(Some)
    }

    fn handle_end(&mut self, namespace: Option<&str>, raw_local_name: &[u8]) -> Result<(), String> {
        let local_name = std::str::from_utf8(raw_local_name)
            .map_err(|error| format!("document body element name: {error}"))?;
        if is_w_ns(namespace) && local_name == "body" && self.body_content_depth == Some(self.depth)
        {
            self.body_content_depth = None;
            self.body_closed = true;
        }
        if let Some(sdt) = self.sdts.last_mut() {
            if sdt.properties_depth == Some(self.depth) {
                sdt.properties_depth = None;
            }
            if sdt.content_depth == Some(self.depth) {
                sdt.content_depth = None;
            }
            if is_w_ns(namespace) && local_name == "sdt" && sdt.open_depth == self.depth {
                let completed = self.sdts.pop().expect("matching SDT frame exists");
                if completed.cover_pages {
                    if let Some(ordinal) = completed.last_block_ordinal {
                        if let Some(flag) = self.cover_break_after.get_mut(ordinal) {
                            *flag = true;
                        }
                    }
                }
            }
        }
        Ok(())
    }

    fn observe_sdt_metadata(
        &mut self,
        is_word: bool,
        local_name: &str,
        element: &BytesStart<'_>,
        next_depth: usize,
    ) -> Result<(), String> {
        let Some(sdt) = self.sdts.last_mut() else {
            return Ok(());
        };
        if is_word && self.depth == sdt.open_depth {
            if local_name == "sdtPr" {
                sdt.properties_depth = Some(next_depth);
            } else if local_name == "sdtContent" {
                sdt.content_depth = Some(next_depth);
            }
        }
        if is_word
            && local_name == "docPartGallery"
            && sdt
                .properties_depth
                .is_some_and(|depth| self.depth >= depth)
            && attribute_value(element, "val")?.as_deref() == Some("Cover Pages")
        {
            sdt.cover_pages = true;
        }
        Ok(())
    }

    fn at_logical_level(&self) -> bool {
        let Some(container_depth) = self
            .sdts
            .last()
            .and_then(|sdt| sdt.content_depth)
            .or(self.body_content_depth)
        else {
            return false;
        };
        self.depth == container_depth
    }

    fn append_capture(&mut self, event: Event<'static>) -> Result<(), String> {
        let capture = self
            .capture
            .as_mut()
            .ok_or_else(|| "document body capture is not active".to_string())?;
        bounded_xml::append_projected_event(
            &mut capture.xml,
            &event,
            0,
            HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize,
            "document body block",
        )
        .map_err(|error| {
            self.map_bounded_error(error, HardResourceLimitKind::DocxBodyBlockXmlBytes)
        })
    }

    fn finish_capture(&mut self) -> Result<ProjectedBodyBlock, String> {
        let capture = self
            .capture
            .take()
            .ok_or_else(|| "document body capture is not active".to_string())?;
        let ordinal = self.cover_break_after.len();
        self.cover_break_after.push(false);
        for sdt in &mut self.sdts {
            if sdt
                .content_depth
                .is_some_and(|depth| capture.root_depth > depth)
            {
                sdt.last_block_ordinal = Some(ordinal);
            }
        }
        Ok(ProjectedBodyBlock {
            ordinal,
            source_offset: capture.source_offset,
            local_name: capture.local_name,
            xml: capture.xml,
        })
    }

    fn map_read_error(&self, error: BoundedXmlReadError) -> String {
        match error {
            BoundedXmlReadError::Xml(message) => message,
            BoundedXmlReadError::EventLimit { limit, observed } => {
                self.report_limit(HardResourceLimitKind::XmlEventBytes, limit, observed)
            }
        }
    }

    fn map_bounded_error(&self, error: BoundedXmlError, kind: HardResourceLimitKind) -> String {
        match error {
            BoundedXmlError::Xml(message) => message,
            BoundedXmlError::Limit { limit, observed } => self.report_limit(kind, limit, observed),
        }
    }

    fn report_limit(&self, kind: HardResourceLimitKind, limit: usize, observed: usize) -> String {
        match bounded_xml::report_hard_limit(
            self.reporter.as_ref(),
            kind,
            Some(DOCUMENT_PART),
            limit,
            observed,
        ) {
            Err(message) => message,
            Ok(()) => format!("document body projector limit exceeded: {observed} > {limit}"),
        }
    }
}

fn local_name(element: &BytesStart<'_>) -> Result<String, String> {
    std::str::from_utf8(element.local_name().as_ref())
        .map(str::to_string)
        .map_err(|error| format!("document body element name: {error}"))
}

fn is_body_block(local_name: &str) -> bool {
    matches!(local_name, "p" | "tbl" | "sectPr")
}

fn attribute_value(element: &BytesStart<'_>, local_name: &str) -> Result<Option<String>, String> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| format!("document body attribute: {error}"))?;
        let key = attribute.key.as_ref();
        let key = key.rsplit(|byte| *byte == b':').next().unwrap_or(key);
        if key != local_name.as_bytes() {
            continue;
        }
        let value = attribute
            .normalized_value(XmlVersion::Implicit1_0)
            .map_err(|error| format!("document body attribute value: {error}"))?;
        return Ok(Some(value.into_owned()));
    }
    Ok(None)
}

#[cfg(test)]
mod tests {
    use std::io::{BufReader, Cursor};

    use super::*;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn project(xml: &str) -> (Vec<ProjectedBodyBlock>, DocumentBodyPlan) {
        let mut projector =
            DocumentBodyProjector::new(BufReader::new(Cursor::new(xml.as_bytes().to_vec())), None);
        let mut blocks = Vec::new();
        while let Some(block) = projector.next_block().unwrap() {
            blocks.push(block);
        }
        let plan = projector.plan().unwrap();
        (blocks, plan)
    }

    #[test]
    fn projects_body_blocks_without_retaining_the_document_wrapper() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:r="urn:rels"><w:body>
              <w:p><w:r><w:t>A</w:t></w:r></w:p>
              <w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
              <w:sectPr/>
            </w:body></w:document>"#,
        );
        let (blocks, plan) = project(&xml);
        assert_eq!(
            blocks
                .iter()
                .map(|block| block.local_name.as_str())
                .collect::<Vec<_>>(),
            ["p", "tbl", "sectPr"]
        );
        assert_eq!(plan.cover_break_after, [false, false, false]);
        assert_eq!(blocks[0].source_offset, xml.find("<w:p>").unwrap());
        assert_eq!(blocks[1].source_offset, xml.find("<w:tbl>").unwrap());
        for block in blocks {
            let text = String::from_utf8(block.xml).unwrap();
            assert!(
                text.contains("xmlns:w="),
                "inherited namespace must be repaired: {text}"
            );
            roxmltree::Document::parse(&text).unwrap();
        }
    }

    #[test]
    fn flattens_nested_block_sdts_and_marks_only_the_cover_tail() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:body>
              <w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Cover Pages"/></w:docPartObj></w:sdtPr>
                <w:sdtContent><w:p/><w:sdt><w:sdtContent><w:p/><w:tbl/></w:sdtContent></w:sdt></w:sdtContent>
              </w:sdt>
              <w:p/>
            </w:body></w:document>"#,
        );
        let (blocks, plan) = project(&xml);
        assert_eq!(
            blocks
                .iter()
                .map(|block| block.local_name.as_str())
                .collect::<Vec<_>>(),
            ["p", "p", "tbl", "p"]
        );
        assert_eq!(plan.cover_break_after, [false, false, true, false]);
    }

    #[test]
    fn ignores_unsupported_top_level_elements_but_validates_the_tail() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:altChunk/><w:p/></w:body></w:document>"#,
        );
        let (blocks, _) = project(&xml);
        assert_eq!(blocks.len(), 1);
        assert_eq!(blocks[0].local_name, "p");
    }

    #[test]
    fn missing_body_is_fatal_for_the_streamed_required_part() {
        let xml = format!(r#"<w:document xmlns:w="{W}"/>"#);
        let mut projector =
            DocumentBodyProjector::new(BufReader::new(Cursor::new(xml.into_bytes())), None);
        assert!(projector.next_block().unwrap_err().contains("no <w:body>"));
    }

    #[test]
    fn exact_block_projection_limit_is_inclusive() {
        fn paragraph_inner(projected_bytes: usize) -> String {
            let root_bytes = format!(r#"<w:p xmlns:w="{W}"></w:p>"#).len();
            let unit_overhead = "<w:r><w:t></w:t></w:r>".len();
            let mut remaining = projected_bytes - root_bytes;
            let mut inner = String::new();
            while remaining > unit_overhead + 256 * 1024 {
                inner.push_str("<w:r><w:t>");
                inner.push_str(&"x".repeat(256 * 1024));
                inner.push_str("</w:t></w:r>");
                remaining -= unit_overhead + 256 * 1024;
            }
            inner.push_str("<w:r><w:t>");
            inner.push_str(&"x".repeat(remaining - unit_overhead));
            inner.push_str("</w:t></w:r>");
            inner
        }
        let exact_inner = paragraph_inner(HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize);
        let exact = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p>{exact_inner}</w:p></w:body></w:document>"#,
        );
        let mut projector =
            DocumentBodyProjector::new(BufReader::new(Cursor::new(exact.into_bytes())), None);
        let block = projector.next_block().unwrap().unwrap();
        assert_eq!(block.xml.len(), HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize);

        let over_inner = paragraph_inner(HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize + 1);
        let over = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p>{over_inner}</w:p></w:body></w:document>"#,
        );
        let mut projector =
            DocumentBodyProjector::new(BufReader::new(Cursor::new(over.into_bytes())), None);
        assert!(projector
            .next_block()
            .unwrap_err()
            .contains("limit exceeded"));
    }
}

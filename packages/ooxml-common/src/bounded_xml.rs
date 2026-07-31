//! Format-neutral bounded XML streaming primitives.
//!
//! Host-schema state machines remain in the consuming format crates. This
//! module owns only lexical event bounds, source positions, inherited namespace
//! context, bounded event projection, and the namespace-based core of MCE
//! `Choice` selection.

use crate::mce::{
    classify_choice_requires, ChoiceRequiresClassification, RequiredNamespaceSupport,
};
use crate::package_session::PackageLimitReporter;
use crate::resource::{observe_hard_limit, HardResourceLimitKind};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{QName, ResolveResult};
use quick_xml::{NsReader, Writer};
use std::collections::HashSet;
use std::io::{BufRead, Read};
use std::rc::Rc;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XmlSourceSpan {
    pub start: u64,
    pub end: u64,
}

#[derive(Debug)]
pub struct BoundedXmlEvent {
    pub namespace: Option<String>,
    pub event: Event<'static>,
    pub span: XmlSourceSpan,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BoundedXmlReadError {
    Xml(String),
    EventLimit { limit: usize, observed: usize },
}

/// A `BufRead` adapter whose allowance is reset before every quick-xml event.
/// It prevents one unterminated or hostile token from growing quick-xml's
/// caller-owned event buffer to the complete part size.
pub struct EventBudgetReader<R> {
    inner: R,
    limit: usize,
    remaining: usize,
    exceeded: bool,
    current_event_bytes: usize,
    max_event_bytes: usize,
}

impl<R> EventBudgetReader<R> {
    pub fn new(inner: R, limit: usize) -> Self {
        Self {
            inner,
            limit,
            remaining: limit,
            exceeded: false,
            current_event_bytes: 0,
            max_event_bytes: 0,
        }
    }

    fn reset_event_budget(&mut self) {
        self.remaining = self.limit;
        self.exceeded = false;
        self.current_event_bytes = 0;
    }

    pub fn max_event_bytes(&self) -> usize {
        self.max_event_bytes
    }
}

impl<R: BufRead> Read for EventBudgetReader<R> {
    fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
        let available = self.fill_buf()?;
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

impl<R: BufRead> BufRead for EventBudgetReader<R> {
    fn fill_buf(&mut self) -> std::io::Result<&[u8]> {
        if self.remaining == 0 {
            if self.inner.fill_buf()?.is_empty() {
                return Ok(&[]);
            }
            self.exceeded = true;
            return Err(std::io::Error::other(
                "XML event exceeded its hidden byte ceiling",
            ));
        }
        let available = self.inner.fill_buf()?;
        Ok(&available[..available.len().min(self.remaining)])
    }

    fn consume(&mut self, amount: usize) {
        let consumed = amount.min(self.remaining);
        self.inner.consume(consumed);
        self.remaining -= consumed;
        self.current_event_bytes = self.current_event_bytes.saturating_add(consumed);
        self.max_event_bytes = self.max_event_bytes.max(self.current_event_bytes);
    }
}

/// Namespace-aware, source-positioned event reader over one bounded XML source.
pub struct BoundedXmlReader<R: BufRead> {
    reader: NsReader<EventBudgetReader<R>>,
    event_buf: Vec<u8>,
    label: &'static str,
    terminal: Option<BoundedXmlReadError>,
}

impl<R: BufRead> BoundedXmlReader<R> {
    pub fn new(source: R, event_limit: usize, label: &'static str) -> Self {
        let mut reader = NsReader::from_reader(EventBudgetReader::new(source, event_limit));
        reader.config_mut().trim_text(false);
        Self {
            reader,
            event_buf: Vec::with_capacity(event_limit),
            label,
            terminal: None,
        }
    }

    pub fn reader(&self) -> &NsReader<EventBudgetReader<R>> {
        &self.reader
    }

    pub fn max_event_bytes(&self) -> usize {
        self.reader.get_ref().max_event_bytes()
    }

    pub fn read_event(&mut self) -> Result<BoundedXmlEvent, BoundedXmlReadError> {
        if let Some(error) = &self.terminal {
            return Err(error.clone());
        }
        self.reader.get_mut().reset_event_budget();
        let start = self.reader.buffer_position();
        let read = self.reader.read_resolved_event_into(&mut self.event_buf);
        let (namespace, event) = match read {
            Ok(read) => read,
            Err(error) => {
                let terminal = if self.reader.get_ref().exceeded {
                    let limit = self.reader.get_ref().limit;
                    BoundedXmlReadError::EventLimit {
                        limit,
                        observed: limit.saturating_add(1),
                    }
                } else {
                    BoundedXmlReadError::Xml(format!("{} XML stream: {error}", self.label))
                };
                self.terminal = Some(terminal.clone());
                return Err(terminal);
            }
        };
        let namespace = match resolved_namespace_uri(&namespace, self.label) {
            Ok(namespace) => namespace,
            Err(error) => {
                let terminal = BoundedXmlReadError::Xml(error);
                self.terminal = Some(terminal.clone());
                return Err(terminal);
            }
        };
        let event = event.into_owned();
        let end = self.reader.buffer_position();
        self.event_buf.clear();
        Ok(BoundedXmlEvent {
            namespace,
            event,
            span: XmlSourceSpan { start, end },
        })
    }
}

pub fn resolved_namespace_is(
    namespace: &ResolveResult<'_>,
    predicate: impl FnOnce(&str) -> bool,
) -> bool {
    match namespace {
        ResolveResult::Bound(namespace) => std::str::from_utf8(namespace.as_ref())
            .ok()
            .map(predicate)
            .unwrap_or(false),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => false,
    }
}

fn resolved_namespace_uri(
    namespace: &ResolveResult<'_>,
    label: &str,
) -> Result<Option<String>, String> {
    match namespace {
        ResolveResult::Bound(namespace) => {
            let namespace = std::str::from_utf8(namespace.as_ref()).map_err(|e| e.to_string())?;
            quick_xml::escape::unescape(namespace)
                .map(|namespace| Some(namespace.into_owned()))
                .map_err(|e| format!("{label} namespace URI: {e}"))
        }
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(format!(
            "{label} XML uses unbound namespace prefix {}",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BoundedXmlError {
    Xml(String),
    Limit { limit: usize, observed: usize },
}

impl From<String> for BoundedXmlError {
    fn from(error: String) -> Self {
        Self::Xml(error)
    }
}

/// Persistent in-scope namespace declarations. A derived context stores only
/// declarations introduced by one element and shares its parent by `Rc`.
#[derive(Default)]
pub struct NamespaceContext {
    parent: Option<Rc<NamespaceContext>>,
    declarations: Vec<(Option<String>, String)>,
    active_bytes: usize,
}

impl NamespaceContext {
    pub fn root() -> Rc<Self> {
        Rc::new(Self::default())
    }

    pub fn derive(
        element: &BytesStart<'_>,
        inherited: &Rc<Self>,
        context_limit: usize,
        label: &str,
    ) -> Result<Rc<Self>, BoundedXmlError> {
        let mut local_context_bytes = 0usize;
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| {
                BoundedXmlError::Xml(format!("{label} namespace attribute: {error}"))
            })?;
            let key = attribute.key.as_ref();
            if key == b"xmlns" || key.starts_with(b"xmlns:") {
                local_context_bytes = local_context_bytes
                    .checked_add(key.len())
                    .and_then(|total| total.checked_add(attribute.value.len()))
                    .unwrap_or(usize::MAX);
            }
        }
        if local_context_bytes == 0 {
            return Ok(Rc::clone(inherited));
        }
        let active_bytes = inherited.active_bytes.saturating_add(local_context_bytes);
        if active_bytes > context_limit {
            return Err(BoundedXmlError::Limit {
                limit: context_limit,
                observed: active_bytes,
            });
        }
        let mut declarations = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| {
                BoundedXmlError::Xml(format!("{label} namespace attribute: {error}"))
            })?;
            let key = attribute.key.as_ref();
            let prefix = if key == b"xmlns" {
                Some(None)
            } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                Some(Some(
                    std::str::from_utf8(prefix)
                        .map_err(|error| BoundedXmlError::Xml(error.to_string()))?
                        .to_string(),
                ))
            } else {
                None
            };
            if let Some(prefix) = prefix {
                let namespace = attribute
                    .unescape_value()
                    .map_err(|error| BoundedXmlError::Xml(error.to_string()))?
                    .into_owned();
                declarations.push((prefix, namespace));
            }
        }
        Ok(Rc::new(Self {
            parent: Some(Rc::clone(inherited)),
            declarations,
            active_bytes,
        }))
    }

    pub fn active_bytes(&self) -> usize {
        self.active_bytes
    }

    pub fn namespace_for_prefix(&self, prefix: &str) -> Option<&str> {
        self.declarations
            .iter()
            .rev()
            .find_map(|(candidate, namespace)| {
                (candidate.as_deref() == Some(prefix)).then_some(namespace.as_str())
            })
            .or_else(|| {
                self.parent
                    .as_deref()
                    .and_then(|parent| parent.namespace_for_prefix(prefix))
            })
    }

    pub fn effective_bindings(&self) -> Vec<(Option<String>, String)> {
        let mut chain = Vec::new();
        let mut context = Some(self);
        while let Some(current) = context {
            chain.push(current);
            context = current.parent.as_deref();
        }
        let mut bindings: Vec<(Option<String>, String)> = Vec::new();
        for current in chain.into_iter().rev() {
            for (prefix, namespace) in &current.declarations {
                if let Some(existing) = bindings
                    .iter_mut()
                    .find(|(candidate, _)| candidate == prefix)
                {
                    existing.1.clone_from(namespace);
                } else {
                    bindings.push((prefix.clone(), namespace.clone()));
                }
            }
        }
        bindings
    }
}

/// Persistent namespace-resolved MCE state. Parsing MCE attributes and deciding
/// which namespaces a host format understands remain responsibilities of the
/// consuming format crate.
#[derive(Default)]
pub struct MceScope {
    parent: Option<Rc<MceScope>>,
    ignorable: HashSet<String>,
    process_content: HashSet<(String, String)>,
    active_bytes: usize,
}

impl MceScope {
    pub fn root() -> Rc<Self> {
        Rc::new(Self::default())
    }

    pub fn derive(
        inherited: &Rc<Self>,
        ignorable: HashSet<String>,
        process_content: HashSet<(String, String)>,
        derived_bytes: usize,
    ) -> Rc<Self> {
        if ignorable.is_empty() && process_content.is_empty() {
            return Rc::clone(inherited);
        }
        Rc::new(Self {
            parent: Some(Rc::clone(inherited)),
            ignorable,
            process_content,
            active_bytes: inherited.active_bytes.saturating_add(derived_bytes),
        })
    }

    pub fn active_bytes(&self) -> usize {
        self.active_bytes
    }

    pub fn is_ignorable(&self, namespace: &str) -> bool {
        self.ignorable.contains(namespace)
            || self
                .parent
                .as_deref()
                .is_some_and(|parent| parent.is_ignorable(namespace))
    }

    pub fn processes_content(&self, namespace: &str, local_name: &str) -> bool {
        self.process_content
            .iter()
            .any(|(candidate_namespace, candidate_local)| {
                candidate_namespace == namespace
                    && (candidate_local == local_name || candidate_local == "*")
            })
            || self
                .parent
                .as_deref()
                .is_some_and(|parent| parent.processes_content(namespace, local_name))
    }
}

/// Repair namespace declarations when preprocessing removes physical XML
/// ancestors but retains and serializes one of their descendants.
pub fn inject_missing_namespaces(
    element: &mut BytesStart<'static>,
    element_context: &NamespaceContext,
    processed_parent_context: Option<&NamespaceContext>,
    context_limit: usize,
    label: &str,
) -> Result<(), BoundedXmlError> {
    let mut locally_declared = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            BoundedXmlError::Xml(format!("{label} namespace attribute: {error}"))
        })?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" {
            locally_declared.insert(None);
        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
            locally_declared.insert(Some(
                std::str::from_utf8(prefix)
                    .map_err(|error| BoundedXmlError::Xml(error.to_string()))?
                    .to_string(),
            ));
        }
    }
    let element_namespaces = element_context.effective_bindings();
    let processed_parent_namespaces = processed_parent_context
        .map(NamespaceContext::effective_bindings)
        .unwrap_or_default();
    let mut missing = Vec::new();
    let mut injected_bytes = 0usize;
    for (prefix, namespace) in element_namespaces {
        let parent_preserves_binding =
            processed_parent_namespaces
                .iter()
                .any(|(parent_prefix, parent_namespace)| {
                    parent_prefix == &prefix && parent_namespace == &namespace
                });
        if locally_declared.contains(&prefix) || parent_preserves_binding {
            continue;
        }
        let escaped_namespace_bytes = quick_xml::escape::escape(&namespace).len();
        injected_bytes = injected_bytes
            .checked_add(escaped_namespace_bytes)
            .and_then(|total| total.checked_add(prefix.as_ref().map_or(9, |p| p.len() + 10)))
            .unwrap_or(usize::MAX);
        missing.push((prefix, namespace));
    }
    let observed = element.len().saturating_add(injected_bytes);
    if observed > context_limit {
        return Err(BoundedXmlError::Limit {
            limit: context_limit,
            observed,
        });
    }
    // Preflight every declaration before mutating the start tag, so a limit
    // failure cannot leave a partially repaired fragment.
    for (prefix, namespace) in &missing {
        match prefix {
            Some(prefix) => {
                let key = format!("xmlns:{prefix}");
                element.push_attribute((key.as_str(), namespace.as_str()));
            }
            None => element.push_attribute(("xmlns", namespace.as_str())),
        }
    }
    Ok(())
}

pub fn projected_event_bytes(event: &Event<'_>) -> usize {
    match event {
        Event::Start(event) => event.len().saturating_add(2),
        Event::End(event) => event.len().saturating_add(3),
        Event::Empty(event) => event.len().saturating_add(3),
        Event::Text(event) => event.len(),
        Event::Comment(event) => event.len().saturating_add(7),
        Event::CData(event) => event.len().saturating_add(12),
        Event::Decl(event) => event.len().saturating_add(4),
        Event::PI(event) => event.len().saturating_add(4),
        Event::DocType(event) => event.len().saturating_add(11),
        Event::GeneralRef(event) => event.len().saturating_add(2),
        Event::Eof => 0,
    }
}

pub fn append_projected_event(
    target: &mut Vec<u8>,
    event: &Event<'_>,
    retained_overhead: usize,
    limit: usize,
    label: &str,
) -> Result<(), BoundedXmlError> {
    let event_bytes = projected_event_bytes(event);
    let observed = retained_overhead
        .checked_add(target.len())
        .and_then(|current| current.checked_add(event_bytes))
        .unwrap_or(usize::MAX);
    if observed > limit {
        return Err(BoundedXmlError::Limit { limit, observed });
    }
    target
        .try_reserve_exact(event_bytes)
        .map_err(|error| BoundedXmlError::Xml(error.to_string()))?;
    Writer::new(target)
        .write_event(event.borrow())
        .map_err(|error| BoundedXmlError::Xml(format!("{label} XML projection: {error}")))
}

/// Parse and neutrally classify the namespace-only core of MCE `Choice`
/// selection. The host decides whether missing/blank non-conformance is fatal.
pub fn classify_mce_choice_requires(
    choice: &BytesStart<'_>,
    context: &NamespaceContext,
    understands: &dyn Fn(&str) -> bool,
    label: &str,
) -> Result<ChoiceRequiresClassification, String> {
    let mut requires = None;
    for attribute in choice.attributes() {
        let attribute = attribute.map_err(|e| format!("{label} MCE Choice attribute: {e}"))?;
        if attribute.key == QName(b"Requires") {
            requires = Some(
                attribute
                    .unescape_value()
                    .map_err(|e| format!("{label} MCE Requires: {e}"))?
                    .into_owned(),
            );
        }
    }
    Ok(classify_choice_requires(
        requires.as_deref(),
        |prefix| match context.namespace_for_prefix(prefix) {
            None => RequiredNamespaceSupport::Unresolved,
            Some(namespace) if understands(namespace) => RequiredNamespaceSupport::Understood,
            Some(_) => RequiredNamespaceSupport::Unsupported,
        },
    ))
}

/// Attribute a format-owned hard XML ceiling to its active package operation.
pub fn report_hard_limit(
    reporter: Option<&PackageLimitReporter>,
    kind: HardResourceLimitKind,
    part: Option<&str>,
    limit: usize,
    observed: usize,
) -> Result<(), String> {
    let limit = u64::try_from(limit).unwrap_or(u64::MAX);
    let observed = u64::try_from(observed).unwrap_or(u64::MAX);
    if let Some(reporter) = reporter {
        reporter.observe_hard_limit(kind, part, limit, observed)
    } else {
        observe_hard_limit(kind, part, limit, observed)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::events::Event;
    use std::io::{BufReader, Cursor};

    #[test]
    fn event_reader_bounds_one_token_and_reports_source_spans() {
        let xml = br#"<r xmlns:a="urn:a"><a:x>v</a:x></r>"#;
        let source = BufReader::with_capacity(1, Cursor::new(xml.as_slice()));
        let mut reader = BoundedXmlReader::new(source, 64, "test");

        let root = reader.read_event().unwrap();
        assert!(matches!(root.event, Event::Start(_)));
        assert_eq!(root.span.start, 0);
        assert!(root.span.end > root.span.start);
        assert!(reader.event_buf.is_empty());

        let child = reader.read_event().unwrap();
        assert_eq!(child.namespace.as_deref(), Some("urn:a"));
        assert!(child.span.start >= root.span.end);

        let over = BufReader::new(Cursor::new(b"<r>12345</r>".as_slice()));
        let mut over = BoundedXmlReader::new(over, 4, "test");
        assert!(matches!(over.read_event().unwrap().event, Event::Start(_)));
        let terminal = over.read_event().unwrap_err();
        assert!(matches!(
            terminal,
            BoundedXmlReadError::EventLimit {
                limit: 4,
                observed: 5
            }
        ));
        let buffered = over.event_buf.clone();
        let position = over.reader.buffer_position();
        assert!(!buffered.is_empty());
        for _ in 0..3 {
            assert_eq!(over.read_event().unwrap_err(), terminal);
            assert_eq!(over.event_buf, buffered);
            assert_eq!(over.reader.buffer_position(), position);
            assert_eq!(over.max_event_bytes(), 4);
        }
    }

    #[test]
    fn xml_error_is_terminal_and_does_not_reuse_the_event_buffer() {
        let source = BufReader::new(Cursor::new(b"<u:x/>tail".as_slice()));
        let mut reader = BoundedXmlReader::new(source, 64, "test");

        let terminal = reader.read_event().unwrap_err();
        assert!(matches!(terminal, BoundedXmlReadError::Xml(_)));
        let buffered = reader.event_buf.clone();
        let position = reader.reader.buffer_position();
        assert!(!buffered.is_empty());
        for _ in 0..3 {
            assert_eq!(reader.read_event().unwrap_err(), terminal);
            assert_eq!(reader.event_buf, buffered);
            assert_eq!(reader.reader.buffer_position(), position);
        }

        let source = BufReader::new(Cursor::new(b"<r><".as_slice()));
        let mut malformed = BoundedXmlReader::new(source, 64, "test");
        assert!(matches!(
            malformed.read_event().unwrap().event,
            Event::Start(_)
        ));
        let terminal = malformed.read_event().unwrap_err();
        let buffered = malformed.event_buf.clone();
        let position = malformed.reader.buffer_position();
        assert!(matches!(terminal, BoundedXmlReadError::Xml(_)));
        assert_eq!(malformed.read_event().unwrap_err(), terminal);
        assert_eq!(malformed.event_buf, buffered);
        assert_eq!(malformed.reader.buffer_position(), position);
    }

    #[test]
    fn namespace_context_is_persistent_and_rebinding_is_effective() {
        let root = NamespaceContext::root();
        let first = quick_xml::events::BytesStart::from_content(
            r#"r xmlns="urn:default" xmlns:a="urn:a""#,
            1,
        );
        let first = NamespaceContext::derive(&first, &root, 1024, "test").unwrap();
        let second = quick_xml::events::BytesStart::from_content(r#"a:x xmlns:a="urn:other""#, 3);
        let second = NamespaceContext::derive(&second, &first, 1024, "test").unwrap();

        assert_eq!(first.namespace_for_prefix("a"), Some("urn:a"));
        assert_eq!(second.namespace_for_prefix("a"), Some("urn:other"));
        assert_eq!(
            second.effective_bindings(),
            vec![
                (None, "urn:default".to_string()),
                (Some("a".to_string()), "urn:other".to_string())
            ]
        );
    }

    #[test]
    fn event_projection_preflights_the_complete_write() {
        let event = Event::Text(quick_xml::events::BytesText::new("abc"));
        let mut exact = Vec::new();
        append_projected_event(&mut exact, &event, 2, 5, "test").unwrap();
        assert_eq!(exact, b"abc");

        let mut over = Vec::new();
        assert_eq!(
            append_projected_event(&mut over, &event, 3, 5, "test"),
            Err(BoundedXmlError::Limit {
                limit: 5,
                observed: 6
            })
        );
        assert!(over.is_empty());
    }

    #[test]
    fn mce_choice_uses_inherited_prefixes_and_the_host_predicate() {
        let root = NamespaceContext::root();
        let declarations = quick_xml::events::BytesStart::from_content(
            r#"r xmlns:k="urn:known" xmlns:u="urn:unknown""#,
            1,
        );
        let context = NamespaceContext::derive(&declarations, &root, 1024, "test").unwrap();
        let known = quick_xml::events::BytesStart::from_content(r#"mc:Choice Requires="k""#, 9);
        let mixed = quick_xml::events::BytesStart::from_content(r#"mc:Choice Requires="k u""#, 9);
        let missing = quick_xml::events::BytesStart::new("mc:Choice");
        let blank = quick_xml::events::BytesStart::from_content(r#"mc:Choice Requires="   ""#, 9);
        let unbound =
            quick_xml::events::BytesStart::from_content(r#"mc:Choice Requires="missing""#, 9);
        let understands = |namespace: &str| namespace == "urn:known";

        assert_eq!(
            classify_mce_choice_requires(&known, &context, &understands, "test").unwrap(),
            ChoiceRequiresClassification::Understood
        );
        assert_eq!(
            classify_mce_choice_requires(&mixed, &context, &understands, "test").unwrap(),
            ChoiceRequiresClassification::Unsupported
        );
        assert_eq!(
            classify_mce_choice_requires(&missing, &context, &understands, "test").unwrap(),
            ChoiceRequiresClassification::Missing
        );
        assert_eq!(
            classify_mce_choice_requires(&blank, &context, &understands, "test").unwrap(),
            ChoiceRequiresClassification::Blank
        );
        assert_eq!(
            classify_mce_choice_requires(&unbound, &context, &understands, "test").unwrap(),
            ChoiceRequiresClassification::Unresolved
        );

        let bad_escape =
            quick_xml::events::BytesStart::from_content(r#"mc:Choice Requires="&bogus;""#, 9);
        assert!(classify_mce_choice_requires(&bad_escape, &context, &understands, "test").is_err());
        let malformed_attribute =
            quick_xml::events::BytesStart::from_content(r#"mc:Choice Requires="k" broken"#, 9);
        assert!(
            classify_mce_choice_requires(&malformed_attribute, &context, &understands, "test")
                .is_err()
        );
    }

    #[test]
    fn mce_scope_persists_inherited_ignorable_and_process_content_state() {
        let root = MceScope::root();
        let parent = MceScope::derive(
            &root,
            HashSet::from(["urn:future".to_string()]),
            HashSet::new(),
            10,
        );
        let child = MceScope::derive(
            &parent,
            HashSet::new(),
            HashSet::from([("urn:future".to_string(), "keep".to_string())]),
            7,
        );

        assert!(child.is_ignorable("urn:future"));
        assert!(child.processes_content("urn:future", "keep"));
        assert!(!child.processes_content("urn:future", "drop"));
        assert_eq!(child.active_bytes(), 17);
    }

    #[test]
    fn fragment_namespace_repair_preflights_escaped_serialized_bytes() {
        let root = NamespaceContext::root();
        let declaration = BytesStart::from_content(r#"r xmlns:f="urn:a&amp;b""#, 1);
        let context = NamespaceContext::derive(&declaration, &root, 1024, "test").unwrap();
        let mut element = BytesStart::new("f:item");
        let injected_bytes = 10 + "f".len() + "urn:a&amp;b".len();
        let exact_limit = element.len() + injected_bytes;

        inject_missing_namespaces(&mut element, &context, None, exact_limit, "test").unwrap();
        assert_eq!(element.len(), exact_limit);
        assert!(String::from_utf8_lossy(element.as_ref()).contains(r#"xmlns:f="urn:a&amp;b""#));

        let mut over = BytesStart::new("f:item");
        let original_len = over.len();
        assert_eq!(
            inject_missing_namespaces(&mut over, &context, None, exact_limit - 1, "test"),
            Err(BoundedXmlError::Limit {
                limit: exact_limit - 1,
                observed: exact_limit,
            })
        );
        assert_eq!(over.len(), original_len);
    }
}

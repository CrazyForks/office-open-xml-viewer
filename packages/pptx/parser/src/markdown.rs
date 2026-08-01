//! The PPTX → GitHub-flavoured-markdown projection. All public conversion paths
//! feed canonical slides through this one bounded writer and renderer.

use crate::types::*;
use ooxml_common::math::nodes_to_text;
use std::fmt;

/// UTF-8 markdown sink that stops retaining bytes at the configured ceiling.
/// The parser reports the crossing through its package-scoped limit reporter;
/// this type's single responsibility is preventing an oversized projection from
/// first becoming an oversized allocation.
pub(crate) struct MarkdownWriter {
    value: String,
    limit: u64,
    observed: u64,
    exceeded: bool,
}

impl MarkdownWriter {
    pub(crate) fn new(limit: u64) -> Self {
        Self {
            value: String::new(),
            limit,
            observed: 0,
            exceeded: false,
        }
    }

    pub(crate) fn observed(&self) -> u64 {
        self.observed
    }

    pub(crate) fn into_string(self) -> String {
        self.value
    }

    pub(crate) fn push_str(&mut self, value: &str) {
        if self.exceeded {
            return;
        }
        self.observed = self
            .observed
            .saturating_add(u64::try_from(value.len()).unwrap_or(u64::MAX));
        if self.observed > self.limit {
            self.exceeded = true;
            return;
        }
        self.value.push_str(value);
    }

    pub(crate) fn push(&mut self, value: char) {
        let mut encoded = [0; 4];
        self.push_str(value.encode_utf8(&mut encoded));
    }
}

impl fmt::Write for MarkdownWriter {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.push_str(value);
        if self.exceeded {
            Err(fmt::Error)
        } else {
            Ok(())
        }
    }
}

pub(crate) fn render_slide_md(slide: &Slide, out: &mut MarkdownWriter) {
    use std::fmt::Write as _;
    let title = slide_title_md(slide);
    if let Some(t) = title {
        let _ = writeln!(out, "# {} (slide {})\n", t, slide.slide_number);
    } else {
        let _ = writeln!(out, "# Slide {}\n", slide.slide_number);
    }
    for el in &slide.elements {
        render_element_md(el, out);
    }
    if let Some(notes) = &slide.notes {
        let trimmed = notes.trim();
        if !trimmed.is_empty() {
            let _ = writeln!(out, "## Speaker notes\n\n{}\n", trimmed);
        }
    }
    if !slide.comments.is_empty() {
        let _ = writeln!(out, "## Comments\n");
        for c in &slide.comments {
            let author = c.author.as_deref().unwrap_or("(unknown)");
            let _ = writeln!(out, "> **{}**: {}", author, c.text.trim());
        }
        out.push('\n');
    }
}

pub(crate) fn slide_title_md(slide: &Slide) -> Option<String> {
    for el in &slide.elements {
        if let SlideElement::Shape(s) = el {
            let ph = s.placeholder_type.as_deref().unwrap_or("");
            if ph == "title" || ph == "ctrTitle" {
                let txt = shape_text_plain(s);
                if let Some(t) = txt {
                    if !t.is_empty() {
                        return Some(t);
                    }
                }
            }
        }
    }
    None
}

pub(crate) fn shape_text_plain(s: &ShapeElement) -> Option<String> {
    let tb = s.text_body.as_ref()?;
    let mut buf = String::new();
    for para in &tb.paragraphs {
        for run in &para.runs {
            if let TextRun::Text(t) = run {
                buf.push_str(&t.text);
            }
        }
        buf.push(' ');
    }
    let trimmed = buf.trim().to_string();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed)
    }
}

pub(crate) fn render_element_md(el: &SlideElement, out: &mut MarkdownWriter) {
    match el {
        SlideElement::Shape(s) => {
            let ph = s.placeholder_type.as_deref().unwrap_or("");
            // The slide-level # heading already used the title placeholder's
            // text — skip it here to avoid duplicating it inside the body.
            if ph == "title" || ph == "ctrTitle" {
                return;
            }
            // Drop auto-generated metadata placeholders (slide number, date,
            // footer, header). Their text is always a single token like "3" or
            // "2026-05-11" that's pure noise for an agent reading the content.
            if matches!(ph, "sldNum" | "dt" | "ftr" | "hdr") {
                return;
            }
            render_shape_md(s, out);
        }
        SlideElement::Table(t) => render_table_md(t, out),
        SlideElement::Chart(c) => render_chart_md(c, out),
        // Pictures / media / connectors carry no readable text; intentionally
        // dropped in the markdown projection. Use `pptx_get_pictures` or the
        // raw JSON path when you need to inspect them.
        SlideElement::Picture(_) | SlideElement::Media(_) => {}
    }
}

pub(crate) fn render_shape_md(s: &ShapeElement, out: &mut MarkdownWriter) {
    let Some(tb) = &s.text_body else { return };
    if tb.paragraphs.is_empty() {
        return;
    }
    // Body / subtitle placeholders inherit bullet formatting from the layout's
    // lstStyle (ECMA-376 §19.7.10) — treat `Bullet::Inherit` paragraphs there
    // as bulleted, mirroring what PowerPoint draws. Free text boxes default to
    // plain paragraphs.
    let ph = s.placeholder_type.as_deref().unwrap_or("");
    let inherit_means_bullet = matches!(ph, "body" | "subTitle" | "obj" | "tx" | "ftr" | "hdr");
    for para in &tb.paragraphs {
        render_paragraph_md(para, inherit_means_bullet, out);
    }
    out.push('\n');
}

pub(crate) enum ParaKind {
    Plain,
    Bullet,
    Number,
}

pub(crate) fn paragraph_kind(b: &Bullet, inherit_means_bullet: bool) -> ParaKind {
    match b {
        Bullet::None => ParaKind::Plain,
        Bullet::Char { .. } => ParaKind::Bullet,
        // A picture bullet is still an unordered list item for markdown export.
        Bullet::Blip { .. } => ParaKind::Bullet,
        Bullet::AutoNum { .. } => ParaKind::Number,
        Bullet::Inherit => {
            if inherit_means_bullet {
                ParaKind::Bullet
            } else {
                ParaKind::Plain
            }
        }
    }
}

pub(crate) fn render_paragraph_md(
    para: &Paragraph,
    inherit_means_bullet: bool,
    out: &mut MarkdownWriter,
) {
    use std::fmt::Write as _;
    if !runs_have_visible_text(&para.runs) {
        out.push('\n');
        return;
    }
    for _ in 0..para.lvl {
        out.push_str("  ");
    }
    match paragraph_kind(&para.bullet, inherit_means_bullet) {
        ParaKind::Plain => {}
        ParaKind::Bullet => out.push_str("- "),
        // We deliberately emit `1.` for every numbered paragraph rather than
        // tracking the real counter — every markdown renderer auto-renumbers
        // sequential ordered-list items, so the visual output is correct and
        // we don't need to carry per-list state.
        ParaKind::Number => out.push_str("1. "),
    }
    render_runs_md(&para.runs, out);
    let _ = writeln!(out);
}

fn runs_have_visible_text(runs: &[TextRun]) -> bool {
    runs.iter().any(|run| match run {
        TextRun::Break => false,
        TextRun::Math { nodes, .. } => !nodes_to_text(nodes).trim().is_empty(),
        TextRun::Text(text) => !text.text.trim().is_empty(),
    })
}

pub(crate) fn render_runs_md(runs: &[TextRun], out: &mut MarkdownWriter) {
    for run in runs {
        match run {
            // Intra-paragraph soft break (<a:br/>) → markdown hard line break
            // (two trailing spaces + newline).
            TextRun::Break => out.push_str("  \n"),
            // Equations have no faithful markdown form; emit their flattened text.
            TextRun::Math { nodes, .. } => out.push_str(&nodes_to_text(nodes)),
            TextRun::Text(t) => {
                let raw = &t.text;
                // Empty / whitespace-only runs (separators between formatted
                // spans) shouldn't trigger bold/italic wrappers — `**   **`
                // is awkward and most renderers drop the formatting anyway.
                if raw.chars().all(|c| c.is_whitespace()) {
                    out.push_str(raw);
                    continue;
                }
                // Preserve leading/trailing whitespace OUTSIDE the formatting
                // wrappers so `(bold)" Title "` becomes ` **Title** ` not
                // `**" Title "**`. This is how every markdown renderer treats
                // strong/emphasis spans (they're trimmed of whitespace).
                let leading_len = raw.len() - raw.trim_start().len();
                let trail_start = raw.trim_end().len();
                let leading = &raw[..leading_len];
                let trailing = &raw[trail_start..];
                let trimmed = &raw[leading_len..trail_start];
                out.push_str(leading);
                if t.italic == Some(true) {
                    out.push('*');
                }
                if t.bold == Some(true) {
                    out.push_str("**");
                }
                if t.hyperlink.is_some() {
                    out.push('[');
                }
                write_escaped_inline_md(trimmed, out);
                if let Some(url) = &t.hyperlink {
                    out.push_str("](");
                    out.push_str(url);
                    out.push(')');
                }
                if t.bold == Some(true) {
                    out.push_str("**");
                }
                if t.italic == Some(true) {
                    out.push('*');
                }
                out.push_str(trailing);
            }
        }
    }
}

/// Escape the markdown inline metacharacters that would otherwise be parsed as
/// formatting. We deliberately don't escape every potential metachar — pptx
/// body text contains so much punctuation that aggressive escaping makes the
/// output noisier than the structure it's trying to expose. Pipe is handled
/// separately in `render_table_cell_md` since it only matters inside tables.
fn write_escaped_inline_md(value: &str, out: &mut MarkdownWriter) {
    for character in value.chars() {
        if matches!(character, '\\' | '*' | '_' | '`') {
            out.push('\\');
        }
        out.push(character);
    }
}

pub(crate) fn render_table_md(t: &TableElement, out: &mut MarkdownWriter) {
    use std::fmt::Write as _;
    if t.rows.is_empty() {
        return;
    }
    let cols = t.rows[0].cells.len();
    if cols == 0 {
        return;
    }
    let header_cells: Vec<String> = t.rows[0].cells.iter().map(render_table_cell_md).collect();
    let _ = writeln!(out, "| {} |", header_cells.join(" | "));
    let sep: Vec<&str> = (0..cols).map(|_| "---").collect();
    let _ = writeln!(out, "| {} |", sep.join(" | "));
    for row in t.rows.iter().skip(1) {
        let cells: Vec<String> = row.cells.iter().map(render_table_cell_md).collect();
        let _ = writeln!(out, "| {} |", cells.join(" | "));
    }
    out.push('\n');
}

pub(crate) fn render_table_cell_md(cell: &TableCell) -> String {
    // Continuation cells of a merge carry no content — leave empty so the row
    // alignment stays intact.
    if cell.h_merge || cell.v_merge {
        return String::new();
    }
    let Some(tb) = &cell.text_body else {
        return String::new();
    };
    let mut buf = String::new();
    for (i, para) in tb.paragraphs.iter().enumerate() {
        if i > 0 {
            buf.push_str("<br>");
        }
        for run in &para.runs {
            if let TextRun::Text(t) = run {
                buf.push_str(&t.text);
            }
        }
    }
    buf.trim().replace('|', "\\|")
}

pub(crate) fn render_chart_md(c: &ChartElement, out: &mut MarkdownWriter) {
    use std::fmt::Write as _;
    let chart = &c.chart;
    let title = chart.title.as_deref().unwrap_or("(untitled)");
    let _ = writeln!(out, "**Chart ({}): {}**\n", chart.chart_type, title);
    if !chart.categories.is_empty() {
        let _ = writeln!(out, "- Categories: {}", chart.categories.join(", "));
    }
    for s in &chart.series {
        let values: Vec<String> = s
            .values
            .iter()
            .map(|v| match v {
                Some(n) => format!("{n}"),
                None => "—".to_string(),
            })
            .collect();
        let _ = writeln!(out, "- {}: {}", s.name, values.join(", "));
    }
    out.push('\n');
}

/// Materialized-model oracle retained for compatibility/degraded paths and
/// sequential-output equivalence tests.
pub(crate) fn render_presentation_md(pres: &Presentation) -> String {
    let mut out = MarkdownWriter::new(u64::MAX);
    for (i, slide) in pres.slides.iter().enumerate() {
        if i > 0 {
            out.push_str("\n---\n\n");
        }
        render_slide_md(slide, &mut out);
    }
    out.into_string()
}

#[cfg(test)]
mod tests {
    use super::MarkdownWriter;

    #[test]
    fn writer_counts_utf8_bytes_and_stops_retaining_at_the_first_crossing() {
        let mut output = MarkdownWriter::new(3);
        output.push_str("é");
        output.push_str("é");
        output.push_str("ignored after crossing");

        assert_eq!(output.observed(), 4);
        assert_eq!(output.into_string(), "é");
    }
}

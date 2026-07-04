//! Chart element parsing: the legacy DrawingML chart (`c:` namespace) and the
//! newer chartEx (`cx:` namespace) parsers, plus the pptx `ColorResolver` the
//! shared `ooxml_common::chart` helpers use to resolve `<a:solidFill>` colours.
//! Extracted verbatim from `lib.rs`. The general colour grammar
//! (`parse_color_node`) and the shared XML helpers (`child`, `attr`) stay in
//! `lib.rs` and are imported here.

use crate::types::*;
use crate::{attr, parse_color_node};
use std::collections::HashMap;

/// `ooxml_common::chart::ColorResolver` implementation backed by pptx's
/// `HashMap<String, String>` theme palette and PowerPoint's tint formula.
/// Used by chart helpers in ooxml-common that need to resolve
/// `<a:solidFill>` text colors without owning the theme storage.
pub(crate) struct PptxColorResolver<'a> {
    pub(crate) theme: &'a HashMap<String, String>,
}

impl ooxml_common::chart::ColorResolver for PptxColorResolver<'_> {
    fn resolve_solid_fill(&self, node: roxmltree::Node<'_, '_>) -> Option<String> {
        parse_color_node(node, self.theme)
    }

    fn theme_major_font_latin(&self) -> Option<String> {
        // pptx stores the theme major/minor Latin faces under the `+mj-lt` /
        // `+mn-lt` keys of its color+font map (see lib.rs parse_theme_colors).
        self.theme.get("+mj-lt").cloned()
    }

    fn theme_minor_font_latin(&self) -> Option<String> {
        self.theme.get("+mn-lt").cloned()
    }
}

/// Parse a legacy OOXML chart (`c:` namespace) — barChart / lineChart etc.
///
/// Thin pptx adapter over the shared
/// [`ooxml_common::chart::parse_chart_part`]: it builds a [`PptxColorResolver`]
/// from the theme palette, delegates the entire chart-structure parse, and
/// wraps the resulting [`ChartModel`] in a pptx [`ChartElement`] graphic frame.
/// The frame geometry (`x`/`y`/`width`/`height`) is filled in by the caller
/// from the slide's `<p:graphicFrame><a:xfrm>`; here it defaults to 0.
pub(crate) fn parse_legacy_chart(
    xml: &str,
    theme: &HashMap<String, String>,
) -> Option<ChartElement> {
    let doc = roxmltree::Document::parse(xml).ok()?;
    let root = doc.root_element();
    let resolver = PptxColorResolver { theme };
    let chart = ooxml_common::chart::parse_chart_part(root, &resolver)?;
    Some(ChartElement {
        x: 0,
        y: 0,
        width: 0,
        height: 0,
        chart,
    })
}

/// Parse a modern chartEx (cx: namespace) — waterfall, treemap, etc.
pub(crate) fn parse_chartex(xml: &str, theme: &HashMap<String, String>) -> Option<ChartElement> {
    let doc = roxmltree::Document::parse(xml).ok()?;
    let root = doc.root_element();

    // Chart type from series layoutId attribute
    let series_node = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "series")?;
    let layout_id = attr(&series_node, "layoutId").unwrap_or_default();
    let chart_type = layout_id; // "waterfall", "treemap", etc.

    // Categories from chartData > data > strDim[@type="cat"] > lvl > pt
    let categories: Vec<String> = root
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "strDim"
                && attr(n, "type").as_deref() == Some("cat")
        })
        .and_then(|dim| {
            dim.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "lvl")
        })
        .map(|lvl| {
            lvl.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                .filter_map(|pt| pt.text().map(|t| t.replace('\n', " ")))
                .collect()
        })
        .unwrap_or_default();

    let pt_count = categories.len().max(1);

    // Values from chartData > data > numDim[@type="val"] > lvl > pt
    let raw_values: Vec<Option<f64>> = root
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "numDim"
                && attr(n, "type").as_deref() == Some("val")
        })
        .and_then(|dim| {
            dim.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "lvl")
        })
        .map(|lvl| {
            let mut vals: Vec<Option<f64>> = vec![None; pt_count];
            for (i, pt) in lvl
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                .enumerate()
            {
                if i < vals.len() {
                    vals[i] = pt.text().and_then(|t| t.parse().ok());
                }
            }
            vals
        })
        .unwrap_or_else(|| vec![None; pt_count]);

    // Subtotal indices (idx=0 is always implicit; add from cx:subtotals)
    let mut subtotal_indices: Vec<u32> = vec![0];
    if let Some(subtotals_node) = series_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "subtotals")
    {
        for idx_node in subtotals_node
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "idx")
        {
            if let Some(v) = attr(&idx_node, "val").and_then(|v| v.parse::<u32>().ok()) {
                if v != 0 {
                    subtotal_indices.push(v);
                }
            }
        }
    }

    // Series color (first dataPt or series spPr)
    let color = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(|sp| {
            sp.children()
                .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
        })
        .and_then(|fill| parse_color_node(fill, theme));

    // Per-idx data-label colors. ChartEx writes `<cx:dataLabels>` with
    // `<cx:dataLabel idx="N">` overrides; each carries its own `<cx:txPr>`
    // whose first `<a:solidFill>` is the label colour for that bar. Sample-2
    // waterfall uses this to paint negative-bar labels in accent1 (red) while
    // positive-bar labels stay tx1 (black).
    let mut data_label_colors_vec: Vec<Option<String>> = vec![None; raw_values.len().max(1)];
    let mut has_per_label_color = false;
    for dl in series_node
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dataLabel")
    {
        let Some(idx) = attr(&dl, "idx").and_then(|v| v.parse::<usize>().ok()) else {
            continue;
        };
        if idx >= data_label_colors_vec.len() {
            continue;
        }
        // First `<a:solidFill>` inside the per-idx <cx:txPr>.
        let txpr = match dl
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "txPr")
        {
            Some(n) => n,
            None => continue,
        };
        for desc in txpr.descendants().filter(|n| n.is_element()) {
            if desc.tag_name().name() != "solidFill" {
                continue;
            }
            if let Some(c) = parse_color_node(desc, theme) {
                data_label_colors_vec[idx] = Some(c);
                has_per_label_color = true;
                break;
            }
        }
    }

    let series = vec![ChartSeriesData {
        name: String::new(),
        values: raw_values,
        color,
        data_point_colors: None,
        data_label_colors: if has_per_label_color {
            Some(data_label_colors_vec)
        } else {
            None
        },
        categories: None,
        bubble_sizes: None,
        val_format_code: None,
        label_color: None,
        series_type: None,
        use_secondary_axis: None,
        show_marker: None,
        marker_symbol: None,
        marker_size: None,
        marker_fill: None,
        marker_line: None,
        data_point_overrides: None,
        data_label_overrides: None,
        series_data_labels: None,
        err_bars: None,
        // chartEx (waterfall) has no `<c:smooth>` concept.
        smooth: None,
        // chartEx series carry no classic `<c:trendline>`.
        trend_lines: None,
    }];

    // ChartEx axis visibility — shared helper that pairs each `<cx:axis hidden>`
    // with its `<cx:catScaling>` / `<cx:valScaling>` child to disambiguate cat
    // vs. val (chartEx doesn't declare axis kind via the `id` attribute).
    let (cat_axis_hidden, val_axis_hidden) = ooxml_common::chart::extract_chartex_axis_hidden(root);

    // `<cx:catScaling gapWidth>` (chartEx) — same semantics as legacy
    // `<c:gapWidth>` but stored as a *fraction* (e.g. 0.8 ≡ 80%) instead of
    // an integer percentage. Convert to the legacy percentage form so the
    // shared renderer's `barW = catGap / (1 + gapWidth/100)` formula works
    // uniformly across chart types. Default 1.5 (= legacy 150%) per PowerPoint
    // when the attribute is omitted.
    let bar_gap_width = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "catScaling")
        .and_then(|n| attr(&n, "gapWidth"))
        .and_then(|v| v.parse::<f64>().ok())
        .map(|frac| (frac * 100.0).round() as i32);

    Some(ChartElement {
        x: 0,
        y: 0,
        width: 0,
        height: 0,
        chart: ChartModel {
            chart_type,
            title: None,
            categories,
            series,
            val_max: None,
            val_min: None,
            subtotal_indices,
            show_data_labels: false,
            cat_axis_hidden,
            val_axis_hidden,
            plot_area_bg: None,
            chart_bg: None,
            show_legend: false,
            cat_axis_cross_between: "between".to_string(),
            val_axis_major_tick_mark: "cross".to_string(),
            cat_axis_major_tick_mark: "cross".to_string(),
            title_font_size_hpt: None,
            title_font_color: None,
            title_font_face: None,
            cat_axis_font_size_hpt: None,
            val_axis_font_size_hpt: None,
            cat_axis_font_color: None,
            val_axis_font_color: None,
            cat_axis_line_color: None,
            cat_axis_line_width_emu: None,
            cat_axis_line_hidden: false,
            val_axis_line_color: None,
            val_axis_line_width_emu: None,
            val_axis_line_hidden: false,
            data_label_font_size_hpt: None,
            legend_pos: None,
            bar_gap_width,
            bar_overlap: None,
            data_label_position: None,
            data_label_font_color: None,
            data_label_format_code: None,
            val_axis_format_code: None,
            plot_area_manual_layout: None,
            scatter_style: None,
            // chartEx (waterfall/treemap/etc.) has its own axis model and is not
            // wired for axis titles or an explicit chartSpace border yet.
            cat_axis_title: None,
            val_axis_title: None,
            cat_axis_title_font_size_hpt: None,
            cat_axis_title_font_bold: None,
            cat_axis_title_font_color: None,
            val_axis_title_font_size_hpt: None,
            val_axis_title_font_bold: None,
            val_axis_title_font_color: None,
            title_font_bold: None,
            cat_axis_font_bold: None,
            val_axis_font_bold: None,
            chart_border_color: None,
            chart_border_width_emu: None,
            secondary_val_axis: None,
            // chartEx charts (waterfall/treemap/etc.) are not pie/doughnut and
            // don't carry `<c:txPr>` axis/legend faces; only the theme fallback
            // fonts are threaded so their data labels can pick up the body font.
            hole_size: None,
            first_slice_angle: None,
            cat_axis_font_face: None,
            val_axis_font_face: None,
            cat_axis_title_font_face: None,
            val_axis_title_font_face: None,
            data_label_font_face: None,
            legend_font_face: None,
            legend_font_color: None,
            legend_font_size_hpt: None,
            legend_font_bold: None,
            theme_major_font_latin: theme.get("+mj-lt").cloned(),
            theme_minor_font_latin: theme.get("+mn-lt").cloned(),
            val_axis_minor_tick_mark: None,
            cat_axis_minor_tick_mark: None,
            legend_manual_layout: None,
            title_manual_layout: None,
            cat_axis_crosses: None,
            cat_axis_crosses_at: None,
            val_axis_crosses: None,
            val_axis_crosses_at: None,
            cat_axis_format_code: None,
            cat_axis_min: None,
            cat_axis_max: None,
            radar_style: None,
            // chartEx (cx: namespace) has its own date-axis model; the legacy
            // `<c:date1904>` element does not apply here, so keep the 1900
            // default until/unless a chartEx date system is wired.
            date1904: false,
            // chartEx waterfall has no line/area blanks to display.
            disp_blanks_as: None,
            // chartEx (cx:) has its own axis model (`<cx:axis>`); the classic
            // `<c:catAx>`/`<c:valAx>` scale properties don't apply, so leave the
            // CH6 fields unset — the renderer keeps its defaults (byte-stable).
            val_axis_major_gridlines: None,
            cat_axis_major_gridlines: None,
            val_axis_gridline_color: None,
            val_axis_gridline_width_emu: None,
            cat_axis_gridline_color: None,
            cat_axis_gridline_width_emu: None,
            val_axis_minor_gridlines: None,
            val_axis_major_unit: None,
            val_axis_minor_unit: None,
            val_axis_log_base: None,
            val_axis_orientation: None,
            cat_axis_orientation: None,
            cat_axis_tick_label_pos: None,
            val_axis_tick_label_pos: None,
            cat_axis_label_rotation: None,
        },
    })
}

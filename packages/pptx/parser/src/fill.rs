//! Fill / colour / stroke / effect / 3-D scene / custom-geometry parsing
//! (pptx-specific DrawingML). Extracted verbatim from `lib.rs`. The general
//! colour-node grammar (`parse_color_node` / `parse_color_node_tint`) lives here;
//! it uses the `PptxSchemeResolver` (from `theme`) for `<a:schemeClr>` lookups.
//! Shared XML helpers (`child`, `children_vec`, `attr`, `attr_r`, `attr_i64`,
//! `attr_f64`) stay in `lib.rs` and are imported here.

use crate::theme::PptxSchemeResolver;
use crate::types::*;
use crate::{attr, attr_f64, attr_i64, attr_r, child, children_vec, parse_preflighted_pptx_xml};
use ooxml_common::blip::{mime_from_ext, parse_blip_duotone};
use ooxml_common::color::ThemeResolver;
use std::collections::HashMap;

/// Parse `<a:blip><a:alphaModFix amt="..."/></a:blip>` from a blipFill node
/// (ECMA-376 §20.1.8.6). Thin re-export of the shared
/// [`ooxml_common::blip::parse_blip_alpha`] so the three formats read the blip
/// alpha identically (previously a pptx-local copy). Returns the fraction
/// `amt/100000` when present and < 1.0; `None` otherwise.
pub(crate) use ooxml_common::blip::parse_blip_alpha;
pub(crate) use ooxml_common::fill::{parse_fill_rect, parse_tile};

// ===========================
//  Color parsing
// ===========================

/// Resolve a color node (solidFill child / run rPr child) to a hex string.
/// Handles srgbClr, sysClr, prstClr, and schemeClr (with transform support).
pub(crate) fn parse_color_node(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<String> {
    parse_color_node_tint(node, theme, ooxml_common::color::TintMode::PowerPointLinear)
}

/// Like `parse_color_node`, but lets the caller pick how `<a:tint>` is interpreted.
/// Table styles (`<a:tcStyle>` band fills) use `TintMode::WordLiteral` — the literal
/// ECMA-376 §20.1.2.3.34 definition (`val·input + (1-val)·white`, so a 20% tint is a
/// near-white wash) — which is how PowerPoint renders table band tints. The SmartArt
/// accent-recolor path keeps `PowerPointLinear` (see `apply_color_transforms`).
///
/// Thin wrapper over the shared [`ooxml_common::color::parse_color_node`]: the
/// grammar + transforms live there; [`PptxSchemeResolver`] supplies the
/// pptx-specific theme-slot lookup. Output is unchanged (uppercase hex, no `#`).
pub(crate) fn parse_color_node_tint(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    tint_mode: ooxml_common::color::TintMode,
) -> Option<String> {
    ooxml_common::color::parse_color_node(node, &PptxSchemeResolver { theme }, tint_mode)
}

// ===========================
//  Fill / Stroke parsing
// ===========================

pub(crate) fn parse_fill(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Fill> {
    parse_fill_tint(node, theme, ooxml_common::color::TintMode::PowerPointLinear)
}

/// Parse DrawingML fill properties with the caller-selected tint semantics.
/// Normal presentation backgrounds and theme style-matrix fills use the
/// literal ECMA-376 retained-input definition. The historical generic path
/// keeps PowerPointLinear for SmartArt compatibility.
fn parse_fill_tint(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    tint_mode: ooxml_common::color::TintMode,
) -> Option<Fill> {
    parse_fill_with_resolver(node, &PptxSchemeResolver { theme }, tint_mode)
}

fn parse_fill_with_resolver<R: ThemeResolver + ?Sized>(
    node: roxmltree::Node<'_, '_>,
    resolver: &R,
    tint_mode: ooxml_common::color::TintMode,
) -> Option<Fill> {
    for c in node.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "solidFill" => {
                // If the color resolves, use it. If not (e.g. phClr with no theme slot),
                // return None so the caller can fall back to the shape style color.
                if let Some(color) = ooxml_common::color::parse_color_node(c, resolver, tint_mode) {
                    return Some(Fill::Solid { color });
                }
                // Unresolvable → don't default to black; let fallback logic handle it
            }
            "noFill" => return Some(Fill::None),
            "pattFill" => {
                // ECMA-376 §20.1.8.40 — preset pattern with fg/bg colours.
                // Shared parse (ooxml_common::fill); colors resolve with pptx's
                // PowerPointLinear tint via PptxSchemeResolver.
                let ooxml_common::fill::PatternFill { fg, bg, preset } =
                    ooxml_common::fill::parse_patt_fill(c, resolver, tint_mode);
                return Some(Fill::Pattern { fg, bg, preset });
            }
            "gradFill" => {
                // Shared parse (ooxml_common::fill). Returns None when there are
                // no resolvable stops, so we keep scanning sibling fill elements.
                if let Some(g) = ooxml_common::fill::parse_grad_fill(c, resolver, tint_mode) {
                    return Some(Fill::Gradient {
                        stops: g.stops,
                        angle: g.angle,
                        grad_type: g.grad_type,
                    });
                }
            }
            _ => {}
        }
    }
    None
}

struct StyleMatrixSchemeResolver<'a> {
    theme: &'a HashMap<String, String>,
    placeholder_color: Option<&'a str>,
}

impl ThemeResolver for StyleMatrixSchemeResolver<'_> {
    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        if name == "phClr" {
            return self.placeholder_color.map(str::to_owned);
        }
        PptxSchemeResolver { theme: self.theme }.resolve_scheme_color(name)
    }
}

/// Resolve a shape `fillRef` or slide/master `bgRef` through the theme's format
/// style matrix. `phClr` inside the selected style is substituted with the
/// reference element's colour before its own transforms are applied.
///
/// ECMA-376 Part 1 §19.3.1.3: bgRef 1..999 indexes fillStyleLst, 1001+
/// indexes bgFillStyleLst (1001 = first); 0 and 1000 mean no background.
pub(crate) fn parse_style_matrix_fill(
    style_ref: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    background: bool,
) -> Option<Fill> {
    use ooxml_common::color::TintMode::WordLiteral;

    let idx = attr(&style_ref, "idx")?.parse::<u32>().ok()?;
    let key = if background {
        match idx {
            0 | 1000 => return Some(Fill::None),
            1..=999 => format!("+fillStyle-{idx}"),
            _ => format!("+bgFillStyle-{}", idx - 1000),
        }
    } else if idx == 0 {
        return Some(Fill::None);
    } else {
        format!("+fillStyle-{idx}")
    };
    let fragment = theme.get(&key)?;

    let placeholder_color = parse_color_node_tint(style_ref, theme, WordLiteral);

    let wrapped = format!(
        r#"<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{fragment}</root>"#
    );
    let doc = parse_preflighted_pptx_xml(&wrapped).ok()?;
    let resolver = StyleMatrixSchemeResolver {
        theme,
        placeholder_color: placeholder_color.as_deref(),
    };
    parse_fill_with_resolver(doc.root_element(), &resolver, WordLiteral)
}

/// ECMA-376 §20.1.8.14 `a:blipFill` → `Fill::Image`. The `resolve_blip`
/// closure maps the `<a:blip r:embed>` rId to the blip's embedded **zip path**
/// using the caller's rels (each inheritance level resolves against its own
/// part); the mime is derived from that path. The renderer fetches the bytes
/// lazily by path rather than from an inlined data URL.
///
/// Both fill-modes are honoured and mutually exclusive:
/// - `stretch` (§20.1.8.56): the `fillRect` (§20.1.8.30) is captured so the
///   renderer can place the (possibly overscanned) image into the box.
/// - `tile` (§20.1.8.58): the tile offset/scale/flip/align descriptor is
///   captured so the renderer can repeat the blip at its native (scaled) size.
///
/// When neither child is present the blip defaults to full-box placement
/// (stretch with no fillRect).
///
/// `theme` resolves the `<a:duotone>` (§20.1.8.23) endpoint colours through the
/// slide palette (PowerPoint linear tint), so a picture FILL recolours exactly
/// like a `<p:pic>` picture element does.
pub(crate) fn parse_blip_fill<F: FnMut(&str) -> Option<String>>(
    blip_fill: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> Option<Fill> {
    let r_id = child(blip_fill, "blip").and_then(|b| attr_r(&b, "embed"))?;
    let image_path = resolve_blip(&r_id)?;
    let mime_type = mime_from_ext(&image_path).to_owned();
    let alpha = parse_blip_alpha(blip_fill);
    // §20.1.8.23 duotone recolour, resolved through the theme with PowerPoint's
    // linear tint (same call the `<p:pic>` paths use). `None` ⇒ no effect.
    let duotone = parse_blip_duotone(
        blip_fill,
        &PptxSchemeResolver { theme },
        ooxml_common::color::TintMode::PowerPointLinear,
    );
    // §20.1.8.58 tile takes precedence when present (stretch/tile are an
    // either-or choice in CT_BlipFillProperties).
    if let Some(tile_node) = child(blip_fill, "tile") {
        return Some(Fill::Image {
            image_path,
            mime_type,
            fill_rect: None,
            tile: Some(parse_tile(tile_node)),
            alpha,
            duotone,
        });
    }
    let fill_rect = child(blip_fill, "stretch").and_then(parse_fill_rect);
    Some(Fill::Image {
        image_path,
        mime_type,
        fill_rect,
        tile: None,
        alpha,
        duotone,
    })
}

pub(crate) fn parse_arrow_end(node: roxmltree::Node<'_, '_>) -> ArrowEnd {
    let kind = attr(&node, "type").unwrap_or_else(|| "none".to_owned());
    let w = attr(&node, "w").unwrap_or_else(|| "med".to_owned());
    let len = attr(&node, "len").unwrap_or_else(|| "med".to_owned());
    ArrowEnd { kind, w, len }
}

pub(crate) fn parse_stroke(
    ln_node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Stroke> {
    if child(ln_node, "noFill").is_some() {
        return None;
    }
    let width = attr_i64(&ln_node, "w").unwrap_or(9525);
    // CT_LineProperties uses EG_LineFillProperties (§20.1.8.38), so a line can
    // carry the same solid/gradient/pattern paints as a shape. Keep a solid
    // fallback colour for arrowheads and consumers that do not yet understand
    // non-solid line paint.
    let parsed_fill = parse_fill(ln_node, theme)?;
    let color = match &parsed_fill {
        Fill::Solid { color } => color.clone(),
        Fill::Gradient { stops, .. } => stops
            .iter()
            .rev()
            .find(|stop| !stop.color.ends_with("00"))
            .or_else(|| stops.last())
            .map(|stop| stop.color.clone())?,
        Fill::Pattern { fg, .. } => fg.clone(),
        Fill::None | Fill::Image { .. } => return None,
    };
    let fill = match parsed_fill {
        Fill::Gradient { .. } | Fill::Pattern { .. } => Some(parsed_fill),
        Fill::Solid { .. } => None,
        Fill::None | Fill::Image { .. } => unreachable!(),
    };
    let dash_style = child(ln_node, "prstDash")
        .and_then(|n| attr(&n, "val"))
        .filter(|v| v != "solid");
    let line_cap = attr(&ln_node, "cap").and_then(|cap| match cap.as_str() {
        "rnd" => Some("round".to_owned()),
        "sq" => Some("square".to_owned()),
        "flat" => Some("butt".to_owned()),
        _ => None,
    });
    // Arrow ends — only emit when type != "none"
    let head_end = child(ln_node, "headEnd")
        .map(parse_arrow_end)
        .filter(|a| a.kind != "none");
    let tail_end = child(ln_node, "tailEnd")
        .map(parse_arrow_end)
        .filter(|a| a.kind != "none");
    // ECMA-376 §20.1.8.42 ST_CompoundLine. Default "sng" stays absent so the
    // renderer keeps its single-stroke fast path.
    let cmpd = attr(&ln_node, "cmpd").filter(|v| v != "sng");
    Some(Stroke {
        color,
        width,
        fill,
        dash_style,
        line_cap,
        head_end,
        tail_end,
        cmpd,
    })
}

// ===========================
//  Shadow parsing
// ===========================

/// Parse spPr > effectLst > outerShdw into a Shadow.
pub(crate) fn parse_shadow(
    effect_lst: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Shadow> {
    parse_shadow_node(child(effect_lst, "outerShdw")?, theme)
}

/// Parse spPr > effectLst > innerShdw into a Shadow. ECMA-376 §20.1.8.21
/// (CT_InnerShadowEffect) — same field shape as outerShdw, semantics differ
/// at render time (cast inward).
#[cfg(test)]
pub(crate) fn parse_inner_shadow(
    effect_lst: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Shadow> {
    parse_shadow_node(child(effect_lst, "innerShdw")?, theme)
}

/// Shared field reader for innerShdw / outerShdw. Both elements expose
/// blurRad, dist, dir, and a color child with optional alphaModFix.
pub(crate) fn parse_shadow_node(
    n: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Shadow> {
    parse_shadow_node_with_resolver(
        n,
        &PptxSchemeResolver { theme },
        ooxml_common::color::TintMode::PowerPointLinear,
    )
}

fn parse_shadow_node_with_resolver<R: ThemeResolver + ?Sized>(
    n: roxmltree::Node<'_, '_>,
    resolver: &R,
    tint_mode: ooxml_common::color::TintMode,
) -> Option<Shadow> {
    let blur = attr_i64(&n, "blurRad").unwrap_or(0);
    let dist = attr_i64(&n, "dist").unwrap_or(0);
    let dir = attr_f64(&n, "dir").unwrap_or(0.0) / 60_000.0;

    let color_str = ooxml_common::color::parse_color_node(n, resolver, tint_mode)
        .unwrap_or_else(|| "000000".to_owned());
    let (color, alpha) = if color_str.len() >= 8 {
        let a = u8::from_str_radix(&color_str[6..8], 16).unwrap_or(255) as f64 / 255.0;
        (color_str[..6].to_owned(), a)
    } else {
        (color_str, 1.0)
    };

    Some(Shadow {
        color,
        alpha,
        blur,
        dist,
        dir,
    })
}

/// Parse spPr > effectLst > glow into a Glow effect — ECMA-376 §20.1.8.17
/// (CT_GlowEffect): a coloured halo with a blur radius, no offset.
#[cfg(test)]
pub(crate) fn parse_glow(
    effect_lst: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Glow> {
    let g = child(effect_lst, "glow")?;
    parse_glow_node_with_resolver(
        g,
        &PptxSchemeResolver { theme },
        ooxml_common::color::TintMode::PowerPointLinear,
    )
}

fn parse_glow_node_with_resolver<R: ThemeResolver + ?Sized>(
    g: roxmltree::Node<'_, '_>,
    resolver: &R,
    tint_mode: ooxml_common::color::TintMode,
) -> Option<Glow> {
    let radius = attr_i64(&g, "rad").unwrap_or(0);
    let color_str = ooxml_common::color::parse_color_node(g, resolver, tint_mode)
        .unwrap_or_else(|| "000000".to_owned());
    let (color, alpha) = if color_str.len() >= 8 {
        let a = u8::from_str_radix(&color_str[6..8], 16).unwrap_or(255) as f64 / 255.0;
        (color_str[..6].to_owned(), a)
    } else {
        (color_str, 1.0)
    };
    Some(Glow {
        color,
        alpha,
        radius,
    })
}

/// Parse spPr > effectLst > softEdge into a SoftEdge — ECMA-376 §20.1.8.31.
pub(crate) fn parse_soft_edge(effect_lst: roxmltree::Node<'_, '_>) -> Option<SoftEdge> {
    let n = child(effect_lst, "softEdge")?;
    let radius = attr_i64(&n, "rad").unwrap_or(0);
    Some(SoftEdge { radius })
}

/// Parse spPr > effectLst > reflection — ECMA-376 §20.1.8.27. Defaults
/// follow the spec table: blur=0, dist=0, dir=0, stA=100000 (=1.0),
/// stPos=0, endA=0, endPos=100000 (=1.0), sx=100000, sy=-100000.
pub(crate) fn parse_reflection(effect_lst: roxmltree::Node<'_, '_>) -> Option<Reflection> {
    let r = child(effect_lst, "reflection")?;
    let pct = |name: &str, default: f64| -> f64 {
        attr_f64(&r, name).map(|v| v / 100_000.0).unwrap_or(default)
    };
    Some(Reflection {
        blur: attr_i64(&r, "blurRad").unwrap_or(0),
        dist: attr_i64(&r, "dist").unwrap_or(0),
        dir: attr_f64(&r, "dir").unwrap_or(0.0) / 60_000.0,
        st_a: pct("stA", 1.0),
        st_pos: pct("stPos", 0.0),
        end_a: pct("endA", 0.0),
        end_pos: pct("endPos", 1.0),
        sx: pct("sx", 1.0),
        sy: pct("sy", -1.0),
    })
}

/// Effects pulled from `spPr > effectLst`. The five members are independent
/// siblings inside `CT_EffectList` — ECMA-376 §20.1.8.16. Used by both shapes
/// (`p:sp`) and pictures (`p:pic`): `p:spPr` is `CT_ShapeProperties` in both
/// cases (§19.3.1.37), so `effectLst` applies equally to images.
#[derive(Default)]
pub(crate) struct EffectLst {
    pub(crate) shadow: Option<Shadow>,
    pub(crate) inner_shadow: Option<Shadow>,
    pub(crate) glow: Option<Glow>,
    pub(crate) soft_edge: Option<SoftEdge>,
    pub(crate) reflection: Option<Reflection>,
}

/// Read every `effectLst` child shapes and pictures share. `effect_lst` is the
/// optional `<a:effectLst>` node; missing nodes yield an all-`None` result.
pub(crate) fn parse_effect_lst(
    effect_lst: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
) -> EffectLst {
    parse_effect_lst_with_resolver(
        effect_lst,
        &PptxSchemeResolver { theme },
        ooxml_common::color::TintMode::PowerPointLinear,
    )
}

fn parse_effect_lst_with_resolver<R: ThemeResolver + ?Sized>(
    effect_lst: Option<roxmltree::Node<'_, '_>>,
    resolver: &R,
    tint_mode: ooxml_common::color::TintMode,
) -> EffectLst {
    EffectLst {
        shadow: effect_lst
            .and_then(|node| child(node, "outerShdw"))
            .and_then(|node| parse_shadow_node_with_resolver(node, resolver, tint_mode)),
        inner_shadow: effect_lst
            .and_then(|node| child(node, "innerShdw"))
            .and_then(|node| parse_shadow_node_with_resolver(node, resolver, tint_mode)),
        glow: effect_lst
            .and_then(|node| child(node, "glow"))
            .and_then(|node| parse_glow_node_with_resolver(node, resolver, tint_mode)),
        soft_edge: effect_lst.and_then(parse_soft_edge),
        reflection: effect_lst.and_then(parse_reflection),
    }
}

/// Resolve `p:style/a:effectRef` through the theme format matrix.
///
/// `effectRef@idx` is one-based into `a:effectStyleLst`. Any `phClr` inside
/// that effect style is supplied by the color child of the reference before
/// the ordinary DrawingML transforms are applied.
pub(crate) fn parse_style_matrix_effects(
    effect_ref: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> EffectLst {
    let Some(idx) = attr(&effect_ref, "idx").and_then(|value| value.parse::<u32>().ok()) else {
        return EffectLst::default();
    };
    if idx == 0 {
        return EffectLst::default();
    }
    let Some(fragment) = theme.get(&format!("+effectStyle-{idx}")) else {
        return EffectLst::default();
    };

    let placeholder_color = parse_color_node_tint(
        effect_ref,
        theme,
        ooxml_common::color::TintMode::WordLiteral,
    );
    let wrapped = format!(
        r#"<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{fragment}</root>"#
    );
    let Ok(doc) = parse_preflighted_pptx_xml(&wrapped) else {
        return EffectLst::default();
    };
    let effect_lst = doc
        .root_element()
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "effectLst");
    let resolver = StyleMatrixSchemeResolver {
        theme,
        placeholder_color: placeholder_color.as_deref(),
    };
    parse_effect_lst_with_resolver(
        effect_lst,
        &resolver,
        ooxml_common::color::TintMode::WordLiteral,
    )
}

// ===========================
//  3D scene parsing (scene3d / sp3d)
// ===========================

/// Parse `<a:rot>` (`CT_SphereCoords`, ECMA-376 §20.1.5.11). Angles are stored
/// in the XML as 60000ths of a degree; we convert to degrees. All three
/// attributes are required by the schema, but we default missing ones to 0 to
/// stay tolerant of malformed input.
pub(crate) fn parse_rot3d(rot: roxmltree::Node<'_, '_>) -> Rot3d {
    let deg = |name: &str| attr_f64(&rot, name).unwrap_or(0.0) / 60_000.0;
    Rot3d {
        lat: deg("lat"),
        lon: deg("lon"),
        rev: deg("rev"),
    }
}

/// Parse `<a:scene3d>` (`CT_Scene3D`, ECMA-376 §20.1.4.1.41). Requires a
/// `<a:camera>` child (§20.1.5.5); `<a:lightRig>` is optional for our purposes
/// (Phase A renders the camera only). Returns None when no camera is present.
pub(crate) fn parse_scene3d(sppr: roxmltree::Node<'_, '_>) -> Option<Scene3d> {
    let scene = child(sppr, "scene3d")?;
    let cam = child(scene, "camera")?;
    let camera = Camera3d {
        prst: attr(&cam, "prst")?,
        // §20.1.5.5: fov is an ST_FOVAngle in 60000ths of a degree.
        fov: attr_f64(&cam, "fov").map(|v| v / 60_000.0),
        // zoom is an ST_PositivePercentage (100000 = 100%).
        zoom: attr_f64(&cam, "zoom").map(|v| v / 100_000.0),
        rot: child(cam, "rot").map(parse_rot3d),
    };
    let light_rig = child(scene, "lightRig").and_then(|lr| {
        Some(LightRig {
            rig: attr(&lr, "rig")?,
            dir: attr(&lr, "dir")?,
            rot: child(lr, "rot").map(parse_rot3d),
        })
    });
    Some(Scene3d { camera, light_rig })
}

/// Parse `<a:bevel>` (`CT_Bevel`, ECMA-376 §20.1.5.3). `w`/`h` default to
/// 76200 EMU and `prst` to "circle" per the schema.
pub(crate) fn parse_bevel3d(bevel: roxmltree::Node<'_, '_>) -> Bevel3d {
    Bevel3d {
        w: attr_i64(&bevel, "w").unwrap_or(76_200),
        h: attr_i64(&bevel, "h").unwrap_or(76_200),
        prst: attr(&bevel, "prst").unwrap_or_else(|| "circle".into()),
    }
}

/// Parse `<a:sp3d>` (`CT_Shape3D`, ECMA-376 §20.1.5.12). Defaults follow the
/// schema: z=0, extrusionH=0, contourW=0, prstMaterial="warmMatte". Parsed in
/// full but not rendered in Phase A.
pub(crate) fn parse_sp3d(sppr: roxmltree::Node<'_, '_>) -> Option<Sp3d> {
    let n = child(sppr, "sp3d")?;
    // contourClr is colour-only here; pass an empty theme map because sp3d
    // contour colours in practice are srgbClr (no theme lookup needed) and this
    // parser has the theme threaded only into the line/fill paths.
    let contour_clr = child(n, "contourClr").and_then(|c| parse_color_node(c, &HashMap::new()));
    Some(Sp3d {
        z: attr_i64(&n, "z").unwrap_or(0),
        extrusion_h: attr_i64(&n, "extrusionH").unwrap_or(0),
        contour_w: attr_i64(&n, "contourW").unwrap_or(0),
        contour_clr,
        prst_material: attr(&n, "prstMaterial").unwrap_or_else(|| "warmMatte".into()),
        bevel_t: child(n, "bevelT").map(parse_bevel3d),
        bevel_b: child(n, "bevelB").map(parse_bevel3d),
    })
}

// ===========================
//  Custom geometry parsing
// ===========================

const DRAWINGML_FULL_CIRCLE: f64 = 21_600_000.0;
const DRAWINGML_ANGLE_TO_RAD: f64 = std::f64::consts::TAU / DRAWINGML_FULL_CIRCLE;

/// Evaluate the ordered guide program carried by `a:custGeom`.
///
/// ECMA-376 Part 1 §20.1.9.11 defines 17 prefix operators and requires guides
/// to be evaluated in document order. Coordinates in `pathLst` may refer to
/// these names instead of containing numeric literals. Keeping the evaluation
/// in the parser preserves the compact, normalized `PathCmd` wire model and
/// avoids repeating formula work on every canvas repaint.
fn custom_geometry_guides(
    cust_geom: roxmltree::Node<'_, '_>,
    shape_w: f64,
    shape_h: f64,
) -> HashMap<String, f64> {
    let first_path = child(cust_geom, "pathLst").and_then(|paths| {
        paths
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "path")
    });
    let w = if shape_w > 0.0 {
        shape_w
    } else {
        first_path
            .and_then(|path| attr_f64(&path, "w"))
            .unwrap_or(1.0)
            .max(1.0)
    };
    let h = if shape_h > 0.0 {
        shape_h
    } else {
        first_path
            .and_then(|path| attr_f64(&path, "h"))
            .unwrap_or(1.0)
            .max(1.0)
    };
    let ss = w.min(h);
    let ls = w.max(h);
    let mut env = HashMap::new();

    macro_rules! builtins {
        ($($name:expr => $value:expr),* $(,)?) => {
            $(env.insert($name.to_owned(), $value);)*
        };
    }
    builtins! {
        "w" => w, "h" => h,
        "l" => 0.0, "t" => 0.0, "r" => w, "b" => h,
        "hc" => w / 2.0, "vc" => h / 2.0,
        "wd2" => w / 2.0, "wd3" => w / 3.0, "wd4" => w / 4.0,
        "wd5" => w / 5.0, "wd6" => w / 6.0, "wd8" => w / 8.0,
        "wd10" => w / 10.0, "wd12" => w / 12.0, "wd16" => w / 16.0,
        "wd32" => w / 32.0,
        "hd2" => h / 2.0, "hd3" => h / 3.0, "hd4" => h / 4.0,
        "hd5" => h / 5.0, "hd6" => h / 6.0, "hd8" => h / 8.0,
        "hd10" => h / 10.0, "hd12" => h / 12.0, "hd16" => h / 16.0,
        "hd32" => h / 32.0,
        "ss" => ss, "ssd2" => ss / 2.0, "ssd4" => ss / 4.0,
        "ssd6" => ss / 6.0, "ssd8" => ss / 8.0, "ssd16" => ss / 16.0,
        "ssd32" => ss / 32.0,
        "ls" => ls, "lsd2" => ls / 2.0, "lsd4" => ls / 4.0,
        "lsd6" => ls / 6.0, "lsd8" => ls / 8.0, "lsd16" => ls / 16.0,
        "lsd32" => ls / 32.0,
        "cd" => DRAWINGML_FULL_CIRCLE,
        "cd2" => DRAWINGML_FULL_CIRCLE / 2.0,
        "cd4" => DRAWINGML_FULL_CIRCLE / 4.0,
        "cd8" => DRAWINGML_FULL_CIRCLE / 8.0,
        "3cd4" => 3.0 * DRAWINGML_FULL_CIRCLE / 4.0,
        "3cd8" => 3.0 * DRAWINGML_FULL_CIRCLE / 8.0,
        "5cd8" => 5.0 * DRAWINGML_FULL_CIRCLE / 8.0,
        "7cd8" => 7.0 * DRAWINGML_FULL_CIRCLE / 8.0,
    }

    for list_name in ["avLst", "gdLst"] {
        let Some(list) = child(cust_geom, list_name) else {
            continue;
        };
        for guide in list
            .children()
            .filter(|node| node.is_element() && node.tag_name().name() == "gd")
        {
            let (Some(name), Some(formula)) = (attr(&guide, "name"), attr(&guide, "fmla")) else {
                continue;
            };
            if let Some(value) = evaluate_geometry_formula(&formula, &env) {
                env.insert(name, value);
            }
        }
    }
    env
}

fn resolve_geometry_value(token: &str, env: &HashMap<String, f64>) -> Option<f64> {
    env.get(token)
        .copied()
        .or_else(|| token.parse::<f64>().ok())
        .filter(|value| value.is_finite())
}

fn evaluate_geometry_formula(formula: &str, env: &HashMap<String, f64>) -> Option<f64> {
    let mut tokens = formula.split_whitespace();
    let op = tokens.next()?;
    let args: Vec<f64> = tokens
        .map(|token| resolve_geometry_value(token, env))
        .collect::<Option<_>>()?;
    let arg = |index: usize| args.get(index).copied();
    let value = match op {
        "val" => arg(0)?,
        "*/" => arg(0)? * arg(1)? / arg(2)?,
        "+-" => arg(0)? + arg(1)? - arg(2)?,
        "+/" => (arg(0)? + arg(1)?) / arg(2)?,
        "?:" => {
            if arg(0)? > 0.0 {
                arg(1)?
            } else {
                arg(2)?
            }
        }
        "abs" => arg(0)?.abs(),
        "at2" => arg(1)?.atan2(arg(0)?) / DRAWINGML_ANGLE_TO_RAD,
        "cat2" => arg(0)? * arg(2)?.atan2(arg(1)?).cos(),
        "cos" => arg(0)? * (arg(1)? * DRAWINGML_ANGLE_TO_RAD).cos(),
        "max" => arg(0)?.max(arg(1)?),
        "min" => arg(0)?.min(arg(1)?),
        "mod" => (arg(0)?.powi(2) + arg(1)?.powi(2) + arg(2)?.powi(2)).sqrt(),
        "pin" => {
            let (low, value, high) = (arg(0)?, arg(1)?, arg(2)?);
            if value < low {
                low
            } else if value > high {
                high
            } else {
                value
            }
        }
        "sat2" => arg(0)? * arg(2)?.atan2(arg(1)?).sin(),
        "sin" => arg(0)? * (arg(1)? * DRAWINGML_ANGLE_TO_RAD).sin(),
        "sqrt" => arg(0)?.max(0.0).sqrt(),
        "tan" => arg(0)? * (arg(1)? * DRAWINGML_ANGLE_TO_RAD).tan(),
        _ => return None,
    };
    value.is_finite().then_some(value)
}

fn geometry_attr(
    node: &roxmltree::Node<'_, '_>,
    name: &str,
    env: &HashMap<String, f64>,
) -> Option<f64> {
    let token = attr(node, name)?;
    resolve_geometry_value(&token, env)
}

/// Parse a single path command node; coordinates are normalised to [0,1].
pub(crate) fn parse_path_cmd(
    cmd_node: roxmltree::Node<'_, '_>,
    path_w: f64,
    path_h: f64,
    env: &HashMap<String, f64>,
) -> Option<PathCmd> {
    match cmd_node.tag_name().name() {
        "moveTo" => {
            let pt = child(cmd_node, "pt")?;
            let x = geometry_attr(&pt, "x", env)? / path_w;
            let y = geometry_attr(&pt, "y", env)? / path_h;
            Some(PathCmd::MoveTo { x, y })
        }
        "lnTo" => {
            let pt = child(cmd_node, "pt")?;
            let x = geometry_attr(&pt, "x", env)? / path_w;
            let y = geometry_attr(&pt, "y", env)? / path_h;
            Some(PathCmd::LineTo { x, y })
        }
        "cubicBezTo" => {
            let pts: Vec<_> = children_vec(cmd_node, "pt");
            if pts.len() < 3 {
                return None;
            }
            let x1 = geometry_attr(&pts[0], "x", env)? / path_w;
            let y1 = geometry_attr(&pts[0], "y", env)? / path_h;
            let x2 = geometry_attr(&pts[1], "x", env)? / path_w;
            let y2 = geometry_attr(&pts[1], "y", env)? / path_h;
            let x = geometry_attr(&pts[2], "x", env)? / path_w;
            let y = geometry_attr(&pts[2], "y", env)? / path_h;
            Some(PathCmd::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x,
                y,
            })
        }
        "arcTo" => {
            // wR/hR are radii in path-local units; stAng/swAng in 60000ths of a degree
            let wr = geometry_attr(&cmd_node, "wR", env).unwrap_or(0.0) / path_w;
            let hr = geometry_attr(&cmd_node, "hR", env).unwrap_or(0.0) / path_h;
            let st_ang = geometry_attr(&cmd_node, "stAng", env).unwrap_or(0.0) / 60000.0;
            let sw_ang = geometry_attr(&cmd_node, "swAng", env).unwrap_or(0.0) / 60000.0;
            Some(PathCmd::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            })
        }
        "close" => Some(PathCmd::Close),
        _ => None,
    }
}

/// Parse custGeom > pathLst into a list of sub-paths (one per <a:path> element).
pub(crate) fn parse_cust_geom(
    cust_geom: roxmltree::Node<'_, '_>,
    shape_w: f64,
    shape_h: f64,
) -> Vec<Vec<PathCmd>> {
    let path_lst = match child(cust_geom, "pathLst") {
        Some(n) => n,
        None => return vec![],
    };

    let env = custom_geometry_guides(cust_geom, shape_w, shape_h);
    path_lst
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "path")
        .map(|path_node| {
            // CT_Path2D defaults w/h to zero. A zero/omitted coordinate-system
            // size means the path's guide values are already in shape space;
            // normalize them by the shape extents. Treating the schema default
            // as one makes otherwise valid guide-based paths enormous.
            let path_w = attr_f64(&path_node, "w")
                .filter(|value| *value > 0.0)
                .unwrap_or(shape_w.max(1.0));
            let path_h = attr_f64(&path_node, "h")
                .filter(|value| *value > 0.0)
                .unwrap_or(shape_h.max(1.0));
            path_node
                .children()
                .filter(|n| n.is_element())
                .filter_map(|cmd| parse_path_cmd(cmd, path_w, path_h, &env))
                .collect()
        })
        .collect()
}

// ===========================
//  Transform (a:xfrm)
// ===========================

pub(crate) fn parse_xfrm(xfrm: roxmltree::Node<'_, '_>) -> Transform {
    let rot = attr_f64(&xfrm, "rot").unwrap_or(0.0) / 60000.0;
    let flip_h = attr(&xfrm, "flipH")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);
    let flip_v = attr(&xfrm, "flipV")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);
    let off = child(xfrm, "off");
    let ext = child(xfrm, "ext");
    Transform {
        x: off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
        y: off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
        cx: ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
        cy: ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
        rot,
        flip_h,
        flip_v,
    }
}

// ===========================
//  Slide background
// ===========================

/// ECMA-376 §19.3.1.1 `p:bg`. `resolve_blip` maps a `<a:blip r:embed>` rId to a
/// base64 data URL using the rels + zip of the part this `c_sld` belongs to
/// (slide / layout / master), so an image background (§20.1.8.14) is resolved
/// against the correct relationship base.
pub(crate) fn parse_background<F: FnMut(&str) -> Option<String>>(
    c_sld: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> Option<Fill> {
    let bg = child(c_sld, "bg")?;
    // bgPr contains an explicit fill specification
    if let Some(bg_pr) = child(bg, "bgPr") {
        // §20.1.8.14 — an image background lives in `bgPr > blipFill`. Try it
        // first so the embedded blip is resolved; fall back to the generic
        // solid/gradient/pattern parser for non-image bgPr fills.
        if let Some(blip_fill) = child(bg_pr, "blipFill") {
            if let Some(fill) = parse_blip_fill(blip_fill, theme, resolve_blip) {
                return Some(fill);
            }
        }
        return parse_fill_tint(bg_pr, theme, ooxml_common::color::TintMode::WordLiteral);
    }
    // bgRef references a theme background style; its child is a color element
    if let Some(bg_ref) = child(bg, "bgRef") {
        return parse_style_matrix_fill(bg_ref, theme, true)
            .or_else(|| parse_color_node(bg_ref, theme).map(|c| Fill::Solid { color: c }));
    }
    None
}

/// Resolve a table-style `<a:fill>` wrapper's colour. Identical to `parse_fill`
/// for the common solid/no-fill cases, except `<a:tint>` uses the literal
/// ECMA-376 §20.1.2.3.34 formula (`TintMode::WordLiteral`) so a band's
/// `accent + tint 20%` renders as the near-white wash PowerPoint draws, rather
/// than the saturated linear-lerp used for SmartArt accents. Gradient/pattern/
/// blip fills (rare in table styles) defer to the generic `parse_fill`.
pub(crate) fn parse_table_style_fill(
    fill_wrapper: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Fill> {
    use ooxml_common::color::TintMode::WordLiteral;
    for c in fill_wrapper.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "noFill" => return Some(Fill::None),
            "solidFill" => {
                return parse_color_node_tint(c, theme, WordLiteral)
                    .map(|color| Fill::Solid { color });
            }
            _ => {}
        }
    }
    parse_fill(fill_wrapper, theme)
}

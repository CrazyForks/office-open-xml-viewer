//! OOXML color transforms (lumMod, lumOff, satMod, satOff, hueMod, hueOff,
//! shade, tint, alpha and friends) shared between the docx and pptx parsers.
//!
//! Word and PowerPoint diverge on the `tint` transform — Word reads `val` as
//! the *retained fraction of the input color* (the literal ECMA-376
//! §20.1.2.3.34 reading: `result = val·input + (1-val)·white`), while
//! PowerPoint applies it as a `lerp(input, white, val)` in linear sRGB. Empirical
//! comparison against PDF exports confirms each app does its own thing — see
//! `TintMode` and the per-app `apply_color_transforms_with` flag.
//!
//! Everything else (shade, lumMod/Off, satMod/Off, hueMod/Off, alpha
//! family) is identical between the two and lives here uncopied.

use roxmltree::Node;

/// Default DrawingML logical-color → theme scheme-slot mapping
/// (ECMA-376 §19.3.1.6 `clrMap` / CT_ColorMapping, with the standard
/// PowerPoint default attribute values).
///
/// A theme's `<a:clrScheme>` stores twelve named slots — `dk1`, `lt1`, `dk2`,
/// `lt2`, `accent1`..`accent6`, `hlink`, `folHlink`. Documents reference colors
/// by *logical* name (`bg1`, `tx1`, `bg2`, `tx2`, the accents, and the two
/// hyperlink names) and a `clrMap` indirection layer resolves each logical name
/// to a slot. When no explicit `clrMap` override is present the default mapping
/// applies:
///
/// - `bg1` → `lt1`, `tx1` → `dk1` (background 1 is the light slot, text 1 the
///   dark slot)
/// - `bg2` → `lt2`, `tx2` → `dk2`
/// - `accent1`..`accent6`, `hlink`, `folHlink` map to the identically named slot
///
/// This is the single source of truth for that default table. Each parser keeps
/// its own *resolution* (storage layout and per-app tint), but the logical→slot
/// names live here. See also [`default_scheme_slot`] for a lookup that also
/// accepts a raw slot name and returns it unchanged.
pub const SCHEME_DEFAULT_SLOTS: &[(&str, &str)] = &[
    ("bg1", "lt1"),
    ("tx1", "dk1"),
    ("bg2", "lt2"),
    ("tx2", "dk2"),
    ("accent1", "accent1"),
    ("accent2", "accent2"),
    ("accent3", "accent3"),
    ("accent4", "accent4"),
    ("accent5", "accent5"),
    ("accent6", "accent6"),
    ("hlink", "hlink"),
    ("folHlink", "folHlink"),
];

/// Resolve a DrawingML color name to its default theme scheme slot
/// (ECMA-376 §19.3.1.6, default `clrMap`).
///
/// Logical names listed in [`SCHEME_DEFAULT_SLOTS`] (`bg1`/`tx1`/`bg2`/`tx2`)
/// map to their slot; every other input — including the raw slot names
/// (`dk1`, `lt1`, …), the accents and the hyperlink names — is returned
/// unchanged. This mirrors how a parser walks a `schemeClr@val` that may carry
/// either a logical name *or* a slot name and wants the underlying slot.
pub fn default_scheme_slot(name: &str) -> &str {
    SCHEME_DEFAULT_SLOTS
        .iter()
        .find(|(logical, _)| *logical == name)
        .map(|(_, slot)| *slot)
        .unwrap_or(name)
}

/// Selects the formula applied to `<a:tint val>` modifiers. The OOXML spec
/// is consistent (val = retained input), but the two desktop apps render
/// templates differently in practice — see ECMA-376 §20.1.2.3.34 and the
/// commit history of pptx-parser for the linear-sRGB derivation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TintMode {
    /// Word: `result = val·input + (1-val)·white` in sRGB. Matches Word's
    /// rendering of resume / cover templates that use accent recolors with
    /// tint values.
    WordLiteral,
    /// PowerPoint: `lerp(input, white, val)` in linear sRGB. Matches
    /// PowerPoint's rendering of SmartArt accent recolors pixel-for-pixel.
    PowerPointLinear,
}

/// Apply OOXML color transforms to `hex` based on the modifier elements
/// declared as direct children of `node`. Returns 6-char hex when fully
/// opaque, or 8-char hex (RRGGBBAA) when alpha < 1.
pub fn apply_color_transforms(hex: &str, node: Node, tint_mode: TintMode) -> String {
    if hex.len() < 6 {
        return hex.to_owned();
    }
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0);
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0);
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);

    let mut rf = r as f64 / 255.0;
    let mut gf = g as f64 / 255.0;
    let mut bf = b as f64 / 255.0;
    let mut alpha = if hex.len() >= 8 {
        u8::from_str_radix(&hex[6..8], 16).unwrap_or(255) as f64 / 255.0
    } else {
        1.0
    };

    let attr_pct = |t: &Node, name: &str, default: f64| -> f64 {
        t.attribute(name)
            .and_then(|v| v.parse::<f64>().ok())
            .unwrap_or(default)
            / 100_000.0
    };

    for t in node.children().filter(|n| n.is_element()) {
        match t.tag_name().name() {
            "lumMod" => {
                let val = attr_pct(&t, "val", 100_000.0);
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb(h, (l * val).min(1.0), s);
                rf = nr;
                gf = ng;
                bf = nb;
            }
            "lumOff" => {
                let val = attr_pct(&t, "val", 0.0);
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb(h, (l + val).clamp(0.0, 1.0), s);
                rf = nr;
                gf = ng;
                bf = nb;
            }
            "satMod" => {
                let val = attr_pct(&t, "val", 100_000.0);
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb(h, l, (s * val).clamp(0.0, 1.0));
                rf = nr;
                gf = ng;
                bf = nb;
            }
            "satOff" => {
                let val = attr_pct(&t, "val", 0.0);
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb(h, l, (s + val).clamp(0.0, 1.0));
                rf = nr;
                gf = ng;
                bf = nb;
            }
            "hueMod" => {
                let val = attr_pct(&t, "val", 100_000.0);
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb((h * val).rem_euclid(1.0), l, s);
                rf = nr;
                gf = ng;
                bf = nb;
            }
            "hueOff" => {
                // hueOff is in 60000ths of a degree per ECMA-376 §20.1.2.3.16.
                let val_deg = t
                    .attribute("val")
                    .and_then(|v| v.parse::<f64>().ok())
                    .unwrap_or(0.0)
                    / 60_000.0;
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb((h + val_deg / 360.0).rem_euclid(1.0), l, s);
                rf = nr;
                gf = ng;
                bf = nb;
            }
            "shade" => {
                // ECMA-376 §20.1.2.3.31: result = val·input + (1-val)·black.
                let val = attr_pct(&t, "val", 100_000.0);
                rf *= val;
                gf *= val;
                bf *= val;
            }
            "tint" => {
                let val = attr_pct(&t, "val", 0.0);
                match tint_mode {
                    TintMode::WordLiteral => {
                        // `result = val·input + (1-val)·white` per literal spec.
                        rf = val * rf + (1.0 - val);
                        gf = val * gf + (1.0 - val);
                        bf = val * bf + (1.0 - val);
                    }
                    TintMode::PowerPointLinear => {
                        // PowerPoint reads val as the lerp fraction toward
                        // white in LINEAR sRGB. Verified against PDF
                        // exports of SmartArt accent recolors.
                        let lr = srgb_to_linear(rf);
                        let lg = srgb_to_linear(gf);
                        let lb = srgb_to_linear(bf);
                        rf = linear_to_srgb((lr + (1.0 - lr) * val).clamp(0.0, 1.0));
                        gf = linear_to_srgb((lg + (1.0 - lg) * val).clamp(0.0, 1.0));
                        bf = linear_to_srgb((lb + (1.0 - lb) * val).clamp(0.0, 1.0));
                    }
                }
            }
            "alpha" => {
                // ECMA-376 §20.1.2.3.1 — sets absolute alpha.
                alpha = attr_pct(&t, "val", 100_000.0);
            }
            "alphaModFix" => {
                // ECMA-376 §20.1.8.4 — fixed (absolute) alpha modulation.
                alpha = attr_pct(&t, "amt", 100_000.0);
            }
            "alphaMod" => {
                // ECMA-376 §20.1.2.3.2 — multiply current alpha by val/100000.
                alpha *= attr_pct(&t, "val", 100_000.0);
            }
            "alphaOff" => {
                // ECMA-376 §20.1.2.3.3 — additive offset to alpha.
                alpha += attr_pct(&t, "val", 0.0);
            }
            _ => {}
        }
    }

    let r = (rf.clamp(0.0, 1.0) * 255.0).round() as u8;
    let g = (gf.clamp(0.0, 1.0) * 255.0).round() as u8;
    let b = (bf.clamp(0.0, 1.0) * 255.0).round() as u8;
    if (alpha - 1.0).abs() < 0.004 {
        format!("{:02X}{:02X}{:02X}", r, g, b)
    } else {
        let a = (alpha.clamp(0.0, 1.0) * 255.0).round() as u8;
        format!("{:02X}{:02X}{:02X}{:02X}", r, g, b, a)
    }
}

/// A DrawingML color element (`<a:srgbClr>` / `<a:sysClr>` / `<a:prstClr>` /
/// `<a:schemeClr>`), extracted from a color container's first color child but
/// **not yet resolved to a hex string**. This is the format-agnostic middle
/// layer between [`extract_color_source`] (which finds the element) and
/// [`parse_color_node`] (which resolves + transforms it). The three parsers
/// share the *grammar* here; theme-slot resolution (which differs per host —
/// pptx bakes a `clrMap`, xlsx indexes a positional palette, docx keys a slot
/// map) stays behind the [`ThemeResolver`] trait.
///
/// The wrapped `Node` is retained so the caller can apply the color-transform
/// children (lumMod/tint/…) via [`apply_color_transforms`] on the resolved base.
#[derive(Debug, Clone)]
pub enum ColorSource<'a, 'input> {
    /// `<a:srgbClr val="RRGGBB">` — an explicit hex (ECMA-376 §20.1.2.3.32).
    /// `val` is the raw authored string (case as-authored).
    SrgbClr { val: String, node: Node<'a, 'input> },
    /// `<a:sysClr val="windowText" lastClr="RRGGBB">` — a system color
    /// (§20.1.2.3.33). `last_clr` is the cached resolved hex; `val` is the
    /// system-color enum name, retained only as a last-ditch fallback (matches
    /// docx's historical behavior — pptx/xlsx never authored a val-only sysClr).
    SysClr {
        last_clr: Option<String>,
        val: Option<String>,
        node: Node<'a, 'input>,
    },
    /// `<a:prstClr val="black">` — a named preset color (§20.1.2.3.22 /
    /// §20.1.10.48). Resolved through [`preset_color`](crate::theme::preset_color).
    PrstClr { val: String },
    /// `<a:schemeClr val="accent1">` — a theme-slot reference (§20.1.2.3.29).
    /// `val` is the logical/slot name; the [`ThemeResolver`] maps it to a base
    /// hex. The node is retained for the transform children.
    SchemeClr { val: String, node: Node<'a, 'input> },
}

/// Locate the first DrawingML color element among a container's element
/// children and return it as a [`ColorSource`]. Covers the four color-choice
/// members a `<a:solidFill>` / `<a:gs>` / run `<a:rPr>` / shadow etc. carry:
/// `srgbClr`, `sysClr`, `prstClr`, `schemeClr`. Returns `None` when no such
/// child exists (e.g. a `<a:noFill>` sibling, or a `<a:solidFill>` with no
/// color child). The caller supplies the *container* node (the same node its
/// prior inline `for child` loop walked).
pub fn extract_color_source<'a, 'input>(
    container: Node<'a, 'input>,
) -> Option<ColorSource<'a, 'input>> {
    for c in container.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "srgbClr" => {
                return Some(ColorSource::SrgbClr {
                    val: c.attribute("val")?.to_owned(),
                    node: c,
                });
            }
            "sysClr" => {
                return Some(ColorSource::SysClr {
                    last_clr: c.attribute("lastClr").map(str::to_owned),
                    val: c.attribute("val").map(str::to_owned),
                    node: c,
                });
            }
            "prstClr" => {
                return Some(ColorSource::PrstClr {
                    val: c.attribute("val")?.to_owned(),
                });
            }
            "schemeClr" => {
                return Some(ColorSource::SchemeClr {
                    val: c.attribute("val")?.to_owned(),
                    node: c,
                });
            }
            _ => {}
        }
    }
    None
}

/// Host-specific resolution of a `<a:schemeClr val>` name to its **base hex**
/// (6-char, no `#`, ready for [`apply_color_transforms`]). Each parser keeps
/// its own theme storage and logical-name handling:
///
/// - pptx bakes the master's `<p:clrMap>` into its theme map and falls back to
///   the default §19.3.1.6 slot table (plus a `phClr`→dk1 approximation);
/// - xlsx indexes a positional `dk1,lt1,dk2,lt2,accent1..6,hlink,folHlink`
///   palette (with the well-known dk/lt index swap);
/// - docx keys a slot map and substitutes a caller-provided name for `phClr`.
///
/// Returning `None` leaves the color unresolved (the caller then falls back —
/// e.g. to a shape's style color). Implementations must return the base
/// **without** a leading `#` so the shared transform sees clean input.
pub trait ThemeResolver {
    /// Resolve a scheme/logical color name to its base hex (no `#`). `name` is
    /// the raw `<a:schemeClr val>` string.
    fn resolve_scheme_color(&self, name: &str) -> Option<String>;
}

/// Resolve a located color container to a hex string, sharing the DrawingML
/// color grammar across the docx/pptx/xlsx parsers. Finds the first color child
/// ([`extract_color_source`]), resolves it, and applies any color-transform
/// children (lumMod/lumOff/shade/tint/… via [`apply_color_transforms`]):
///
/// - `srgbClr` / `sysClr` → their hex + transforms;
/// - `prstClr` → [`preset_color`](crate::theme::preset_color) (no transforms,
///   matching the parsers' prior behavior — a transformed preset is a latent
///   gap shared by all three);
/// - `schemeClr` → `resolver.resolve_scheme_color(val)` + transforms.
///
/// The returned hex is uppercase and has **no** `#` prefix (that is what
/// [`apply_color_transforms`] emits: `{:02X}`); each caller re-applies its own
/// casing / `#` convention in a thin adapter. `tint_mode` selects the Word vs
/// PowerPoint `<a:tint>` interpretation ([`TintMode`]). Returns `None` when no
/// color child is present or a `schemeClr` fails to resolve.
pub fn parse_color_node<R: ThemeResolver + ?Sized>(
    container: Node<'_, '_>,
    resolver: &R,
    tint_mode: TintMode,
) -> Option<String> {
    match extract_color_source(container)? {
        ColorSource::SrgbClr { val, node } => Some(apply_color_transforms(&val, node, tint_mode)),
        ColorSource::SysClr {
            last_clr,
            val,
            node,
        } => {
            // Prefer the cached lastClr; fall back to the enum name only when
            // lastClr is absent (docx's historical behavior — see ColorSource).
            let hex = last_clr.or(val)?;
            Some(apply_color_transforms(&hex, node, tint_mode))
        }
        ColorSource::PrstClr { val } => crate::theme::preset_color(&val),
        ColorSource::SchemeClr { val, node } => {
            let base = resolver.resolve_scheme_color(&val)?;
            Some(apply_color_transforms(&base, node, tint_mode))
        }
    }
}

/// sRGB → linear light. IEC 61966-2-1 transfer function.
pub fn srgb_to_linear(c: f64) -> f64 {
    if c <= 0.04045 {
        c / 12.92
    } else {
        ((c + 0.055) / 1.055).powf(2.4)
    }
}

/// Linear light → sRGB.
pub fn linear_to_srgb(c: f64) -> f64 {
    if c <= 0.0031308 {
        12.92 * c
    } else {
        1.055 * c.powf(1.0 / 2.4) - 0.055
    }
}

/// RGB → HLS conversion (lightness in the middle, matches Python's colorsys).
/// Returns (h, l, s) with each component in [0, 1].
pub fn rgb_to_hls(r: f64, g: f64, b: f64) -> (f64, f64, f64) {
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    let d = max - min;
    if d < 1e-10 {
        return (0.0, l, 0.0);
    }
    let s = if l > 0.5 {
        d / (2.0 - max - min)
    } else {
        d / (max + min)
    };
    let h = if (max - r).abs() < 1e-10 {
        (g - b) / d + if g < b { 6.0 } else { 0.0 }
    } else if (max - g).abs() < 1e-10 {
        (b - r) / d + 2.0
    } else {
        (r - g) / d + 4.0
    };
    (h / 6.0, l, s)
}

/// HLS → RGB conversion. (h, l, s) are each in [0, 1]; returns linear-RGB
/// triple in [0, 1].
pub fn hls_to_rgb(h: f64, l: f64, s: f64) -> (f64, f64, f64) {
    if s < 1e-10 {
        return (l, l, l);
    }
    fn hue2rgb(p: f64, q: f64, mut t: f64) -> f64 {
        if t < 0.0 {
            t += 1.0;
        }
        if t > 1.0 {
            t -= 1.0;
        }
        if t < 1.0 / 6.0 {
            return p + (q - p) * 6.0 * t;
        }
        if t < 0.5 {
            return q;
        }
        if t < 2.0 / 3.0 {
            return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
        }
        p
    }
    let q = if l < 0.5 {
        l * (1.0 + s)
    } else {
        l + s - l * s
    };
    let p = 2.0 * l - q;
    (
        hue2rgb(p, q, h + 1.0 / 3.0),
        hue2rgb(p, q, h),
        hue2rgb(p, q, h - 1.0 / 3.0),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Pins the default §19.3.1.6 logical→slot table so the three parsers that
    /// consume it can't drift. The twelve pairs are exactly the PowerPoint
    /// default `clrMap`.
    #[test]
    fn scheme_default_slots_match_spec() {
        assert_eq!(
            SCHEME_DEFAULT_SLOTS,
            &[
                ("bg1", "lt1"),
                ("tx1", "dk1"),
                ("bg2", "lt2"),
                ("tx2", "dk2"),
                ("accent1", "accent1"),
                ("accent2", "accent2"),
                ("accent3", "accent3"),
                ("accent4", "accent4"),
                ("accent5", "accent5"),
                ("accent6", "accent6"),
                ("hlink", "hlink"),
                ("folHlink", "folHlink"),
            ]
        );
    }

    // ── parse_color_node / extract_color_source ─────────────────────────────

    use roxmltree::Document;

    /// A minimal ThemeResolver mapping a couple of slot names to base hex,
    /// letting the scheme-color path be exercised without a full parser.
    struct MapResolver;
    impl ThemeResolver for MapResolver {
        fn resolve_scheme_color(&self, name: &str) -> Option<String> {
            match name {
                "accent1" => Some("4472C4".to_owned()),
                "dk1" => Some("000000".to_owned()),
                _ => None,
            }
        }
    }

    fn parse(xml: &str, mode: TintMode) -> Option<String> {
        // Every fixture wraps the color element in a `<a:solidFill>` container,
        // matching how the parsers pass the located fill node.
        let doc = Document::parse(xml).unwrap();
        parse_color_node(doc.root_element(), &MapResolver, mode)
    }

    const NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// srgbClr resolves to its (uppercased) hex, no `#`, under either TintMode.
    #[test]
    fn parse_color_node_srgb_plain() {
        let xml = format!(r#"<a:solidFill xmlns:a="{NS}"><a:srgbClr val="ff8000"/></a:solidFill>"#);
        assert_eq!(
            parse(&xml, TintMode::WordLiteral).as_deref(),
            Some("FF8000")
        );
        assert_eq!(
            parse(&xml, TintMode::PowerPointLinear).as_deref(),
            Some("FF8000")
        );
    }

    /// sysClr uses lastClr; a transform child (lumMod) is applied on top.
    #[test]
    fn parse_color_node_sysclr_lastclr_with_lummod() {
        let xml = format!(
            r#"<a:solidFill xmlns:a="{NS}"><a:sysClr val="windowText" lastClr="FFFFFF"><a:lumMod val="50000"/></a:sysClr></a:solidFill>"#
        );
        // 50% luminance of white → mid gray (uppercase, no #).
        assert_eq!(
            parse(&xml, TintMode::WordLiteral).as_deref(),
            Some("808080")
        );
    }

    /// sysClr with no lastClr falls back to `val` (docx-historical). "windowText"
    /// is not valid hex, but the fallback path must still fire (produces 000000).
    #[test]
    fn parse_color_node_sysclr_val_fallback() {
        let xml = format!(r#"<a:solidFill xmlns:a="{NS}"><a:sysClr val="000000"/></a:solidFill>"#);
        assert_eq!(
            parse(&xml, TintMode::WordLiteral).as_deref(),
            Some("000000")
        );
    }

    /// prstClr resolves via the shared preset table, WITHOUT applying transforms
    /// (matches the parsers' prior behavior — documented latent gap).
    #[test]
    fn parse_color_node_prstclr_via_preset_table() {
        let xml = format!(r#"<a:solidFill xmlns:a="{NS}"><a:prstClr val="orange"/></a:solidFill>"#);
        assert_eq!(
            parse(&xml, TintMode::WordLiteral).as_deref(),
            Some("FFA500")
        );
        // A transform child is intentionally ignored for prstClr.
        let xml2 = format!(
            r#"<a:solidFill xmlns:a="{NS}"><a:prstClr val="black"><a:lumMod val="50000"/></a:prstClr></a:solidFill>"#
        );
        assert_eq!(
            parse(&xml2, TintMode::WordLiteral).as_deref(),
            Some("000000")
        );
    }

    /// schemeClr resolves through the injected ThemeResolver, then transforms.
    #[test]
    fn parse_color_node_schemeclr_via_resolver() {
        let xml =
            format!(r#"<a:solidFill xmlns:a="{NS}"><a:schemeClr val="accent1"/></a:solidFill>"#);
        assert_eq!(
            parse(&xml, TintMode::WordLiteral).as_deref(),
            Some("4472C4")
        );
        // Unknown slot → None (caller falls back).
        let unknown =
            format!(r#"<a:solidFill xmlns:a="{NS}"><a:schemeClr val="accent9"/></a:solidFill>"#);
        assert_eq!(parse(&unknown, TintMode::WordLiteral), None);
    }

    /// The two TintModes diverge on `<a:tint>`: Word reads val as retained input
    /// (a near-white wash at 20%), PowerPoint lerps toward white in linear sRGB.
    /// The two must produce DIFFERENT hex for the same input, proving the mode
    /// is threaded through parse_color_node.
    #[test]
    fn parse_color_node_tint_mode_diverges() {
        let xml = format!(
            r#"<a:solidFill xmlns:a="{NS}"><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill>"#
        );
        let word = parse(&xml, TintMode::WordLiteral).unwrap();
        let ppt = parse(&xml, TintMode::PowerPointLinear).unwrap();
        assert_ne!(word, ppt);
        // Both are 6-char uppercase hex with no '#'.
        assert_eq!(word.len(), 6);
        assert!(!word.contains('#'));
        assert!(word.chars().all(|ch| ch.is_ascii_hexdigit()));
    }

    /// alpha < 1 yields an 8-char RRGGBBAA hex (transform emits the alpha byte).
    #[test]
    fn parse_color_node_alpha_yields_eight_hex() {
        let xml = format!(
            r#"<a:solidFill xmlns:a="{NS}"><a:srgbClr val="112233"><a:alpha val="50000"/></a:srgbClr></a:solidFill>"#
        );
        let out = parse(&xml, TintMode::WordLiteral).unwrap();
        assert_eq!(out.len(), 8);
        assert!(out.starts_with("112233"));
    }

    /// A container with no color child (e.g. a lone noFill sibling) → None.
    #[test]
    fn parse_color_node_none_when_no_color_child() {
        let xml = format!(r#"<a:solidFill xmlns:a="{NS}"><a:noFill/></a:solidFill>"#);
        assert_eq!(parse(&xml, TintMode::WordLiteral), None);
    }

    /// extract_color_source returns the first color element and skips others.
    #[test]
    fn extract_color_source_picks_first_color_child() {
        let xml = format!(
            r#"<a:solidFill xmlns:a="{NS}"><a:srgbClr val="ABCDEF"/><a:schemeClr val="accent1"/></a:solidFill>"#
        );
        let doc = Document::parse(&xml).unwrap();
        match extract_color_source(doc.root_element()) {
            Some(ColorSource::SrgbClr { val, .. }) => assert_eq!(val, "ABCDEF"),
            other => panic!("expected SrgbClr first, got {other:?}"),
        }
    }

    #[test]
    fn default_scheme_slot_maps_logicals_and_passes_through_slots() {
        // Logical names resolve to their slot.
        assert_eq!(default_scheme_slot("bg1"), "lt1");
        assert_eq!(default_scheme_slot("tx1"), "dk1");
        assert_eq!(default_scheme_slot("bg2"), "lt2");
        assert_eq!(default_scheme_slot("tx2"), "dk2");
        // Raw slot names and accents/hyperlinks pass through unchanged.
        assert_eq!(default_scheme_slot("dk1"), "dk1");
        assert_eq!(default_scheme_slot("lt1"), "lt1");
        assert_eq!(default_scheme_slot("accent3"), "accent3");
        assert_eq!(default_scheme_slot("hlink"), "hlink");
        // Unknown input is returned verbatim (caller decides what to do).
        assert_eq!(default_scheme_slot("phClr"), "phClr");
    }
}

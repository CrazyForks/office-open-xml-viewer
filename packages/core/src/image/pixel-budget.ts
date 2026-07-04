// ── Shared raster pixel-dimension budget (DoS / decode-bomb guard) ───────────
//
// One source of truth for the caps that bound how large a decoded raster surface
// may be, used both by the metafile-embedded DIB decoder (`./dib.ts`) and by the
// pre-`createImageBitmap` header sniff (`./raster-dimensions.ts`). Keeping the
// two paths on the SAME numbers means a raster that is refused as a standalone
// blip is also refused when embedded in a WMF/EMF, and vice versa.

/**
 * Maximum width or height (px) accepted for a decoded raster. 32767 is the
 * largest dimension every major browser accepts for a `<canvas>` /
 * `OffscreenCanvas` (Chrome, Firefox and Safari all top out at 32767 on at least
 * one axis). A raster wider or taller than this could never be drawn to a canvas
 * anyway, so it is rejected before decode.
 */
export const MAX_RASTER_DIMENSION = 32767;

/**
 * Megapixel budget for a decoded raster: 64 MP (2^26 px). A decoded surface is
 * `width × height × 4` bytes of RGBA, so 64 MP bounds it to 256 MiB — the
 * practical ceiling for a document-embedded image (a 600-DPI A4 scan is ~35 MP,
 * well under this). A crafted 60000×60000 header (~3.6e9 px → ~14 GB RGBA)
 * exceeds the budget by ~50× and is refused before any allocation. With both
 * axes ≤ MAX_RASTER_DIMENSION the product stays ≤ ~1.07e9 — exact in an
 * IEEE-754 double — so a plain numeric comparison suffices (no BigInt).
 */
export const MAX_RASTER_PIXELS = 1 << 26; // 67_108_864 px = 64 MP

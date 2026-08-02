// ── Shared raster pixel-dimension budget (DoS / decode-bomb guard) ───────────
//
// One source of truth for the caps that bound how large a decoded raster surface
// may be, used both by the metafile-embedded DIB decoder (`./dib.ts`) and by the
// pre-`createImageBitmap` header sniff (`./raster-dimensions.ts`). Keeping the
// two paths on the SAME numbers means a raster that is refused as a standalone
// blip is also refused when embedded in a WMF/EMF, and vice versa.

/**
 * Implementation hard ceiling for one raster axis. This rejects obviously huge
 * inputs early; it is not a portable browser capability claim. A runtime or
 * device may impose a lower canvas/decode limit, whose ordinary decode/draw
 * failure remains possible below this ceiling.
 */
export const MAX_RASTER_DIMENSION = 32767;

/**
 * Pixel budget for one decoded raster: 32 MP (2^25 px). A decoded surface is
 * `width × height × 4` bytes of RGBA, so this bounds one bitmap to 128 MiB.
 * A crafted 60000×60000 header (~3.6e9 px → ~14 GB RGBA)
 * is refused before any allocation. With both
 * axes ≤ MAX_RASTER_DIMENSION the product stays ≤ ~1.07e9 — exact in an
 * IEEE-754 double — so a plain numeric comparison suffices (no BigInt).
 */
export const MAX_RASTER_PIXELS = 1 << 25; // 33_554_432 px = 32 MP / 128 MiB RGBA

/** Maximum decoded RGBA ownership retained or leased per document. */
export const MAX_DECODED_IMAGE_BYTES = MAX_RASTER_PIXELS * 4;

/** Keep simultaneous browser decoders bounded even before exact pixels exist. */
export const MAX_CONCURRENT_IMAGE_DECODES = 2;

export type OoxmlDecodedImageLimitMetric = 'image-pixels' | 'active-decoded-bytes';

/** Catchable hard-quota crossing for decoded image surfaces. */
export class OoxmlDecodedImageLimitError extends RangeError {
  readonly code = 'ooxml-decoded-image-limit' as const;

  constructor(
    readonly metric: OoxmlDecodedImageLimitMetric,
    readonly limit: number,
    readonly observed: number,
  ) {
    super(`OOXML decoded image limit exceeded: ${metric} ${observed} > ${limit}`);
    this.name = 'OoxmlDecodedImageLimitError';
    Object.setPrototypeOf(this, OoxmlDecodedImageLimitError.prototype);
  }
}

export function isOoxmlDecodedImageLimitError(
  error: unknown,
): error is OoxmlDecodedImageLimitError {
  return error instanceof OoxmlDecodedImageLimitError
    || (!!error && typeof error === 'object'
      && (error as { code?: unknown }).code === 'ooxml-decoded-image-limit');
}

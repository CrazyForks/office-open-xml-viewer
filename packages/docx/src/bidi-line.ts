// Per-line bidi ordering for the docx renderer.
//
// `buildSegments` splits each run's text into space-delimited word pieces, so
// WITHIN a run Arabic joining never crosses a segment boundary (a mid-word
// run split — e.g. one letter bolded — still seams at the run boundary; that
// pre-existing limitation is tracked for Phase 4). That lets us reorder at SEGMENT
// granularity (1:1 with the laid-out segments — every per-segment property is
// preserved) using the shared UAX#9 engine, and let Canvas shape/mirror each
// segment internally when it is drawn with `ctx.direction` set to the segment's
// resolved direction. Inline objects (image / math / tab) participate as a
// single neutral object-replacement character.

import { getDefaultBidiEngine } from '@silurus/ooxml-core';

/** Strong-RTL scripts (Hebrew, Arabic, Syriac, Thaana, NKo, Samaritan, …) +
 *  Arabic presentation forms. Used only as a cheap gate to decide whether a
 *  line needs the (exact) bidi pass at all — never for ordering itself. */
const RTL_GATE =
  // strong-RTL blocks incl. presentation forms, Plane-1 RTL blocks
  // (Phoenician..Old Hungarian U+10800-10FFF; Mende Kikakui/Adlam/Arabic
  // Math U+1E800-1EFFF), and RTL-implicating controls (RLM/RLE/RLO/RLI).
  /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF\u200F\u202B\u202E\u2067]|[\u{10800}-\u{10FFF}\u{1E800}-\u{1EFFF}]/u;

/** A laid-out segment as seen here: only its optional text matters for bidi.
 *  Typed as `unknown` element so the renderer's LayoutSeg union (whose image /
 *  math / tab members carry no `text`) assigns cleanly. */
const segText = (s: unknown): string | undefined => {
  const t = (s as { text?: unknown }).text;
  return typeof t === 'string' ? t : undefined;
};

/** Cheap test: does this run of segments contain any strong-RTL character? */
export function segmentsHaveRtl(segments: readonly unknown[]): boolean {
  for (const s of segments) {
    const t = segText(s);
    if (t !== undefined && RTL_GATE.test(t)) return true;
  }
  return false;
}

export interface LineVisualOrder {
  /** Logical segment indices in visual (left-to-right) order. */
  order: number[];
  /** Per-LOGICAL-index resolved direction (true = RTL) for `ctx.direction`. */
  rtl: boolean[];
}

const OBJECT_PLACEHOLDER = '￼'; // OBJECT REPLACEMENT CHARACTER (bidi class ON)

/**
 * Compute the visual draw order of a line's segments under `baseRtl`. Text
 * segments contribute their text; non-text segments contribute one neutral
 * placeholder so they take the surrounding direction. Each segment is assigned
 * the embedding level of its first code unit (segments are single-script in
 * practice because they are space-split); Canvas resolves any residual
 * intra-segment bidi when the slice is drawn with the matching `ctx.direction`.
 */
export function computeLineVisualOrder(
  segments: readonly unknown[],
  baseRtl: boolean,
): LineVisualOrder {
  const n = segments.length;
  if (n === 0) return { order: [], rtl: [] };

  let full = '';
  const segStart: number[] = new Array(n);
  for (let i = 0; i < n; i++) {
    segStart[i] = full.length;
    const t = segText(segments[i]) ?? '';
    full += t.length > 0 ? t : OBJECT_PLACEHOLDER;
  }

  const engine = getDefaultBidiEngine();
  const { levels, paragraphLevel } = engine.computeLevels(full, baseRtl ? 'rtl' : 'ltr');

  const segLevels = new Uint8Array(n);
  for (let i = 0; i < n; i++) {
    const lvl = levels[segStart[i]];
    // 255 = removed by X9 (no glyph); fall back to the paragraph level.
    segLevels[i] = lvl === 255 ? paragraphLevel : lvl;
  }

  const order = engine.reorderVisual(segLevels, 0, n);
  const rtl: boolean[] = new Array(n);
  for (let i = 0; i < n; i++) rtl[i] = (segLevels[i] & 1) === 1;
  return { order, rtl };
}

/** Physical edge a line aligns to, resolving logical start/end against base direction. */
export type AlignEdge = 'left' | 'right' | 'center' | 'justify';

/**
 * Resolve a paragraph's `w:jc` value (and base direction) to a physical edge.
 * ALL edge values are logical in WordprocessingML: `start`/`end` by definition
 * (§17.18.44), and the transitional `left`/`right` are defined as
 * "semantically equivalent to start/end" (ECMA-376 Part 4 §14.11.2) — so every
 * edge flips under an RTL base. An unset alignment defaults to the leading
 * (logical-start) edge.
 */
export function resolveAlignEdge(alignment: string | undefined, baseRtl: boolean): AlignEdge {
  switch (alignment) {
    case 'center':
      return 'center';
    case 'both':
    case 'justify':
    case 'distribute':
      return 'justify';
    case 'end':
    case 'right':
      return baseRtl ? 'left' : 'right';
    case 'start':
    case 'left':
    case undefined:
    default:
      return baseRtl ? 'right' : 'left';
  }
}

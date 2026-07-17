// Float-wrap geometry for DOCX anchors (ECMA-376 §20.4.2.x).
//
// Pure layout math: given the floats active on a page (as FloatRect exclusion
// records) it answers "where may this line sit?" and "where must this new float
// be re-seated to avoid a clash?". No canvas/drawing or document-model deps, so
// it can be unit-reasoned and shared by the renderer and the paginator.
//
// IMPORTANT: which parts of the behavior here are ECMA-376-mandated and which
// are documented Office-compatibility observations is recorded inline on
// resolveFloatOverlap and resolveLineFloatWindow. Compatibility behavior must
// stay evidence-backed and apply to the whole OOXML construct, never one file.

import {
  compilePolygonWrap,
  polygonBandExactIntervalFunctions,
  polygonLineTopEventYPts,
  projectPolygonExactLineIntervals,
  type CompiledPolygonWrap,
} from './polygon-wrap.js';
import {
  compareExactRational,
  decodeBinary64,
  exactRationalToNumberDown,
  exactRationalToNumber,
  exactRationalToNumberUp,
  type ExactRational,
} from './exact-geometry.js';
import {
  axisAlignedRectsOverlap,
  resolveAxisAlignedOverlap,
} from './axis-aligned-overlap.js';

function unreducedExactFromNumber(value: number): ExactRational {
  const decoded = decodeBinary64(value);
  return decoded.exponent >= 0
    ? {
        numerator: decoded.coefficient << BigInt(decoded.exponent),
        denominator: 1n,
      }
    : {
        numerator: decoded.coefficient,
        denominator: 1n << BigInt(-decoded.exponent),
      };
}

function addUnreducedExact(left: ExactRational, right: ExactRational): ExactRational {
  return {
    numerator: left.numerator * right.denominator
      + right.numerator * left.denominator,
    denominator: left.denominator * right.denominator,
  };
}

function subtractUnreducedExact(left: ExactRational, right: ExactRational): ExactRational {
  return {
    numerator: left.numerator * right.denominator
      - right.numerator * left.denominator,
    denominator: left.denominator * right.denominator,
  };
}

function multiplyUnreducedExact(left: ExactRational, right: ExactRational): ExactRational {
  return {
    numerator: left.numerator * right.numerator,
    denominator: left.denominator * right.denominator,
  };
}

function divideUnreducedExact(left: ExactRational, right: ExactRational): ExactRational {
  const negative = right.numerator < 0n;
  return {
    numerator: (negative ? -left.numerator : left.numerator) * right.denominator,
    denominator: left.denominator * (negative ? -right.numerator : right.numerator),
  };
}

function exactBinary64Midpoint(left: number, right: number): number {
  const exactLeft = unreducedExactFromNumber(left);
  const exactRight = unreducedExactFromNumber(right);
  return exactRationalToNumber({
    numerator: exactLeft.numerator * exactRight.denominator
      + exactRight.numerator * exactLeft.denominator,
    denominator: 2n * exactLeft.denominator * exactRight.denominator,
  });
}

/**
 * Floating object that affects text wrap on the current page. `FloatRect` is a
 * historical name: tight/through records carry their authoritative polygon;
 * xLeft/xRight/yTop/yBottom are its padded acquisition bounds, not a replacement
 * rectangle for polygon line queries.
 */
export interface FloatRect {
  /** What kind of object reserved this float. Used to scope overlap avoidance:
   *  ECMA-376 §17.4.56 (tblOverlap="never") only forbids a floating table from
   *  overlapping OTHER FLOATING TABLES — not DrawingML anchors (§20.4.2.3) or
   *  text frames. resolveFloatOverlap reads this to limit a never-overlap
   *  table's blockers to kind==='table'. 'shape' = DrawingML wp:anchor shape,
   *  'frame' = <w:framePr> text frame; both also cover anchor images. */
  kind: 'table' | 'shape' | 'frame';
  mode: 'square' | 'topAndBottom';
  /** Exact retained wrap semantics; `mode` remains the coarse legacy routing key. */
  authoredWrap?: 'square' | 'tight' | 'through' | 'topAndBottom';
  anchorOccurrenceId?: string;
  acquisitionOccurrenceId?: string;
  wrapPolygon?: readonly Readonly<{ xPt: number; yPt: number }>[];
  /** Hex key of the image bitmap (used to defer drawing until final Y is known). */
  imageKey: string;
  /** Absolute canvas X of the image box (without dist padding). */
  imageX: number;
  imageY: number;
  imageW: number;
  imageH: number;
  /** Padded exclusion rectangle for text wrap. */
  xLeft: number;
  xRight: number;
  yTop: number;
  yBottom: number;
  /** ST_WrapText: "bothSides" | "left" | "right" | "largest". */
  side: string;
  /** dist* padding (px) — needed when displacing a float to keep its exclusion
   *  padding when re-seating next to a blocking float (ECMA-376 §20.4.2.x). */
  distLeft: number;
  distRight: number;
  distTop: number;
  distBottom: number;
  /** Identifier of the anchoring paragraph. Used only by the observed Office
   *  compatibility rule under allowOverlap=true: floats with the SAME paraId
   *  never displace each other, while different-paragraph floats do. ECMA-376
   *  does not define this scoping; see resolveFloatOverlap. */
  paraId: number;
  /** true once the image itself has been drawn (drawn after its paragraph lays out). */
  drawn: boolean;
}

export type WrapSide = 'bothSides' | 'left' | 'right' | 'largest';

export function normalizeWrapSide(side: string | null | undefined): WrapSide {
  switch (side) {
    case 'left':
    case 'right':
    case 'largest':
    case 'bothSides': return side;
    default: return 'bothSides';
  }
}

/** A horizontal interval [l, r] in absolute canvas px. */
export interface Gap {
  l: number;
  r: number;
}

// ── Float-layout tolerances (px) ──────────────────────────────────────────────
// Sub-pixel slack used so floating-point coordinate noise (margin/anchor/dist
// arithmetic at the current scale) doesn't read as a real overlap or a real gap.

/** Overlap epsilon: two exclusion rects must overlap by MORE than this to count
 *  as intersecting, so coincident/touching edges (and FP noise) are not a clash. */
export const FLOAT_OVERLAP_EPS = 0.01;

/** Slack added to the page-right edge when testing whether a displaced float
 *  still fits horizontally — a float ending within this many px of the page edge
 *  is treated as fitting (it would otherwise be pushed down by FP rounding).
 *  Looser than FLOAT_OVERLAP_EPS because it guards a half-pixel rounding of a
 *  full-width displacement, not an edge-touch test. */
export const FLOAT_PAGE_RIGHT_SLACK = 0.5;

/** Minimum horizontal space (pt) a free side-gap must have before Word will
 *  START a CONTENT (text / inline-object) line beside a square float, rather than
 *  flowing it below the float band. Measured — NOT from ECMA-376, which mandates
 *  no side-gap minimum (§20.4.2.17 only says text wraps around the rectangle;
 *  §17.18.3 `<w:br w:clear>` is the sole spec-mandated flow onto a float-free
 *  region). Controlled Office comparisons documented in issue #676 establish
 *  an exact 1-inch (1440-twip) rule: 70pt flows below while 72pt starts beside.
 *  The boundary is independent of text, font size, and line spacing, so it is
 *  an absolute width rather than an em- or line-height-relative quantity. The
 *  evidence covers square wrapping only; tight/through use their
 *  §20.4.2.18/.19 polygon openings without inheriting this compatibility
 *  policy. Callers convert to px with `WORD_MIN_LINE_START_PT * scale`
 *  (renderer scale is px/pt).
 *
 *  SCOPE — content lines only. An EMPTY paragraph-mark line (a literally-empty or
 *  anchor-only paragraph's pilcrow, no width-bearing content) does NOT obey this
 *  1-inch rule: Word keeps such a mark beside a float whenever the gap can hold
 *  the pilcrow itself, dropping it below only when the gap is narrower than that
 *  (effectively a full-width float band). Office comparisons in issue #676 show
 *  that a full-width band moves the mark below, while a sub-inch side gap that
 *  can hold the pilcrow keeps authoring blank-line marks beside the float. The
 *  narrow threshold is therefore the paragraph-mark em; it governs the
 *  literally-empty paths — the paint pass `resolveEmptyMarkTop` and paginator
 *  mirror `flowMarkLine` — plus the anchorHost-only metric line inside
 *  `layoutLines`. */
export const WORD_MIN_LINE_START_PT = 72;

/** Tolerance (pt) subtracted from the 1-inch requirement when testing a side
 *  gap, to make Word's INCLUSIVE ≥ 1-inch boundary robust to coordinate noise.
 *  Word places a line beside a float at a gap of exactly 1 inch (issue #676).
 *  But a gap
 *  that is nominally 1 inch is computed as content-width − frame-width through
 *  twip→EMU→px conversions and lands slightly under 72: this renderer computes
 *  71.963716pt for the 72.0pt frame (a ~0.036pt deficit — sub-twip conversion
 *  rounding, not pure IEEE-754). Without tolerance the inclusive boundary
 *  flips to below and disagrees with Word. One twip (1/20 pt = 0.05pt) is the
 *  authoring granularity of a frame width, so a gap short of 1 inch by less
 *  than one twip is treated as exactly 1 inch. One twip covers the observed
 *  0.036pt deficit (a half twip, 0.025pt, would NOT) yet is ≪ the 2pt step
 *  that discriminates the fixtures (70pt stays below, 72pt goes beside), so it
 *  never promotes a genuinely sub-inch gap. Applied in the render's px space
 *  as `× scale` (see resolveLineFloatWindow). Same rationale as
 *  FLOAT_PAGE_RIGHT_SLACK: a tolerance sized to the coordinate-rounding
 *  granularity it absorbs. */
export const LINE_START_GAP_EPS_PT = 0.05; // one twip (1/20 pt)

/** The square-only compatibility width passed separately from polygon geometry:
 *  Word's 1-inch minimum
 *  side-gap, minus the one-twip rounding tolerance, at the render scale (px/pt).
 *  Single source of truth so the paint pass and both paginator mirrors agree
 *  bit-for-bit on the flow/beside decision. Empty paragraph-mark lines use the
 *  narrower `paragraphMarkEmPx` threshold instead (see WORD_MIN_LINE_START_PT's
 *  SCOPE note). See WORD_MIN_LINE_START_PT and LINE_START_GAP_EPS_PT (issue
 *  #676). */
export function wordMinLineStartPx(scale: number): number {
  return (WORD_MIN_LINE_START_PT - LINE_START_GAP_EPS_PT) * scale;
}

/** Minimum width (px) a polygon free gap must have to hold a line start. It also
 *  floors a zero-width direct probe so
 *  a `requiredWidth === 0` call still rejects sub-pixel slivers between
 *  full-width floats. Square queries additionally apply the independently
 *  supplied compatibility width; tight/through deliberately do not. */
export const MIN_LINE_GAP = 1;

export function isWrapFloat(mode?: string | null): boolean {
  return mode === 'square' || mode === 'topAndBottom' || mode === 'tight' || mode === 'through';
}

/**
 * Does float `f`'s horizontal extent overlap the paragraph/column text band
 * [paraXLeft, paraXRight]? Touching edges (within FLOAT_OVERLAP_EPS) do not count.
 *
 * ECMA-376 §20.4.2.17 (wrapSquare) and §20.4.2.20 (wrapTopAndBottom) both exclude
 * text only where the object is horizontally placed ("text shall wrap around …
 * THIS OBJECT"). Floats are registered in ABSOLUTE page coordinates and the page
 * float set is shared across a section's newspaper columns (§17.6.4), so a float
 * anchored in one column must be filtered out for a line laid out in another
 * column that it does not horizontally overlap. Both wrap modes route through
 * this one predicate so they share identical column-scoping semantics.
 */
export function floatOverlapsColumnX(
  f: FloatRect,
  paraXLeft: number,
  paraXRight: number,
): boolean {
  return f.xRight > paraXLeft + FLOAT_OVERLAP_EPS && f.xLeft < paraXRight - FLOAT_OVERLAP_EPS;
}

/** Two exclusion rects intersect (strict overlap, touching edges allowed). */
export function rectsOverlap(
  aL: number, aR: number, aT: number, aB: number,
  bL: number, bR: number, bT: number, bB: number,
): boolean {
  return axisAlignedRectsOverlap(
    { left: aL, right: aR, top: aT, bottom: aB },
    { left: bL, right: bR, top: bT, bottom: bB },
    FLOAT_OVERLAP_EPS,
  );
}

/**
 * Widest free horizontal interval within [left, right] after removing the
 * `blocked` spans. Returns null when nothing is free. Factored out of
 * resolveLineFloatWindow so the caller holds a properly-typed Gap (the previous
 * inline closure form forced TS to narrow `best` to never, requiring casts).
 */
export function widestFreeGap(blocked: Gap[], left: number, right: number): Gap | null {
  const spans = blocked.slice().sort((a, b) => a.l - b.l);
  let cursor = left;
  let best: Gap | null = null;
  const consider = (l: number, r: number): void => {
    // Adopt only when strictly wider than the current best (0 when none yet),
    // so a zero/negative-width gap never becomes `best`. Matches the prior inline
    // form `r - l > (best ? best.r - best.l : 0)`.
    if (r - l > (best ? best.r - best.l : 0)) best = { l, r };
  };
  for (const b of spans) {
    if (b.l > cursor) consider(cursor, Math.min(b.l, right));
    cursor = Math.max(cursor, Math.min(b.r, right));
    if (cursor >= right) break;
  }
  if (cursor < right) consider(cursor, right);
  return best;
}

interface PreparedFloatRect {
  readonly rect: Readonly<FloatRect>;
  readonly polygon: CompiledPolygonWrap | null;
  readonly wrapMaximumLeftPt: number;
  readonly wrapMaximumRightPt: number;
}

export interface PreparedFloatWrap {
  readonly floats: readonly PreparedFloatRect[];
}

const sweepEventCache = new WeakMap<PreparedFloatWrap, Map<number, readonly number[]>>();

interface ExactAttributedGap {
  readonly l: ExactRational;
  readonly r: ExactRational;
  readonly leftSquareBoundary: boolean;
  readonly rightSquareBoundary: boolean;
}

/** Snapshot and compile geometry once at the line-wrap oracle boundary. */
export function prepareFloatWrap(floats: readonly FloatRect[]): PreparedFloatWrap {
  const prepared = floats.map((float): PreparedFloatRect => {
    const rect = Object.freeze({
      ...float,
      ...(float.wrapPolygon
        ? { wrapPolygon: Object.freeze(float.wrapPolygon.map((point) => Object.freeze({ ...point }))) }
        : {}),
    });
    const polygon = rect.authoredWrap === 'tight' || rect.authoredWrap === 'through'
      ? compilePolygonWrap({
          kind: rect.authoredWrap,
          imageKey: rect.imageKey,
          points: rect.wrapPolygon,
          xLeftPt: rect.xLeft,
          xRightPt: rect.xRight,
          yTopPt: rect.yTop,
          yBottomPt: rect.yBottom,
        })
      : null;
    return Object.freeze({
      rect,
      polygon,
      // §20.4.3.7 chooses "largest" from the positioned object, not from a
      // contour slice. Retain the full padded wrap extent so every line-band
      // projection observes one stable side even when the polygon is skewed.
      wrapMaximumLeftPt: polygon
        ? Math.min(rect.xLeft, polygon.polygonLeftPt)
        : rect.xLeft,
      wrapMaximumRightPt: polygon
        ? Math.max(rect.xRight, polygon.polygonRightPt)
        : rect.xRight,
    });
  });
  const result = Object.freeze({ floats: Object.freeze(prepared) });
  sweepEventCache.set(result, new Map());
  return result;
}

export interface LineFloatReference {
  readonly xLeftPt: number;
  readonly xRightPt: number;
  readonly readingDirection: 'ltr' | 'rtl';
}

function effectiveWrapSide(
  prepared: PreparedFloatRect,
  reference: LineFloatReference,
): Exclude<WrapSide, 'largest'> {
  const side = normalizeWrapSide(prepared.rect.side);
  if (side !== 'largest') return side;
  const leftWidth = subtractUnreducedExact(
    unreducedExactFromNumber(prepared.wrapMaximumLeftPt),
    unreducedExactFromNumber(reference.xLeftPt),
  );
  const rightWidth = subtractUnreducedExact(
    unreducedExactFromNumber(reference.xRightPt),
    unreducedExactFromNumber(prepared.wrapMaximumRightPt),
  );
  const widthComparison = compareExactRational(leftWidth, rightWidth);
  return widthComparison === 0
    ? (reference.readingDirection === 'ltr' ? 'left' : 'right')
    : widthComparison > 0 ? 'left' : 'right';
}

function floatBlockedIntervals(
  prepared: PreparedFloatRect,
  lineTopPt: number,
  probeHeightPt: number,
  paragraphLeftPt: number,
  paragraphRightPt: number,
  reference: LineFloatReference,
): ExactAttributedGap[] {
  const { rect: float, polygon } = prepared;
  const projected = polygon
    ? projectPolygonExactLineIntervals(polygon, lineTopPt, probeHeightPt)
    : [{
        l: unreducedExactFromNumber(float.xLeft),
        r: unreducedExactFromNumber(float.xRight),
      }];
  if (projected.length === 0) return [];
  const squareBoundary = polygon === null;
  const objectLeft = projected.reduce((minimum, interval) =>
    compareExactRational(interval.l, minimum) < 0 ? interval.l : minimum,
  projected[0]!.l);
  const objectRight = projected.reduce((maximum, interval) =>
    compareExactRational(interval.r, maximum) > 0 ? interval.r : maximum,
  projected[0]!.r);
  switch (effectiveWrapSide(prepared, reference)) {
    case 'left': return [{
      l: objectLeft,
      r: unreducedExactFromNumber(paragraphRightPt),
      leftSquareBoundary: squareBoundary,
      rightSquareBoundary: false,
    }];
    case 'right': return [{
      l: unreducedExactFromNumber(paragraphLeftPt),
      r: objectRight,
      leftSquareBoundary: false,
      rightSquareBoundary: squareBoundary,
    }];
    case 'bothSides': return projected.map((interval) => ({
      ...interval,
      leftSquareBoundary: squareBoundary,
      rightSquareBoundary: squareBoundary,
    }));
  }
}

/**
 * Resolve where a single line box may sit relative to the page's active floats.
 *
 * Given the line's intended top Y and minimum horizontal width, this convenience
 * boundary compiles raw geometry and returns the earliest Y plus horizontal
 * window. Production line layout uses the prepared variant below so compilation
 * occurs once per oracle/acquisition rather than in the per-line hot path.
 *
 * Two ECMA-376 wrap rules are applied, in order:
 *   1. topAndBottom floats (§20.4.2.20): a line intersecting one is pushed below
 *      it — text never sits beside a topAndBottom object.
 *   2. square floats (§20.4.2.17): text wraps around the float's rect + dist
 *      padding; tight/through use their compiled polygons (§20.4.2.18/.19).
 *      Multiple objects are composed only after each `largest` object selects
 *      its own permitted side under §20.4.3.7. The widest remaining gap wins.
 *
 * The one-inch line-start rule is a Word compatibility policy observed for
 * square objects, not polygon geometry and not an ECMA-376 mandate. The prepared
 * API therefore receives square and polygon requirements separately.
 */
export function resolveLineFloatWindow(
  topY: number,
  requiredWidth: number,
  probeH: number,
  paraX: number,
  maxWidth: number,
  floats: FloatRect[],
  // The paragraph's RAW COLUMN band, distinct from the indented text band
  // [paraX, paraX + maxWidth]. Step 1 (topAndBottom) gates by the COLUMN band —
  // §20.4.2.20 blocks the FULL column where the object sits, including the
  // paragraph's indent margins — while step 2 (square side-gap) keeps gating by
  // the narrower indented text band (§20.4.2.17). Defaults to the indented band
  // so a direct unit caller that has no separate column band stays correct.
  columnXLeftPt: number = paraX,
  columnXRightPt: number = paraX + maxWidth,
  reference: LineFloatReference = {
    xLeftPt: paraX,
    xRightPt: paraX + maxWidth,
    readingDirection: 'ltr',
  },
): { topY: number; xOffset: number; maxWidth: number } {
  return resolvePreparedLineFloatWindow(
    topY,
    requiredWidth,
    probeH,
    paraX,
    maxWidth,
    prepareFloatWrap(floats),
    columnXLeftPt,
    columnXRightPt,
    reference,
    requiredWidth,
  );
}

function finiteStructuralEvents(
  probeH: number,
  prepared: PreparedFloatWrap,
): readonly number[] {
  const cache = sweepEventCache.get(prepared);
  if (!cache) throw new Error('Prepared float geometry omitted its sweep cache');
  const cached = cache.get(probeH);
  if (cached) return cached;
  const candidates = new Set<number>();
  const appendY = (value: number): void => {
    if (Number.isFinite(value)) candidates.add(value);
  };
  for (const { rect, polygon } of prepared.floats) {
    appendY(exactRationalToNumberUp(subtractUnreducedExact(
      unreducedExactFromNumber(rect.yTop),
      unreducedExactFromNumber(probeH),
    )));
    appendY(rect.yBottom);
    if (!polygon) continue;
    for (const eventYPt of polygonLineTopEventYPts(polygon, probeH)) appendY(eventYPt);
  }
  const events = Object.freeze([...candidates].sort((left, right) => left - right));
  cache.set(probeH, events);
  return events;
}

function mergeAttributedIntervals(
  intervals: readonly ExactAttributedGap[],
): ExactAttributedGap[] {
  const sorted = intervals
    .filter((interval) => compareExactRational(interval.r, interval.l) > 0)
    .slice()
    .sort((left, right) =>
      compareExactRational(left.l, right.l)
        || compareExactRational(left.r, right.r));
  const merged: Array<{
    l: ExactRational;
    r: ExactRational;
    leftSquareBoundary: boolean;
    rightSquareBoundary: boolean;
  }> = [];
  for (const interval of sorted) {
    const previous = merged.at(-1);
    if (!previous || compareExactRational(interval.l, previous.r) > 0) {
      merged.push({ ...interval });
      continue;
    }
    if (compareExactRational(interval.l, previous.l) === 0) {
      previous.leftSquareBoundary = previous.leftSquareBoundary
        && interval.leftSquareBoundary;
    }
    const rightComparison = compareExactRational(interval.r, previous.r);
    if (rightComparison > 0) {
      previous.r = interval.r;
      previous.rightSquareBoundary = interval.rightSquareBoundary;
    } else if (rightComparison === 0) {
      previous.rightSquareBoundary = previous.rightSquareBoundary
        && interval.rightSquareBoundary;
    }
  }
  return merged;
}

function widestUsableFreeGap(
  blocked: readonly ExactAttributedGap[],
  left: number,
  right: number,
  polygonRequiredWidth: number,
  squareRequiredWidth: number,
): Readonly<{
  l: ExactRational;
  r: ExactRational;
  squareConstrained: boolean;
}> | null {
  const merged = mergeAttributedIntervals(blocked);
  const exactLeft = unreducedExactFromNumber(left);
  const exactRight = unreducedExactFromNumber(right);
  const gaps: Array<{
    l: ExactRational;
    r: ExactRational;
    squareConstrained: boolean;
  }> = [];
  const consider = (
    l: ExactRational,
    r: ExactRational,
    squareConstrained: boolean,
  ): void => {
    const clippedLeft = compareExactRational(exactLeft, l) >= 0 ? exactLeft : l;
    const clippedRight = compareExactRational(exactRight, r) <= 0 ? exactRight : r;
    if (compareExactRational(clippedRight, clippedLeft) > 0) {
      gaps.push({ l: clippedLeft, r: clippedRight, squareConstrained });
    }
  };
  let cursor = exactLeft;
  let cursorSquareBoundary = false;
  for (const interval of merged) {
    if (compareExactRational(interval.r, exactLeft) <= 0) {
      cursorSquareBoundary = interval.rightSquareBoundary;
      continue;
    }
    if (compareExactRational(interval.l, exactRight) >= 0) {
      consider(cursor, exactRight, cursorSquareBoundary);
      cursor = exactRight;
      break;
    }
    if (compareExactRational(interval.l, cursor) > 0) {
      consider(
        cursor,
        interval.l,
        cursorSquareBoundary || interval.leftSquareBoundary,
      );
    }
    const rightComparison = compareExactRational(interval.r, cursor);
    if (rightComparison > 0) {
      cursor = interval.r;
      cursorSquareBoundary = interval.rightSquareBoundary;
    } else if (rightComparison === 0) {
      cursorSquareBoundary = cursorSquareBoundary && interval.rightSquareBoundary;
    }
    if (compareExactRational(cursor, exactRight) >= 0) break;
  }
  if (compareExactRational(cursor, exactRight) < 0) {
    consider(cursor, exactRight, cursorSquareBoundary);
  }
  let widestWidth: ExactRational = { numerator: 0n, denominator: 1n };
  for (const gap of gaps) {
    const width = subtractUnreducedExact(gap.r, gap.l);
    if (compareExactRational(width, widestWidth) > 0) widestWidth = width;
  }
  for (const gap of gaps) {
    const width = subtractUnreducedExact(gap.r, gap.l);
    if (compareExactRational(width, widestWidth) !== 0) continue;
    const requirement = Math.max(
      MIN_LINE_GAP,
      gap.squareConstrained ? squareRequiredWidth : polygonRequiredWidth,
    );
    if (compareExactRational(width, unreducedExactFromNumber(requirement)) >= 0) {
      return {
        l: gap.l,
        r: gap.r,
        squareConstrained: gap.squareConstrained,
      };
    }
  }
  return null;
}

function lineWindowAtY(
  topY: number,
  probeH: number,
  paraXLeft: number,
  paraXRight: number,
  maxWidth: number,
  prepared: PreparedFloatWrap,
  columnXLeftPt: number,
  columnXRightPt: number,
  reference: LineFloatReference,
  polygonRequiredWidth: number,
  squareRequiredWidth: number,
): { topY: number; xOffset: number; maxWidth: number } | null {
  const exactTop = unreducedExactFromNumber(topY);
  const exactBottom = addUnreducedExact(
    exactTop,
    unreducedExactFromNumber(probeH),
  );
  const exactRectY = (value: number): ExactRational => unreducedExactFromNumber(value);
  if (prepared.floats.some(({ rect }) => rect.mode === 'topAndBottom'
    && floatOverlapsColumnX(rect as FloatRect, columnXLeftPt, columnXRightPt)
    && compareExactRational(exactBottom, exactRectY(rect.yTop)) > 0
    && compareExactRational(exactTop, exactRectY(rect.yBottom)) < 0)) return null;
  const blocked: ExactAttributedGap[] = [];
  for (const float of prepared.floats) {
    const { rect } = float;
    if (rect.mode !== 'square') continue;
    if (compareExactRational(exactBottom, exactRectY(rect.yTop)) <= 0
      || compareExactRational(exactTop, exactRectY(rect.yBottom)) >= 0) continue;
    if (!floatOverlapsColumnX(rect as FloatRect, paraXLeft, paraXRight)) continue;
    const intervals = floatBlockedIntervals(
      float, topY, probeH, paraXLeft, paraXRight, reference,
    );
    if (intervals.length === 0) continue;
    blocked.push(...intervals);
  }
  if (blocked.length === 0) return { topY, xOffset: 0, maxWidth };
  const best = widestUsableFreeGap(
    blocked,
    paraXLeft,
    paraXRight,
    polygonRequiredWidth,
    squareRequiredWidth,
  );
  if (!best) return null;
  const zero: ExactRational = { numerator: 0n, denominator: 1n };
  const xOffsetCandidate = subtractUnreducedExact(
    best.l,
    unreducedExactFromNumber(paraXLeft),
  );
  const xOffset = compareExactRational(xOffsetCandidate, zero) > 0
    ? xOffsetCandidate
    : zero;
  const exactParagraphLeft = unreducedExactFromNumber(paraXLeft);
  let xOffsetNumber = exactRationalToNumberUp(xOffset);
  let returnedStartNumber = paraXLeft + xOffsetNumber;
  let returnedStart = unreducedExactFromNumber(returnedStartNumber);
  if (compareExactRational(returnedStart, best.l) < 0) {
    const targetStartNumber = exactRationalToNumberUp(best.l);
    xOffsetNumber = exactRationalToNumberUp(subtractUnreducedExact(
      unreducedExactFromNumber(targetStartNumber),
      exactParagraphLeft,
    ));
    returnedStartNumber = paraXLeft + xOffsetNumber;
    returnedStart = unreducedExactFromNumber(returnedStartNumber);
  }
  if (compareExactRational(returnedStart, best.l) < 0) {
    throw new Error('Exact float window could not represent a contained start');
  }
  const exactParagraphRight = unreducedExactFromNumber(paraXRight);
  const exactAvailableEnd =
    compareExactRational(best.r, exactParagraphRight) <= 0
      ? best.r
      : exactParagraphRight;
  const returnedEndNumber = exactRationalToNumberDown(exactAvailableEnd);
  const availableWidth = subtractUnreducedExact(
    unreducedExactFromNumber(returnedEndNumber),
    returnedStart,
  );
  const nonnegativeWidth = compareExactRational(availableWidth, zero) > 0
    ? availableWidth
    : zero;
  const maxWidthNumber = exactRationalToNumberDown(nonnegativeWidth);
  const reconstructedEnd = unreducedExactFromNumber(
    returnedStartNumber + maxWidthNumber,
  );
  if (compareExactRational(reconstructedEnd, exactAvailableEnd) > 0) {
    throw new Error('Exact float window could not represent a contained end');
  }
  return {
    topY,
    xOffset: xOffsetNumber,
    maxWidth: maxWidthNumber,
  };
}

interface AffineBoundary {
  readonly exact: Readonly<{
    slope: ExactRational;
    intercept: ExactRational;
  }>;
  readonly square: boolean;
}

interface AffineBlockedInterval {
  readonly left: AffineBoundary;
  readonly right: AffineBoundary;
}

export interface LineFloatSweepDiagnostics {
  readonly compiledIntersectionCount: number;
  readonly compiledContourSpanCount: number;
  readonly compileOrderComparisonCount: number;
  /** Edge-slot reads for complete vertex scans plus local intersection updates: O(V² + K). */
  readonly compilePairMembershipVisitCount: number;
  readonly structuralEventCount: number;
  readonly localRootCandidateCount: number;
  readonly localRootEventCount: number;
  readonly evaluatedYCount: number;
}

interface MutableLineFloatSweepDiagnostics {
  compiledIntersectionCount: number;
  compiledContourSpanCount: number;
  compileOrderComparisonCount: number;
  compilePairMembershipVisitCount: number;
  structuralEventCount: number;
  localRootCandidateCount: number;
  localRootEventCount: number;
  evaluatedYCount: number;
}

function exactAffineValueAt(boundary: AffineBoundary, y: number): ExactRational {
  return addUnreducedExact(
    multiplyUnreducedExact(boundary.exact.slope, unreducedExactFromNumber(y)),
    boundary.exact.intercept,
  );
}

function sameAffine(left: AffineBoundary, right: AffineBoundary): boolean {
  return compareExactRational(left.exact.slope, right.exact.slope) === 0
    && compareExactRational(left.exact.intercept, right.exact.intercept) === 0;
}

function compareAtRight(left: AffineBoundary, right: AffineBoundary, y: number): number {
  return compareExactRational(
    exactAffineValueAt(left, y),
    exactAffineValueAt(right, y),
  ) || compareExactRational(left.exact.slope, right.exact.slope);
}

function exactAffineRoot(
  left: AffineBoundary['exact'],
  right: AffineBoundary['exact'],
  delta: number,
): ExactRational | null {
  const slope = subtractUnreducedExact(left.slope, right.slope);
  if (slope.numerator === 0n) return null;
  const intercept = subtractUnreducedExact(left.intercept, right.intercept);
  return divideUnreducedExact(
    subtractUnreducedExact(unreducedExactFromNumber(delta), intercept),
    slope,
  );
}

function appendLocalRoot(
  roots: number[],
  root: ExactRational | null,
  lower: number,
  upper: number,
  diagnostics: MutableLineFloatSweepDiagnostics | null,
): void {
  if (root === null
    || compareExactRational(root, unreducedExactFromNumber(lower)) <= 0
    || compareExactRational(root, unreducedExactFromNumber(upper)) >= 0) return;
  roots.push(exactRationalToNumberUp(root));
  if (diagnostics) diagnostics.localRootCandidateCount += 1;
}

function selectEnvelope(
  candidates: readonly AffineBoundary[],
  kind: 'min' | 'max',
  lower: number,
  upper: number,
  roots: number[],
  diagnostics: MutableLineFloatSweepDiagnostics | null,
): AffineBoundary {
  let winner = candidates[0]!;
  for (const candidate of candidates.slice(1)) {
    const comparison = compareAtRight(candidate, winner, lower);
    if ((kind === 'min' && comparison < 0) || (kind === 'max' && comparison > 0)) {
      winner = candidate;
    }
  }
  let square = winner.square;
  for (const candidate of candidates) {
    if (candidate === winner) continue;
    if (sameAffine(candidate, winner)) {
      square = square && candidate.square;
      continue;
    }
    const canTakeOver = kind === 'min'
      ? compareExactRational(candidate.exact.slope, winner.exact.slope) < 0
      : compareExactRational(candidate.exact.slope, winner.exact.slope) > 0;
    if (canTakeOver) {
      appendLocalRoot(
        roots,
        exactAffineRoot(candidate.exact, winner.exact, 0),
        lower,
        upper,
        diagnostics,
      );
    }
  }
  return { exact: winner.exact, square };
}

function constantBoundary(value: number, square = false): AffineBoundary {
  return {
    exact: {
      slope: { numerator: 0n, denominator: 1n },
      intercept: unreducedExactFromNumber(value),
    },
    square,
  };
}

function objectAffineIntervals(
  prepared: PreparedFloatRect,
  probeH: number,
  lower: number,
  upper: number,
  paragraphLeftPt: number,
  paragraphRightPt: number,
  reference: LineFloatReference,
  roots: number[],
  diagnostics: MutableLineFloatSweepDiagnostics | null,
): AffineBlockedInterval[] {
  const { rect, polygon } = prepared;
  const midpoint = exactBinary64Midpoint(lower, upper);
  const exactMidpoint = unreducedExactFromNumber(midpoint);
  const exactMidpointBottom = addUnreducedExact(
    exactMidpoint,
    unreducedExactFromNumber(probeH),
  );
  if (compareExactRational(
    exactMidpointBottom,
    unreducedExactFromNumber(rect.yTop),
  ) <= 0 || compareExactRational(
    exactMidpoint,
    unreducedExactFromNumber(rect.yBottom),
  ) >= 0) return [];
  const square = polygon === null;
  let intervals: AffineBlockedInterval[] = polygon
    ? polygonBandExactIntervalFunctions(polygon, probeH, lower, upper)
      .map((interval) => ({
        left: { exact: interval.left, square: false },
        right: { exact: interval.right, square: false },
      }))
    : [{
        left: constantBoundary(rect.xLeft, true),
        right: constantBoundary(rect.xRight, true),
      }];
  if (intervals.length === 0) return [];
  const objectLeft = selectEnvelope(
    intervals.map((interval) => interval.left),
    'min', lower, upper, roots, diagnostics,
  );
  const objectRight = selectEnvelope(
    intervals.map((interval) => interval.right),
    'max', lower, upper, roots, diagnostics,
  );
  if (polygon?.kind === 'tight') intervals = [{ left: objectLeft, right: objectRight }];
  switch (effectiveWrapSide(prepared, reference)) {
    case 'left': return [{
      left: objectLeft,
      right: constantBoundary(paragraphRightPt),
    }];
    case 'right': return [{
      left: constantBoundary(paragraphLeftPt),
      right: objectRight,
    }];
    case 'bothSides': return intervals;
  }
}

function mergeAffineIntervalsAtRight(
  intervals: readonly AffineBlockedInterval[],
  lower: number,
  upper: number,
  roots: number[],
  diagnostics: MutableLineFloatSweepDiagnostics | null,
): AffineBlockedInterval[] {
  const sorted = intervals.slice().sort((left, right) =>
    compareAtRight(left.left, right.left, lower)
    || compareAtRight(left.right, right.right, lower));
  for (let index = 0; index + 1 < sorted.length; index += 1) {
    appendLocalRoot(
      roots,
      exactAffineRoot(
        sorted[index]!.left.exact,
        sorted[index + 1]!.left.exact,
        0,
      ),
      lower,
      upper,
      diagnostics,
    );
  }
  const merged: AffineBlockedInterval[] = [];
  for (const interval of sorted) {
    const previous = merged.at(-1);
    if (!previous) {
      merged.push(interval);
      continue;
    }
    appendLocalRoot(
      roots,
      exactAffineRoot(interval.left.exact, previous.right.exact, 0),
      lower,
      upper,
      diagnostics,
    );
    const contact = compareAtRight(interval.left, previous.right, lower);
    if (contact > 0) {
      merged.push(interval);
      continue;
    }
    const right = selectEnvelope(
      [previous.right, interval.right],
      'max', lower, upper, roots, diagnostics,
    );
    const left = sameAffine(previous.left, interval.left)
      ? {
          exact: previous.left.exact,
          square: previous.left.square && interval.left.square,
        }
      : previous.left;
    merged[merged.length - 1] = { left, right };
  }
  return merged;
}

function nextLocalSweepEvent(
  lower: number,
  upper: number,
  probeH: number,
  paraXLeft: number,
  paraXRight: number,
  prepared: PreparedFloatWrap,
  columnXLeftPt: number,
  columnXRightPt: number,
  reference: LineFloatReference,
  polygonRequiredWidth: number,
  squareRequiredWidth: number,
  diagnostics: MutableLineFloatSweepDiagnostics | null,
): number | null {
  const midpoint = exactBinary64Midpoint(lower, upper);
  const exactMidpoint = unreducedExactFromNumber(midpoint);
  const exactMidpointBottom = addUnreducedExact(
    exactMidpoint,
    unreducedExactFromNumber(probeH),
  );
  if (prepared.floats.some(({ rect }) => rect.mode === 'topAndBottom'
    && floatOverlapsColumnX(rect as FloatRect, columnXLeftPt, columnXRightPt)
    && compareExactRational(
      exactMidpointBottom,
      unreducedExactFromNumber(rect.yTop),
    ) > 0
    && compareExactRational(
      exactMidpoint,
      unreducedExactFromNumber(rect.yBottom),
    ) < 0)) return null;
  const roots: number[] = [];
  const intervals: AffineBlockedInterval[] = [];
  for (const float of prepared.floats) {
    const { rect } = float;
    if (rect.mode !== 'square') continue;
    if (!floatOverlapsColumnX(rect as FloatRect, paraXLeft, paraXRight)) continue;
    intervals.push(...objectAffineIntervals(
      float, probeH, lower, upper, paraXLeft, paraXRight,
      reference, roots, diagnostics,
    ));
  }
  if (intervals.length === 0) return null;
  const merged = mergeAffineIntervalsAtRight(
    intervals, lower, upper, roots, diagnostics,
  );
  const pageLeft = constantBoundary(paraXLeft);
  const pageRight = constantBoundary(paraXRight);
  const gaps: Array<{ exactWidth: AffineBoundary['exact'] }> = [];
  const appendGapRoot = (
    left: AffineBoundary,
    right: AffineBoundary,
    squareConstrained: boolean,
  ): void => {
    const exactWidth = {
      slope: subtractUnreducedExact(right.exact.slope, left.exact.slope),
      intercept: subtractUnreducedExact(right.exact.intercept, left.exact.intercept),
    };
    gaps.push({ exactWidth });
    const requirement = Math.max(
      MIN_LINE_GAP,
      squareConstrained ? squareRequiredWidth : polygonRequiredWidth,
    );
    const widthAtLower = addUnreducedExact(
      multiplyUnreducedExact(exactWidth.slope, unreducedExactFromNumber(lower)),
      exactWidth.intercept,
    );
    if (compareExactRational(
      widthAtLower,
      unreducedExactFromNumber(requirement),
    ) < 0 && exactWidth.slope.numerator > 0n) {
      appendLocalRoot(
        roots,
        exactAffineRoot(right.exact, left.exact, requirement),
        lower,
        upper,
        diagnostics,
      );
    }
  };
  let gapLeft = pageLeft;
  for (const interval of merged) {
    appendGapRoot(
      gapLeft,
      interval.left,
      gapLeft.square || interval.left.square,
    );
    gapLeft = interval.right;
  }
  appendGapRoot(gapLeft, pageRight, gapLeft.square);
  let widestGap = gaps[0];
  for (const gap of gaps.slice(1)) {
    const comparison = compareExactRational(
      addUnreducedExact(
        multiplyUnreducedExact(gap.exactWidth.slope, unreducedExactFromNumber(lower)),
        gap.exactWidth.intercept,
      ),
      addUnreducedExact(
        multiplyUnreducedExact(
          widestGap!.exactWidth.slope,
          unreducedExactFromNumber(lower),
        ),
        widestGap!.exactWidth.intercept,
      ),
    ) || compareExactRational(gap.exactWidth.slope, widestGap!.exactWidth.slope);
    if (comparison > 0) widestGap = gap;
  }
  if (widestGap) {
    for (const gap of gaps) {
      if (gap === widestGap
        || compareExactRational(
          gap.exactWidth.slope,
          widestGap.exactWidth.slope,
        ) <= 0) continue;
      appendLocalRoot(
        roots,
        exactAffineRoot(gap.exactWidth, widestGap.exactWidth, 0),
        lower,
        upper,
        diagnostics,
      );
    }
  }
  if (roots.length === 0) return null;
  return Math.min(...roots);
}

/**
 * Hot line query over immutable, acquisition-compiled float geometry.
 *
 * Polygon vertices/intersections and rectangle edges form finite structural
 * top-Y slabs. Inside one slab, the compiled active contour pairs yield affine
 * boundaries. The local queue retains only current extrema, adjacent merged
 * boundaries, per-gap threshold roots, and roots where another
 * adjacent-boundary gap overtakes the current widest gap. A `largest` side is
 * invariant across these slabs because §20.4.3.7 selects it from the object's
 * retained maximum horizontal extent. After the earliest local root, the sweep
 * rebuilds that ordering. Every advance therefore consumes one event in a
 * finite affine arrangement. This proves termination without a global all-pairs
 * candidate set, sampling, bisection, or a pass cap.
 */
function resolvePreparedLineFloatWindowCore(
  topY: number,
  polygonRequiredWidth: number,
  probeH: number,
  paraX: number,
  maxWidth: number,
  prepared: PreparedFloatWrap,
  columnXLeftPt: number = paraX,
  columnXRightPt: number = paraX + maxWidth,
  reference: LineFloatReference = {
    xLeftPt: paraX,
    xRightPt: paraX + maxWidth,
    readingDirection: 'ltr',
  },
  squareRequiredWidth: number = polygonRequiredWidth,
  diagnostics: MutableLineFloatSweepDiagnostics | null = null,
): { topY: number; xOffset: number; maxWidth: number } {
  const paraXLeft = paraX;
  const paraXRight = paraX + maxWidth;
  const structuralEvents = finiteStructuralEvents(probeH, prepared);
  if (diagnostics) {
    diagnostics.structuralEventCount = structuralEvents.length;
    for (const { polygon } of prepared.floats) {
      if (!polygon) continue;
      diagnostics.compiledIntersectionCount += polygon.intersectionCount;
      diagnostics.compiledContourSpanCount += polygon.contourSpans.length;
      diagnostics.compileOrderComparisonCount += polygon.compileOrderComparisonCount;
      diagnostics.compilePairMembershipVisitCount += polygon.compilePairMembershipVisitCount;
    }
  }
  const evaluate = (candidateY: number) => {
    if (diagnostics) diagnostics.evaluatedYCount += 1;
    return lineWindowAtY(
      candidateY, probeH, paraXLeft, paraXRight, maxWidth, prepared,
      columnXLeftPt, columnXRightPt, reference,
      polygonRequiredWidth, squareRequiredWidth,
    );
  };
  const current = evaluate(topY);
  if (current) return current;
  let cursor = topY;
  let structuralIndex = structuralEvents.findIndex((eventY) => eventY > cursor);
  while (structuralIndex >= 0 && structuralIndex < structuralEvents.length) {
    const upper = structuralEvents[structuralIndex]!;
    const localEvent = nextLocalSweepEvent(
      cursor, upper, probeH, paraXLeft, paraXRight, prepared,
      columnXLeftPt, columnXRightPt, reference,
      polygonRequiredWidth, squareRequiredWidth, diagnostics,
    );
    if (localEvent !== null) {
      if (diagnostics) diagnostics.localRootEventCount += 1;
      const candidate = evaluate(localEvent);
      if (candidate) return candidate;
      cursor = localEvent;
      continue;
    }
    const candidate = evaluate(upper);
    if (candidate) return candidate;
    cursor = upper;
    do structuralIndex += 1;
    while (structuralIndex < structuralEvents.length
      && structuralEvents[structuralIndex]! <= cursor);
  }
  throw new Error('Finite float line-window event sweep found no usable terminal Y');
}

export function resolvePreparedLineFloatWindow(
  topY: number,
  polygonRequiredWidth: number,
  probeH: number,
  paraX: number,
  maxWidth: number,
  prepared: PreparedFloatWrap,
  columnXLeftPt: number = paraX,
  columnXRightPt: number = paraX + maxWidth,
  reference: LineFloatReference = {
    xLeftPt: paraX,
    xRightPt: paraX + maxWidth,
    readingDirection: 'ltr',
  },
  squareRequiredWidth: number = polygonRequiredWidth,
): { topY: number; xOffset: number; maxWidth: number } {
  return resolvePreparedLineFloatWindowCore(
    topY, polygonRequiredWidth, probeH, paraX, maxWidth, prepared,
    columnXLeftPt, columnXRightPt, reference, squareRequiredWidth,
  );
}

/** Internal instrumentation for deterministic solver-complexity regression tests. */
export function resolvePreparedLineFloatWindowWithDiagnostics(
  topY: number,
  polygonRequiredWidth: number,
  probeH: number,
  paraX: number,
  maxWidth: number,
  prepared: PreparedFloatWrap,
  columnXLeftPt: number = paraX,
  columnXRightPt: number = paraX + maxWidth,
  reference: LineFloatReference = {
    xLeftPt: paraX,
    xRightPt: paraX + maxWidth,
    readingDirection: 'ltr',
  },
  squareRequiredWidth: number = polygonRequiredWidth,
): Readonly<{
  window: { topY: number; xOffset: number; maxWidth: number };
  diagnostics: LineFloatSweepDiagnostics;
}> {
  const diagnostics: MutableLineFloatSweepDiagnostics = {
    compiledIntersectionCount: 0,
    compiledContourSpanCount: 0,
    compileOrderComparisonCount: 0,
    compilePairMembershipVisitCount: 0,
    structuralEventCount: 0,
    localRootCandidateCount: 0,
    localRootEventCount: 0,
    evaluatedYCount: 0,
  };
  const window = resolvePreparedLineFloatWindowCore(
    topY, polygonRequiredWidth, probeH, paraX, maxWidth, prepared,
    columnXLeftPt, columnXRightPt, reference, squareRequiredWidth, diagnostics,
  );
  return Object.freeze({ window, diagnostics: Object.freeze({ ...diagnostics }) });
}

/**
 * Multi-float collision resolution for a NEW wrap float, against floats already
 * registered on the page.
 *
 * What ECMA-376 actually mandates here is narrow, and the mandate differs by
 * what kind of object forbids overlap:
 *   - A DrawingML anchor with @allowOverlap="false" (Part 1 §20.4.2.3): an
 *     object that "cannot overlap other DrawingML object … shall be
 *     repositioned when displayed to prevent this overlap" — i.e. it must avoid
 *     OTHER DRAWINGML OBJECTS. (We never pass allowOverlap=false for shapes/
 *     images today; the default is "true", which only *permits* overlap.)
 *   - A floating table with <w:tblOverlap w:val="never"/> (§17.4.56): the table
 *     "cannot overlap with OTHER FLOATING TABLES in the document." It does NOT
 *     mandate avoiding DrawingML anchors (§20.4.2.3) or text frames — those keep
 *     their own §20.4.2.3 behavior. So a never-overlap table must only avoid
 *     blockers with kind==='table'.
 * allowOverlap="true"/omitted (the default, §20.4.2.3 / §17.4.56) only *permits*
 * overlap; the spec is silent on whether a renderer may avoid it. So:
 *   - allowOverlap === false → spec-mandated avoidance. Scoped by `kind`: a
 *     table avoids only other tables (§17.4.56); any other kind would avoid all
 *     (§20.4.2.3) — not currently exercised, see above.
 *   - allowOverlap === true  → implementation-defined avoidance of floats
 *     anchored in OTHER paragraphs only.
 *
 * EVERYTHING ELSE in this function is implementation-defined — ECMA-376 Part 1
 * does NOT specify it. This is a deterministic compatibility policy, informed
 * by observed Office output but not a normative Word or ECMA-376 requirement:
 *   - the move DIRECTION (right first, then down),
 *   - WHICH float moves (the later/document-order float is the "new" one),
 *   - the "same-paragraph floats never displace each other" gate (the paraId
 *     scoping under allowOverlap=true above),
 *   - and the move AMOUNT. Note the dist* padding reused below is, per
 *     §20.4.2.3/§20.4.2.17, the minimum distance between the float and *text*
 *     (wrapSquare geometry) — it is NOT spec-defined as a float-to-float gap.
 *     Using it to seat one float beside another is our own choice.
 *
 * If the §20.4.2.3 "shall be repositioned" requirement is ever satisfiable in
 * more than one way, the particular re-seating below remains an implementation
 * policy; keep it scoped as such until a specification-backed rule is found.
 *
 * We re-seat horizontally to the right of the blocking float(s) first (margins
 * may be used — Word lets a displaced float sit in the page margin), and only
 * fall back to a vertical push when no horizontal room remains.
 *
 * Coordinates are page-absolute in one caller-selected unit (px or pt). `(x,y)`
 * is the image box origin (no dist); every scalar and blocker must use that same
 * unit. `pageRight` is the page edge and `floats` is the active float set.
 */
export function resolveFloatOverlap(
  x: number, y: number, w: number, h: number,
  dl: number, dr: number, dt: number, db: number,
  paraId: number, allowOverlap: boolean,
  kind: FloatRect['kind'],
  pageRight: number, floats: readonly Pick<
    FloatRect,
    'kind' | 'xLeft' | 'xRight' | 'yTop' | 'yBottom' | 'paraId'
  >[],
): { x: number; y: number } {
  // Which already-registered floats are eligible blockers remains a caller
  // policy; only the axis-aligned right-then-down solver is shared.
  const blockers = floats
    .filter((f) =>
      allowOverlap ? f.paraId !== paraId : kind !== 'table' || f.kind === 'table')
    .map((f) => ({
      get left() { return f.xLeft; },
      get right() { return f.xRight; },
      get top() { return f.yTop; },
      get bottom() { return f.yBottom; },
    }));
  try {
    const resolved = resolveAxisAlignedOverlap(
      {
        left: x - dl,
        right: x + w + dr,
        top: y - dt,
        bottom: y + h + db,
      },
      blockers,
      {
        overlapEpsilon: FLOAT_OVERLAP_EPS,
        rightBoundary: pageRight,
        rightBoundarySlack: FLOAT_PAGE_RIGHT_SLACK,
      },
    );
    return { x: resolved.left + dl, y: resolved.top + dt };
  } catch (error) {
    if (
      error instanceof Error
      && error.message === 'Axis-aligned overlap resolution did not converge'
    ) {
      throw new Error('Float overlap resolution did not converge');
    }
    throw error;
  }
}

/**
 * If y is inside a topAndBottom float that horizontally overlaps the paragraph's
 * column band [paraXLeft, paraXRight], return that float's bottom; otherwise
 * return y. Mirrors `resolveLineFloatWindow` step 1: §20.4.2.20 excludes text
 * only where THIS OBJECT is placed, so a float anchored in another newspaper
 * column (§17.6.4) — the page float set is shared across columns — is filtered
 * out via the shared `floatOverlapsColumnX` predicate.
 */
export function skipPastTopAndBottom(
  y: number,
  floats: FloatRect[],
  paraXLeft: number,
  paraXRight: number,
): number {
  const consumedBottoms = new Set<number>();
  for (;;) {
    let next = y;
    for (const f of floats) {
      if (f.mode !== 'topAndBottom') continue;
      if (!floatOverlapsColumnX(f, paraXLeft, paraXRight)) continue;
      if (y >= f.yTop && y < f.yBottom) next = Math.max(next, f.yBottom);
    }
    if (next === y) return y;
    if (!Number.isFinite(next) || next < y || consumedBottoms.has(next)) {
      throw new Error('Top-and-bottom solver violated strictly increasing finite-bottom progress');
    }
    consumedBottoms.add(next);
    y = next;
  }
}

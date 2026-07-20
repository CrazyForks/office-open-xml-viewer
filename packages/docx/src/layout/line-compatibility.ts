import { defineCompatibilityRule } from './compatibility.js';

export const WORD_EAST_ASIAN_GRID_LINE_ALLOCATION = defineCompatibilityRule({
  id: 'word-east-asian-grid-line-allocation',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/compatibility.test.ts#pins East Asian grid allocation and the untabled Far East metric factor',
  },
  description: 'For an East Asian single-spaced line on a document grid, preserve the measured whole-cell allocation from the intended face design height and use the established 1.3-times-em fallback only when that design height is unavailable.',
});

export const WORD_GRID_AT_LEAST_TALL_LINE_UNSNAPPED = defineCompatibilityRule({
  id: 'word-grid-at-least-tall-line-unsnapped',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/line-box-height.test.ts#does not round tall East Asian content up to an additional grid cell',
  },
  description: 'An explicitly authored atLeast line on an active document grid keeps the maximum of its natural height, authored minimum, and one pitch instead of rounding tall content to another whole cell.',
});

export const WORD_DEGENERATE_LINE_SPACING_SINGLE = defineCompatibilityRule({
  id: 'word-degenerate-line-spacing-single',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-DOC] §2.9.146',
  },
  description: 'Preserve a non-collapsing single-line fallback for exact or automatic line spacing at or below zero, consistent with the native LSPD representation.',
});

export const WORD_AUTO_MULTIPLE_BASELINE_PIN = defineCompatibilityRule({
  id: 'word-auto-multiple-baseline-pin',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/line-spacing-baseline.test.ts#2.0× keeps the baseline at top + ascent (extra 1.0× leading below, NOT centred)',
  },
  description: 'Paint an automatic line-spacing multiplier at or above one with its glyph baseline pinned inside the single design line and place multiplier leading below it; this is draw-only and does not replace the centered trailing-mark pagination metric.',
});

export const WORD_MIXED_ANCHOR_VISIBLE_LINE_METRICS = defineCompatibilityRule({
  id: 'word-mixed-anchor-visible-line-metrics',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/anchor-host-metrics.test.ts#reserves host line height without using its zero-ink box for a visible run baseline',
  },
  description: 'A zero-ink drawing anchor host reserves its line and grid height while visible neighboring glyphs retain their own ascent, descent, and design-line baseline.',
});

export const WORD_JUSTIFICATION_LEADING_INDENT_EXCLUSION = defineCompatibilityRule({
  id: 'word-justification-leading-indent-exclusion',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/text-distribute.test.ts#forwards (segs, slack, firstContentSi, lastDrawnSi) positionally',
  },
  description: 'Keep leading whitespace used as a first-line text indent fixed while distributing justified-line slack across content in a left-to-right line.',
});

export const WORD_RUBY_PARAGRAPH_UNIFORM_LINE_ADVANCE = defineCompatibilityRule({
  id: 'word-ruby-paragraph-uniform-line-advance',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/paragraph-measure.test.ts#uses one uniform snapped advance for every line in a ruby paragraph',
  },
  description: 'Every line in a ruby-bearing paragraph uses the paragraph-wide maximum snapped line advance so its baseline rhythm remains uniform.',
});

export const WORD_FAR_EAST_SINGLE_LINE_FACTOR = 1.3;

export function wordEastAsianGridLineCells(
  naturalHeightPx: number,
  pitchPx: number,
): number {
  return pitchPx > 0 ? Math.max(1, Math.ceil(naturalHeightPx / pitchPx)) : 1;
}

export function wordFarEastSingleLinePx(
  intendedSinglePx: number,
  emPx: number,
): number {
  return intendedSinglePx > 0
    ? intendedSinglePx
    : emPx * WORD_FAR_EAST_SINGLE_LINE_FACTOR;
}

export function wordGridAtLeastLineHeightPx(
  naturalPx: number,
  authoredMinimumPx: number,
  pitchPx: number,
): number {
  return Math.max(naturalPx, authoredMinimumPx, pitchPx);
}

export function wordDegenerateLineSpacingIsSingle(
  rule: string,
  value: number,
): boolean {
  return (rule === 'exact' || rule === 'auto') && value <= 0;
}

export function wordAutoMultipleCenterBoxPx(input: Readonly<{
  autoMultiple: boolean;
  compressedAuto: boolean;
  glyphNaturalPx: number;
  intendedSinglePx: number;
  lineHeightPx: number;
}>): number {
  return input.autoMultiple && !input.compressedAuto
    ? Math.max(input.glyphNaturalPx, input.intendedSinglePx)
    : input.lineHeightPx;
}

export function wordVisibleLineMetricPx(
  reservedMetricPx: number,
  visibleMetricPx: number | undefined,
): number {
  return visibleMetricPx ?? reservedMetricPx;
}

export function wordFirstJustifiedContentSegment(
  segments: readonly unknown[],
  bidi: boolean,
  textOf: (segment: unknown) => string | undefined,
): number {
  if (bidi) return 0;
  for (let index = 0; index < segments.length; index += 1) {
    const text = textOf(segments[index]);
    if (text === undefined || /\S/.test(text)) return index;
  }
  return 0;
}

export function wordRubyUniformLineHeightPx(
  hasRuby: boolean,
  lineHeightsPx: readonly number[],
): number {
  return hasRuby ? Math.max(0, ...lineHeightsPx) : 0;
}

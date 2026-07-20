import { defineCompatibilityRule } from './compatibility.js';
import type { LayoutRect } from './types.js';

export const WORD_EXACT_ROW_HEIGHT_BOTTOM_PADDING = defineCompatibilityRule({
  id: 'word-exact-row-height-bottom-padding',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.180(d)',
  },
  description: 'Word adds the largest bottom cell margin to an exact trHeight instead of treating that margin as part of the authored height.',
});

export const WORD_NIL_TABLE_BORDER_SUPPRESSION = defineCompatibilityRule({
  id: 'word-nil-table-border-suppression',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.169',
  },
  description: 'Word treats a table border value of none as omission while nil remains authored and suppresses the complete shared edge.',
});

export const WORD_SPACED_CELL_INSIDE_BORDER_CONFLICT = defineCompatibilityRule({
  id: 'word-spaced-cell-inside-border-conflict',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.136, 2.1.138',
  },
  description: 'With non-zero cell spacing, Word retains the narrow conditional tcBorders insideH/insideV conflict against the corresponding table inside border.',
});

export const WORD_TABLE_INDENT_ALL_ALIGNMENTS = defineCompatibilityRule({
  id: 'word-table-indent-all-alignments',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.155',
  },
  description: 'Word applies tblInd as a signed leading-edge translation for every table alignment, reversing the translation for bidi visual order.',
});

export const WORD_EXACT_ROW_VERTICAL_CLIP_ONLY = defineCompatibilityRule({
  id: 'word-exact-row-vertical-clip-only',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/table.test.ts#clips an overflowing merged owner when every row in its span is exact',
  },
  description: 'Preserve the established exact-row overflow behavior that clips the owned vertical interval without clipping nested table ink horizontally to the cell box.',
});

export const WORD_OVER_PAGE_CANT_SPLIT_CLIP = defineCompatibilityRule({
  id: 'word-over-page-cant-split-clip',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.120',
  },
  description: 'Word starts an over-page cantSplit row on a fresh page and clips its overflow instead of synthesizing a row continuation.',
});

export const WORD_POSITIONED_TABLE_ADJACENCY_EXCLUSION = defineCompatibilityRule({
  id: 'word-positioned-table-adjacency-exclusion',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.149(a)',
  },
  description: 'Word excludes effectively positioned tables from the logical adjacent-table sequence before retained layout consumes the parser-owned sequence identity.',
});

export function wordExactRowFloorPt(
  authoredHeightPt: number | null,
  bottomCellMarginsPt: readonly number[],
): number {
  return Math.max(0, authoredHeightPt ?? 0)
    + Math.max(0, ...bottomCellMarginsPt);
}

export function wordAuthoredBorderParticipates(
  authoredStyle: string | null | undefined,
): boolean {
  return authoredStyle !== null
    && authoredStyle !== undefined
    && authoredStyle !== 'none';
}

export function wordAlignedTableOriginPt(
  alignedPt: number,
  indentPt: number,
  bidiVisual: boolean,
): number {
  return bidiVisual ? alignedPt - indentPt : alignedPt + indentPt;
}

export function wordSpacedCellUsesSeparatedBorderGrid(spacingPt: number): boolean {
  return spacingPt > 0;
}

export function wordExactRowVerticalClipBounds(
  cellFlowBounds: LayoutRect,
  containingFlowBounds: LayoutRect,
): LayoutRect {
  return Object.freeze({
    xPt: containingFlowBounds.xPt,
    yPt: cellFlowBounds.yPt,
    widthPt: containingFlowBounds.widthPt,
    heightPt: cellFlowBounds.heightPt,
  });
}

export function wordClipsOverPageCantSplitRow(input: Readonly<{
  compatibility: 'word' | 'standard';
  availableHeightPt: number;
  freshPageHeightPt: number;
  epsilonPt: number;
}>): boolean {
  return input.compatibility === 'word'
    && input.availableHeightPt + input.epsilonPt >= input.freshPageHeightPt;
}

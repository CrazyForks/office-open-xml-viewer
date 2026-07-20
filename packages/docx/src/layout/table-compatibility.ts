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

export const WORD_TABLE_BORDER_WEIGHT_PRECEDENCE = defineCompatibilityRule({
  id: 'word-table-border-weight-precedence',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.169',
  },
  description: 'Use the documented Word border numbers for shared-cell conflict weight and force dotted and dashed borders to a complete weight of one.',
});

export const WORD_OMITTED_ROW_HEIGHT_RULE_AT_LEAST = defineCompatibilityRule({
  id: 'word-omitted-row-height-rule-at-least',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.180',
  },
  description: 'Treat an omitted trHeight hRule as atLeast while retaining an explicitly authored auto rule as authored input.',
});

export const WORD_AUTHORED_AUTO_ROW_HEIGHT_FLOOR = defineCompatibilityRule({
  id: 'word-authored-auto-row-height-floor',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/table-row-height.test.ts#auto with @val — @val is honored as a lower bound (Word-compatible)',
  },
  description: 'Preserve the established legacy-model behavior that an auto row with an authored finite height value uses that value as a lower bound.',
});

export const WORD_EFFECTIVE_FLOATING_TABLE_POSITIONING = defineCompatibilityRule({
  id: 'word-effective-floating-table-positioning',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.162',
  },
  description: 'Use parser-retained effective positioning status rather than lexical tblpPr presence to decide whether a table leaves ordinary flow.',
});

export const WORD_TABLE_CELL_SPACING_SCOPE_SHADOW = defineCompatibilityRule({
  id: 'word-table-cell-spacing-scope-shadow',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.152, 2.1.153, 2.1.154',
  },
  description: 'At each table-cell-spacing precedence scope, pct, auto, and nil resolve to zero and shadow lower scopes instead of being treated as absent.',
});

export const WORD_TABLE_MARGIN_SCOPE_SHADOW = defineCompatibilityRule({
  id: 'word-table-margin-scope-shadow',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.116, 2.1.125, 2.1.146, 2.1.177',
  },
  description: 'Preserve the documented scope-specific treatment of non-dxa table cell margins: leading/trailing defaults may resolve to zero while cell/exception and nil top/bottom values remain ignored.',
});

export const WORD_FIRST_ROW_TABLE_EXCEPTION_SCOPE = defineCompatibilityRule({
  id: 'word-first-row-table-exception-scope',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.156, 2.1.158, 2.1.167',
  },
  description: 'Apply the supported first-row table-property exception facts at table scope, including authored preferred-width shadowing.',
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

export function wordSpacedCellInsideBorderOverridesTable(input: Readonly<{
  spacingPt: number;
  directStyle: string | null | undefined;
  conditionalInsideStyle: string | null | undefined;
}>): boolean {
  return input.spacingPt > 0
    && !wordAuthoredBorderParticipates(input.directStyle)
    && wordAuthoredBorderParticipates(input.conditionalInsideStyle);
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

const WORD_BORDER_NUMBER: Readonly<Record<string, number>> = Object.freeze({
  single: 1,
  thick: 2,
  double: 3,
  dotDash: 8,
  dotDotDash: 9,
  triple: 10,
  thinThickSmallGap: 11,
  thickThinSmallGap: 12,
  thinThickThinSmallGap: 13,
  thinThickMediumGap: 14,
  thickThinMediumGap: 15,
  thinThickThinMediumGap: 16,
  thinThickLargeGap: 17,
  thickThinLargeGap: 18,
  thinThickThinLargeGap: 19,
  wave: 20,
  doubleWave: 21,
  dashSmallGap: 22,
  dashDotStroked: 23,
  threeDEmboss: 24,
  threeDEngrave: 25,
  outset: 26,
  inset: 27,
});

export function wordTableBorderWeight(
  style: string,
  widthPt: number,
): number {
  if (style === 'dotted' || style === 'dashed') return 1;
  return Math.max(0, widthPt) * 8 * (WORD_BORDER_NUMBER[style] ?? 0);
}

export function wordNilBorderSuppressesSharedEdge(
  firstStyle: string | null | undefined,
  secondStyle: string | null | undefined,
): boolean {
  return firstStyle === 'nil' || secondStyle === 'nil';
}

export function wordTableRowHeightRule(
  normalizedRule: 'exact' | 'atLeast' | 'auto',
  authored: boolean,
): 'exact' | 'atLeast' | 'auto' {
  return authored ? normalizedRule : 'atLeast';
}

export function wordAuthoredAutoRowHeightUsesFloor(
  rule: string | null | undefined,
  authoredHeight: number | null | undefined,
): boolean {
  return rule === 'auto'
    && authoredHeight !== null
    && authoredHeight !== undefined
    && Number.isFinite(authoredHeight);
}

export function wordTableCellSpacingValuePt(
  kind: string,
  dxaValuePt: number | null,
): number | null {
  if (kind === 'pct' || kind === 'auto' || kind === 'nil') return 0;
  return dxaValuePt;
}

export function wordTableMarginValuePt(input: Readonly<{
  kind: string;
  dxaValuePt: number | null;
  scope: 'cell' | 'exception' | 'table' | 'style';
  edge: 'top' | 'bottom' | 'start' | 'end';
}>): number | null {
  if (input.kind === 'dxa') return input.dxaValuePt;
  if (input.scope === 'cell' || input.scope === 'exception') return null;
  if (input.edge === 'start' || input.edge === 'end') {
    if (input.kind === 'pct' || input.kind === 'auto' || input.kind === 'nil') return 0;
  }
  return null;
}

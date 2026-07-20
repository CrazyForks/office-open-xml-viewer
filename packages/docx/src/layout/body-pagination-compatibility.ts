import type {
  BodyLayoutInput,
  BodyParagraphSourceInput,
} from './body-layout-input.js';
import { defineCompatibilityRule } from './compatibility.js';
import type { ParagraphLayout } from './types.js';

export const WORD_TERMINAL_COLUMN_BREAK = defineCompatibilityRule({
  id: 'word-terminal-column-break',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/pagination.test.ts#ignores a terminal last-column break before a hard page boundary',
  },
  description: 'The established pagination regression contract does not materialize a column transition when no body flow content occurs before the next forced page or non-continuous section boundary.',
});

export const WORD_PRE_BREAK_ANCHOR_PARAGRAPH = defineCompatibilityRule({
  id: 'word-pre-break-anchor-paragraph',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/pagination.test.ts#does not push an anchor-only pre-break paragraph to a new page just for its empty mark',
  },
  description: 'The established pagination regression contract keeps an anchor-only paragraph immediately before an authored page break in the pre-break flow region without charging its otherwise visible paragraph mark.',
});

export const WORD_PRE_BREAK_INLINE_DRAWING_GROUP = defineCompatibilityRule({
  id: 'word-pre-break-inline-drawing-group',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/pagination.test.ts#moves a preceding image with its pre-break callout when the pair only fits fresh',
  },
  description: 'The established pagination regression contract relocates a preceding inline DrawingML resource with an immediately following host-owned anchor paragraph before an authored page break when the pair fits only in a fresh flow region.',
});

export const WORD_CONTINUOUS_SECTION_MARK_SPACING = defineCompatibilityRule({
  id: 'word-continuous-section-mark-spacing',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/body-layout-input.test.ts#projects mutually exclusive collapsed-mark and drop-previous-after roles',
  },
  description: 'The retained body input projects the established continuous-section empty-mark spacing behavior into one mutually exclusive role before pagination.',
});

export const WORD_CONTEXTUAL_SPACING_PER_SIDE = defineCompatibilityRule({
  id: 'word-contextual-spacing-per-side',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/contextual-spacing-body-paint.test.ts#paints the adjudicated six-case gap table',
  },
  description: 'For same-style adjacent paragraphs, contextualSpacing removes only the contribution owned by each toggling side; a current-only toggle preserves the previous paragraph spaceAfter contribution.',
});

/** Compatibility projection governed by {@link WORD_CONTINUOUS_SECTION_MARK_SPACING}. */
export function wordContinuousSectionRole(
  sequence: BodyLayoutInput['sequence'],
  index: number,
): BodyParagraphSourceInput['continuousSectionRole'] {
  const entry = sequence[index];
  if (entry?.kind !== 'body-block' || entry.block.kind !== 'paragraph') return undefined;
  const next = sequence[index + 1];
  const afterNext = sequence[index + 2];
  const suppressBefore = entry.block.inkless === true
    && next?.kind === 'begin-section'
    && next.section.startType === 'continuous';
  if (suppressBefore && entry.block.spaceBeforePt === 0) return 'collapse-mark';
  if (suppressBefore) return 'suppress-before';
  return entry.block.inkless === true
    && next?.kind === 'body-block'
    && next.block.kind === 'paragraph'
    && next.block.inkless === true
    && next.block.spaceBeforePt === 0
    && afterNext?.kind === 'begin-section'
    && afterNext.section.startType === 'continuous'
    ? 'drop-previous-after'
    : undefined;
}

function paragraphHasOnlyDrawingAnchors(layout: ParagraphLayout): boolean {
  return layout.drawings.length > 0 && layout.lines.every((line) =>
    line.placements.every((placement) =>
      placement.kind === 'drawing' || placement.kind === 'anchor-host'));
}

/** A positive-size inline DrawingML resource that participates in line flow. */
export function wordPreBreakInlineDrawingResource(layout: ParagraphLayout): boolean {
  return layout.lines.some((line) => line.placements.some((placement) =>
    placement.kind === 'resource'
    && (placement.resourceKind === 'image' || placement.resourceKind === 'chart')
    && placement.bounds !== undefined
    && placement.bounds.widthPt > 0
    && placement.bounds.heightPt > 0));
}

/** Host-relative drawing extent below the following paragraph's flow cursor. */
export function wordPreBreakHostAnchorExtentPt(
  layout: ParagraphLayout,
  anchorTopPt: number,
): number | null {
  if (!paragraphHasOnlyDrawingAnchors(layout)) return null;
  const drawings = layout.drawings.filter((drawing) => (
    drawing.anchorLayer?.verticalOwnership === 'host'
    && Number.isFinite(drawing.flowBounds.xPt)
    && Number.isFinite(drawing.flowBounds.yPt)
    && Number.isFinite(drawing.flowBounds.widthPt)
    && Number.isFinite(drawing.flowBounds.heightPt)
    && drawing.flowBounds.widthPt > 0
    && drawing.flowBounds.heightPt > 0
  ));
  if (drawings.length !== layout.drawings.length) return null;
  const bottomPt = Math.max(...drawings.map((drawing) =>
    drawing.flowBounds.yPt + drawing.flowBounds.heightPt));
  return Math.max(0, bottomPt - anchorTopPt);
}

/** Returns the retained anchor paragraph with its compatibility-only flow charge removed. */
export function wordFlowNeutralPreBreakAnchorParagraph(
  layout: ParagraphLayout,
): ParagraphLayout | null {
  if (!paragraphHasOnlyDrawingAnchors(layout)) return null;
  const { paragraphMark: _paragraphMark, ...withoutMark } = layout;
  return Object.freeze({
    ...withoutMark,
    advancePt: 0,
    flowBounds: Object.freeze({ ...layout.flowBounds, heightPt: 0 }),
  });
}

/**
 * Plans which authored column breaks have body flow content before the next
 * forced boundary. ECMA-376 §§17.3.3.1 and 17.18.4 define the column advance;
 * suppressing a terminal transition is Word compatibility behavior recorded by
 * {@link WORD_TERMINAL_COLUMN_BREAK}.
 */
export function wordActiveColumnBreakIndexes(
  sequence: BodyLayoutInput['sequence'],
): ReadonlySet<number> {
  const active = new Set<number>();
  let hasFlowContentBeforeBoundary = false;
  for (let index = sequence.length - 1; index >= 0; index -= 1) {
    const entry = sequence[index]!;
    if (entry.kind === 'body-block') {
      hasFlowContentBeforeBoundary = entry.block.kind !== 'paragraph'
        || !entry.block.pageBreakBefore;
      continue;
    }
    if (entry.kind === 'adjacent-table-group') {
      hasFlowContentBeforeBoundary = true;
      continue;
    }
    if (entry.kind === 'authored-break') {
      if (entry.break === 'column') {
        if (hasFlowContentBeforeBoundary) active.add(index);
      } else {
        hasFlowContentBeforeBoundary = false;
      }
      continue;
    }
    if (entry.kind === 'begin-section' && entry.section.startType !== 'continuous') {
      hasFlowContentBeforeBoundary = false;
    }
  }
  return active;
}

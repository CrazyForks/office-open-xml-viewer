import { describe, expect, it } from 'vitest';
import {
  defineCompatibilityRule,
  LINE_START_GAP_EPS_PT,
  WORD_FLOAT_DIFFERENT_PARAGRAPH_DISPLACEMENT,
  WORD_PAGE_ANCHORED_TABLE_COLLISION_DEFERRAL,
  WORD_MIN_LINE_START_PT,
  WORD_SQUARE_LINE_START_ONE_INCH,
  WORD_SECTION_BTLR_TBRL_PAGE_FRAME,
  wordMinLineStartPx,
} from './compatibility.js';
import {
  WORD_PRE_BREAK_ANCHOR_PARAGRAPH,
  WORD_PRE_BREAK_INLINE_DRAWING_GROUP,
  WORD_CONTINUOUS_SECTION_MARK_SPACING,
  WORD_TERMINAL_COLUMN_BREAK,
} from './body-pagination-compatibility.js';
import { ESTABLISHED_NEXT_COLUMN_PAGE_ADVANCE } from './page-flow-compatibility.js';
import {
  WORD_ZERO_RELATIVE_SIZE_EXTENT_FALLBACK,
  wordZeroRelativeSizeUsesExtent,
} from './anchor-compatibility.js';
import {
  WORD_BOOK_FOLD_GUTTER_RIGHT_EDGE,
  WORD_COLUMN_SEPARATOR_SECTION_BAND,
  WORD_CONTINUOUS_SECTION_PAGE_NUMBER_RESTART,
  WORD_DEFAULT_LINE_NUMBER_DISTANCE,
  WORD_DEFAULT_LINE_NUMBER_DISTANCE_PT,
  WORD_TRAILING_EMPTY_MARK_BASELINE_ADMISSION,
  wordBookFoldGutterEdge,
  wordColumnSeparatorBlockBand,
  wordContinuousSectionRestartDisplayNumber,
  wordLineNumberDistancePt,
  wordTrailingEmptyMarkAdmissionAllowancePt,
} from './section-compatibility.js';
import {
  WORD_EXACT_ROW_HEIGHT_BOTTOM_PADDING,
  WORD_EXACT_ROW_VERTICAL_CLIP_ONLY,
  WORD_NIL_TABLE_BORDER_SUPPRESSION,
  WORD_OVER_PAGE_CANT_SPLIT_CLIP,
  WORD_POSITIONED_TABLE_ADJACENCY_EXCLUSION,
  WORD_SPACED_CELL_INSIDE_BORDER_CONFLICT,
  WORD_TABLE_INDENT_ALL_ALIGNMENTS,
  wordAlignedTableOriginPt,
  wordAuthoredBorderParticipates,
  wordClipsOverPageCantSplitRow,
  wordExactRowFloorPt,
  wordExactRowVerticalClipBounds,
  wordSpacedCellUsesSeparatedBorderGrid,
} from './table-compatibility.js';

describe('defineCompatibilityRule', () => {
  it('retains explicit Microsoft evidence as immutable data', () => {
    const rule = defineCompatibilityRule({
      id: 'word-example',
      evidence: { kind: 'microsoft-note', reference: '[MS-OE376] §2.1' },
      description: 'Synthetic compatibility boundary',
    });

    expect(rule.evidence).toEqual({ kind: 'microsoft-note', reference: '[MS-OE376] §2.1' });
    expect(Object.isFrozen(rule)).toBe(true);
    expect(Object.isFrozen(rule.evidence)).toBe(true);
  });

  it('retains a concrete regression-test reference without claiming an Office observation', () => {
    const rule = defineCompatibilityRule({
      id: 'regression-example',
      evidence: {
        kind: 'regression-test',
        reference: 'packages/docx/src/example.test.ts#named behavior',
      },
      description: 'Synthetic compatibility boundary',
    });

    expect(rule.evidence).toEqual({
      kind: 'regression-test',
      reference: 'packages/docx/src/example.test.ts#named behavior',
    });
  });

  it('rejects observation evidence without a reproducible fixture identity', () => {
    expect(() => defineCompatibilityRule({
      id: 'word-observation',
      evidence: {
        kind: 'office-observation',
        syntheticFixtureId: '',
        application: 'Word',
        version: 'current',
        platform: 'Windows',
      },
      description: 'Unreproducible observation',
    })).toThrow(/syntheticFixtureId/);
  });

  it('rejects unstructured evidence and unstable rule identities', () => {
    expect(() => defineCompatibilityRule({
      id: 'Not Stable',
      evidence: { kind: 'microsoft-note', reference: '[MS-OE376] §2.1' },
      description: 'Invalid identity',
    })).toThrow(/kebab-case/);
    expect(() => defineCompatibilityRule({
      id: 'bad-microsoft-reference',
      evidence: { kind: 'microsoft-note', reference: 'MS-OE376 somewhere' },
      description: 'Invalid Microsoft evidence',
    })).toThrow(/Microsoft specification section/);
    expect(() => defineCompatibilityRule({
      id: 'bad-regression-reference',
      evidence: { kind: 'regression-test', reference: 'a test somewhere' },
      description: 'Invalid regression evidence',
    })).toThrow(/path#test-title/);
  });

  it('records the approved Word section btLr page-frame difference', () => {
    expect(WORD_SECTION_BTLR_TBRL_PAGE_FRAME).toEqual({
      id: 'word-section-btlr-tbrl-page-frame',
      evidence: {
        kind: 'regression-test',
        reference: 'packages/docx/src/layout/coordinate-space.test.ts#maps Transitional text direction %s to %s',
      },
      description: expect.stringMatching(
        /Issue #988 comment 4950296007.*normative.*lr.*page frame.*glyph orientation.*paint-owned/i,
      ),
    });
    expect(Object.isFrozen(WORD_SECTION_BTLR_TBRL_PAGE_FRAME)).toBe(true);
    expect(Object.isFrozen(WORD_SECTION_BTLR_TBRL_PAGE_FRAME.evidence)).toBe(true);
  });

  it('pins every regression-test reference to a concrete repository test', () => {
    const rules = [
      WORD_SECTION_BTLR_TBRL_PAGE_FRAME,
      WORD_TERMINAL_COLUMN_BREAK,
      WORD_PRE_BREAK_ANCHOR_PARAGRAPH,
      WORD_PRE_BREAK_INLINE_DRAWING_GROUP,
      WORD_CONTINUOUS_SECTION_MARK_SPACING,
      ESTABLISHED_NEXT_COLUMN_PAGE_ADVANCE,
    ];

    expect(rules.map((rule) => (
      rule.evidence.kind === 'regression-test' ? rule.evidence.reference : null
    ))).toEqual([
      'packages/docx/src/layout/coordinate-space.test.ts#maps Transitional text direction %s to %s',
      'packages/docx/src/pagination.test.ts#ignores a terminal last-column break before a hard page boundary',
      'packages/docx/src/pagination.test.ts#does not push an anchor-only pre-break paragraph to a new page just for its empty mark',
      'packages/docx/src/pagination.test.ts#moves a preceding image with its pre-break callout when the pair only fits fresh',
      'packages/docx/src/body-layout-input.test.ts#projects mutually exclusive collapsed-mark and drop-previous-after roles',
      'packages/docx/src/layout/paginator.test.ts#advances nextColumn to the next page when the outgoing column has no same-page successor',
    ]);
  });
});

describe('float compatibility evidence', () => {
  it('keeps the measured square line-start threshold behind one named rule', () => {
    expect(WORD_SQUARE_LINE_START_ONE_INCH.evidence).toMatchObject({
      kind: 'regression-test',
    });
    expect(WORD_MIN_LINE_START_PT).toBe(72);
    expect(LINE_START_GAP_EPS_PT).toBe(0.05);
    expect(wordMinLineStartPx(1)).toBeCloseTo(71.95, 10);
  });

  it('names the established different-paragraph displacement policy', () => {
    expect(WORD_FLOAT_DIFFERENT_PARAGRAPH_DISPLACEMENT).toMatchObject({
      id: 'word-float-different-paragraph-displacement',
      evidence: { kind: 'regression-test' },
    });
  });

  it('names absolute floating-table collision deferral as compatibility behavior', () => {
    expect(WORD_PAGE_ANCHORED_TABLE_COLLISION_DEFERRAL).toMatchObject({
      id: 'word-page-anchored-table-collision-deferral',
      evidence: { kind: 'regression-test' },
    });
  });
});

describe('layout compatibility inventory', () => {
  it("uses Word's observed 18pt line-number distance only when omitted", () => {
    expect(WORD_DEFAULT_LINE_NUMBER_DISTANCE_PT).toBe(18);
    expect(wordLineNumberDistancePt(undefined)).toBe(18);
    expect(wordLineNumberDistancePt(null)).toBe(18);
    expect(wordLineNumberDistancePt(0)).toBe(0);
    expect(wordLineNumberDistancePt(12)).toBe(12);
  });

  it('keeps each layout observation behind explicit immutable evidence', () => {
    const rules = [
      WORD_ZERO_RELATIVE_SIZE_EXTENT_FALLBACK,
      WORD_DEFAULT_LINE_NUMBER_DISTANCE,
      WORD_CONTINUOUS_SECTION_PAGE_NUMBER_RESTART,
      WORD_TRAILING_EMPTY_MARK_BASELINE_ADMISSION,
      WORD_COLUMN_SEPARATOR_SECTION_BAND,
      WORD_BOOK_FOLD_GUTTER_RIGHT_EDGE,
      WORD_EXACT_ROW_HEIGHT_BOTTOM_PADDING,
      WORD_NIL_TABLE_BORDER_SUPPRESSION,
      WORD_SPACED_CELL_INSIDE_BORDER_CONFLICT,
      WORD_TABLE_INDENT_ALL_ALIGNMENTS,
      WORD_EXACT_ROW_VERTICAL_CLIP_ONLY,
      WORD_OVER_PAGE_CANT_SPLIT_CLIP,
      WORD_POSITIONED_TABLE_ADJACENCY_EXCLUSION,
    ];

    expect(new Set(rules.map((rule) => rule.id)).size).toBe(rules.length);
    expect(rules.every(Object.isFrozen)).toBe(true);
    expect(rules.every((rule) => Object.isFrozen(rule.evidence))).toBe(true);
  });

  it('pins behavior-preserving compatibility decision helpers', () => {
    expect(wordZeroRelativeSizeUsesExtent(0)).toBe(true);
    expect(wordZeroRelativeSizeUsesExtent(-0.1)).toBe(false);
    expect(wordExactRowFloorPt(12, [2, 4, 3])).toBe(16);
    expect(wordAuthoredBorderParticipates('none')).toBe(false);
    expect(wordAuthoredBorderParticipates('nil')).toBe(true);
    expect(wordAlignedTableOriginPt(100, 8, false)).toBe(108);
    expect(wordAlignedTableOriginPt(100, 8, true)).toBe(92);
    expect(wordSpacedCellUsesSeparatedBorderGrid(0)).toBe(false);
    expect(wordSpacedCellUsesSeparatedBorderGrid(0.1)).toBe(true);
    expect(wordExactRowVerticalClipBounds(
      { xPt: 20, yPt: 30, widthPt: 40, heightPt: 50 },
      { xPt: 5, yPt: 10, widthPt: 100, heightPt: 200 },
    )).toEqual({ xPt: 5, yPt: 30, widthPt: 100, heightPt: 50 });
    expect(wordClipsOverPageCantSplitRow({
      compatibility: 'word',
      availableHeightPt: 99.99995,
      freshPageHeightPt: 100,
      epsilonPt: 0.0001,
    })).toBe(true);
    expect(wordClipsOverPageCantSplitRow({
      compatibility: 'standard',
      availableHeightPt: 100,
      freshPageHeightPt: 100,
      epsilonPt: 0.0001,
    })).toBe(false);
    expect(wordContinuousSectionRestartDisplayNumber(4, 8, 6)).toBe(6);
    expect(wordColumnSeparatorBlockBand(10, 40)).toEqual({
      blockStartPt: 10,
      blockEndPt: 40,
    });
    expect(wordBookFoldGutterEdge()).toBe('right');
    const trailingMark = {
      hasContinuationBoundary: false,
      inkless: true,
      undecorated: true,
      keepNext: false,
      markReservePt: 0,
      pageBottomIsUnreserved: true,
      physicalRegionBottomIsActive: true,
      hasFollowingInk: true,
      markBelowBaselinePt: 3,
    };
    expect(wordTrailingEmptyMarkAdmissionAllowancePt(trailingMark)).toBe(3);
    expect(wordTrailingEmptyMarkAdmissionAllowancePt({
      ...trailingMark,
      hasFollowingInk: false,
    })).toBe(0);
  });
});

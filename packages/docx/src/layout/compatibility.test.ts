import { describe, expect, it } from 'vitest';
import {
  defineCompatibilityRule,
  LINE_START_GAP_EPS_PT,
  WORD_FLOAT_DIFFERENT_PARAGRAPH_DISPLACEMENT,
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
import { ESTABLISHED_NEXT_COLUMN_PAGE_ADVANCE } from './paginator.js';

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
});

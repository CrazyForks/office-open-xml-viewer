import { describe, expect, it } from 'vitest';
import {
  resolveDocumentLayoutSettings,
  resolveSectionLayoutContext,
} from '../layout-context.js';
import { bodyAcquisitionInputProjections } from '../parser-model.js';
import type { DocParagraph, DocxDocumentModel, SectionProps } from '../types.js';
import type { BodyAcquisitionState, RetainedTableRecord } from './acquisition-context.js';
import {
  BODY_STORY_CONTEXT,
  retainedTableRecord,
  resolveBodyParagraphLayoutContext,
  resolveStateParagraphLayoutContext,
  withTableCellStory,
} from './acquisition-state.js';

const section = (): SectionProps => ({
  pageWidth: 200,
  pageHeight: 300,
  marginTop: 20,
  marginRight: 20,
  marginBottom: 20,
  marginLeft: 20,
  headerDistance: 10,
  footerDistance: 10,
  docGridType: 'lines',
  docGridLinePitch: 12,
} as SectionProps);

const paragraph = (): DocParagraph => ({
  alignment: 'left',
  indentLeft: 0,
  indentRight: 0,
  indentFirst: 0,
  spaceBefore: 0,
  spaceAfter: 0,
  lineSpacing: null,
  numbering: null,
  tabStops: [],
  runs: [],
} as DocParagraph);

function paragraphState() {
  const document = {
    section: section(),
    body: [],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
  const layoutSettings = resolveDocumentLayoutSettings(document);
  return {
    layoutSettings,
    sectionLayout: resolveSectionLayoutContext(layoutSettings, document.section),
    acquisitionInputs: bodyAcquisitionInputProjections,
  };
}

describe('layout acquisition state', () => {
  it('owns one immutable body-story root and enters table cells without mutating it', () => {
    expect(Object.isFrozen(BODY_STORY_CONTEXT)).toBe(true);
    expect(Object.isFrozen(BODY_STORY_CONTEXT.containers)).toBe(true);

    const state = {
      ...paragraphState(),
      storyContext: BODY_STORY_CONTEXT,
    } as unknown as BodyAcquisitionState;
    const nested = withTableCellStory(state);

    expect(nested).not.toBe(state);
    expect(nested.storyContext).toEqual({
      story: 'body',
      containers: [{ kind: 'tableCell' }],
      lineNumberingEligible: false,
    });
    expect(BODY_STORY_CONTEXT.containers).toEqual([]);
  });

  it('uses body grid policy at the root and table-cell policy for state stories', () => {
    const state = paragraphState();
    expect(resolveBodyParagraphLayoutContext(state, paragraph()).lineGrid.active).toBe(true);
    expect(resolveStateParagraphLayoutContext({
      ...state,
      storyContext: {
        story: 'body',
        containers: [{ kind: 'tableCell' }],
        lineNumberingEligible: false,
      },
    }, paragraph()).lineGrid.active).toBe(false);
  });

  it('returns only retained table records and fails closed when acquisition is absent', () => {
    const record = { sourceIndex: 4 } as RetainedTableRecord;
    const state = {
      retainedTablesBySourceIndex: new Map([[4, record]]),
    } as unknown as BodyAcquisitionState;

    expect(retainedTableRecord(state, 4)).toBe(record);
    expect(() => retainedTableRecord(state, 5)).toThrow(
      'Table placement requires retained table acquisition',
    );
  });
});

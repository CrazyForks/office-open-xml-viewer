import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import type { PageBorders } from '../types.js';
import { createPageFlowSectionContext } from './context.js';
import { finalizeLayoutPage } from './page-factory.js';
import { createPageFlowState, advanceColumnOrPage } from './paginator.js';
import {
  addPageFootnoteReserve,
  commitPageFlowTransition,
  createBodyPaginationState,
  createCanonicalPageDraft,
  setBodyBalanceTarget,
} from './body-pagination.js';

const section: SectionLayoutContext = {
  geometry: {
    pageWidth: 200, pageHeight: 100,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 5, footerDistance: 5,
  },
  columns: [{ xPt: 10, wPt: 180 }],
  columnSeparator: false,
  grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
  textDirection: 'lrTb', verticalAlignment: 'top',
};

const draft = (pageIndex: number) => createCanonicalPageDraft({
  kind: 'content',
  pageIndex,
  physicalPage: {
    widthPt: 200, heightPt: 100, contentTopPt: 10, contentBottomPt: 90,
  },
  sectionOccurrenceId: 'section:0',
  section,
  region: {
    id: `region:${pageIndex}`,
    sectionOccurrenceId: 'section:0',
    section,
    writingMode: 'horizontal-tb',
    blockStartPt: 10,
    blockEndPt: 90,
    columns: [{ inlineStartPt: 10, inlineExtentPt: 180 }],
  },
});

describe('immutable canonical page transitions', () => {
  it.each(['allPages', 'firstPage', 'notFirstPage'] as const)(
    'retains parser-owned %s page borders through region draft finalization',
    (display) => {
      const pageBorders: PageBorders = {
        offsetFrom: 'page', display, zOrder: 'front',
        top: { style: 'single', width: 1, space: 12 },
      };
      const pageDraft = createCanonicalPageDraft({
        kind: 'content',
        pageIndex: 0,
        physicalPage: {
          widthPt: 200, heightPt: 100, contentTopPt: 10, contentBottomPt: 90,
        },
        sectionOccurrenceId: 'section:0',
        section,
        region: {
          id: 'region:0', sectionOccurrenceId: 'section:0', section,
          pageBorders,
          writingMode: 'horizontal-tb', blockStartPt: 10, blockEndPt: 90,
          columns: [{ inlineStartPt: 10, inlineExtentPt: 180 }],
        },
      });

      const page = finalizeLayoutPage(pageDraft.accumulator, {
        displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0',
      });

      expect(page.pageBorders).toEqual(pageBorders);
    },
  );

  it('opens a new draft without mutating the prior state or page', () => {
    const flowSection = createPageFlowSectionContext({
      sectionOccurrenceId: 'section:0',
      geometry: section.geometry,
      columns: section.columns,
      textDirection: section.textDirection,
    });
    const originalPage = draft(0);
    const original = setBodyBalanceTarget(
      addPageFootnoteReserve(
        createBodyPaginationState(createPageFlowState(flowSection), originalPage),
        12,
      ),
      48,
    );
    const transition = advanceColumnOrPage(original.flow, 'overflow');

    const next = commitPageFlowTransition(original, transition, {
      openContentPage: (event) => ({ page: draft(event.pageIndex), flow: transition.state }),
      openParityBlankPage: () => { throw new Error('unused'); },
      openContinuousSectionRegion: () => { throw new Error('unused'); },
    });

    expect(original.pages).toEqual([originalPage]);
    expect(original.footnoteReservePt).toBe(12);
    expect(original.balanceTargetPt).toBe(48);
    expect(next.pages.map((page) => page.accumulator.pageIndex)).toEqual([0, 1]);
    expect(next.flow.pageIndex).toBe(1);
    expect(next.footnoteReservePt).toBe(0);
    expect(next.balanceTargetPt).toBeNull();
    expect(Object.isFrozen(next)).toBe(true);
    expect(Object.isFrozen(next.pages)).toBe(true);
    expect(next).not.toBe(original);
  });
});

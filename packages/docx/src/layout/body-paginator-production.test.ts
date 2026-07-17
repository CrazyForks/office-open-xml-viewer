import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import type { BodyLayoutInput } from './body-layout-input.js';
import type { BodyLayoutKernel } from './body-layout-kernel.js';
import { attachBodyLayoutKernel } from './runtime-state.js';
import { paginateBody } from './body-paginator.js';
import type { LayoutServices, ParagraphLayout, SourceRef, TableLayout } from './types.js';

const emptyFlowRegistrySnapshot = () => ({
  floats: {
    coordinateSpace: 'logical-page-points' as const,
    flowDomainId: 'body',
    entries: [],
    nextParagraphId: 0,
  },
  drawingCollisions: {
    coordinateSpace: 'logical-page-points' as const,
    flowDomainId: 'body',
    entries: [],
  },
});

const source = (index: number): SourceRef => ({
  story: 'body', storyInstance: 'body', path: [index],
});

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

const paragraph = (id: string, src: SourceRef, heightPt: number): ParagraphLayout => ({
  kind: 'paragraph', id, source: src, flowDomainId: 'acquisition',
  flowBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt },
  inkBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt },
  advancePt: heightPt, ordinaryFlow: true,
  spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
  lines: [], borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
});

const table = (id: string, src: SourceRef, heightPt: number): TableLayout => ({
  kind: 'table', id, source: src, flowDomainId: 'acquisition',
  flowBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt },
  inkBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt },
  advancePt: heightPt, ordinaryFlow: true, columnWidthsPt: [180], rows: [], borders: [],
});

const tableWithFootnote = (
  id: string,
  src: SourceRef,
  heightPt: number,
  noteId: string,
): TableLayout => {
  const bounds = { xPt: 0, yPt: 0, widthPt: 180, heightPt };
  const noteParagraph = {
    ...paragraph(`${id}:paragraph`, src, heightPt),
    lines: [{
      range: { start: 0, end: 0 }, bounds, baselinePt: heightPt, advancePt: heightPt,
      placements: [{
        kind: 'text', range: { start: 0, end: 0 }, origin: { xPt: 0, yPt: heightPt },
        bounds: { ...bounds, widthPt: 0 }, advancePt: 0,
        noteReference: { kind: 'footnote', id: noteId },
      }],
    }],
  } as unknown as ParagraphLayout;
  return {
    ...table(id, src, heightPt),
    rows: [{
      kind: 'table-row', id: `${id}:row`, source: src, flowDomainId: 'acquisition',
      flowBounds: bounds, inkBounds: bounds, advancePt: heightPt, ordinaryFlow: true,
      heightPt, contentHeightPt: heightPt,
      cells: [{
        kind: 'table-cell', id: `${id}:cell`, source: src, flowDomainId: 'acquisition',
        flowBounds: bounds, inkBounds: bounds, advancePt: heightPt, ordinaryFlow: true,
        contentBounds: bounds, verticalMerge: 'none', vAlign: 'top',
        blocks: [{ layout: noteParagraph, offsetPt: 0, advancePt: heightPt }],
      }],
    }],
  };
};

const bodyOwner = () => ({
  sectionOccurrenceId: 'section:0', source: source(2), startType: 'nextPage' as const,
  context: section,
  pageNumbering: { start: null, format: null }, titlePage: false, evenAndOddHeaders: false,
  headers: { default: null, first: null, even: null },
  footers: { default: null, first: null, even: null },
  pageBordersAuthored: false, pageBorders: null,
  pageLayout: {
    physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb' as const,
    gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
    bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
  },
});

describe('canonical body producer', () => {
  it('is the sole page owner and returns retained page nodes', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const layouts = [paragraph('p0', source(0), 60), paragraph('p1', source(1), 60)];
    const kernel: BodyLayoutKernel = {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: layouts[input.source.path[0]!]!, blockExtentPt: 60, lineEndBoundaries: [],
        }),
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    };
    attachBodyLayoutKernel(services, kernel);
    const owner = {
      sectionOccurrenceId: 'section:0', source: source(2), startType: 'nextPage' as const,
      context: section,
      pageNumbering: { start: null, format: null }, titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: owner,
      sequence: [0, 1].map((index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      })),
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1]]);
    expect(layout.pages.map((page) => page.readingOrder.length)).toEqual([1, 1]);
    expect(Object.isFrozen(layout)).toBe(true);
  });

  it('commits neither registry component for a rejected body candidate', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const baseFloatEntries = Object.freeze([]);
    const baseCollisionEntries = Object.freeze([]);
    const commits: unknown[] = [];
    let secondMeasurements = 0;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: (request) => {
          const index = request.input.source.path[0]!;
          if (index === 0) {
            return { layout: table('first', request.input.source, 60), blockExtentPt: 60 };
          }
          secondMeasurements += 1;
          const flowRegistryDelta = {
            floats: {
              coordinateSpace: 'logical-page-points' as const,
              flowDomainId: request.location.flowDomainId,
              baseEntries: baseFloatEntries,
              baseNextParagraphId: 0,
              nextParagraphId: 1,
              entries: [{
                kind: 'shape' as const,
                occurrenceId: 'candidate-float',
                paragraphId: 0,
                bounds: { xPt: 0, yPt: 0, widthPt: 10, heightPt: 10 },
                exclusionBounds: { xPt: 0, yPt: 0, widthPt: 10, heightPt: 10 },
              }],
            },
            drawingCollisions: {
              coordinateSpace: 'logical-page-points' as const,
              flowDomainId: request.location.flowDomainId,
              baseEntries: baseCollisionEntries,
              baseEntryCount: 0,
              entries: [{
                occurrenceId: 'candidate-collision',
                bounds: { xPt: 0, yPt: 0, widthPt: 10, heightPt: 10 },
                horizontalOwnership: 'page' as const,
                verticalOwnership: 'page' as const,
              }],
            },
          };
          return request.availableBlockExtentPt < 40
            ? {
                layout: table('rejected', request.input.source, 40),
                blockExtentPt: 40,
                requiresFreshFlowRegion: true,
                flowRegistryDelta,
              }
            : {
                layout: table('accepted', request.input.source, 40),
                blockExtentPt: 40,
                flowRegistryDelta,
              };
        },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: (delta) => { commits.push(delta); },
      }),
    });
    const input = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [0, 1].map((index) => ({
        kind: 'body-block',
        block: { kind: 'table', source: source(index) },
      })),
    } as unknown as BodyLayoutInput;

    paginateBody(input, services, { currentDateMs: 0 });

    expect(secondMeasurements).toBe(2);
    expect(commits).toHaveLength(1);
    expect(commits[0]).toMatchObject({
      floats: { entries: [{ occurrenceId: 'candidate-float' }] },
      drawingCollisions: { entries: [{ occurrenceId: 'candidate-collision' }] },
    });
  });

  it('acquires an adjacent-table sequence as one cursor-bearing logical table', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const acquiredInputs: string[] = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: ({ input }) => {
          acquiredInputs.push(input.kind);
          return { layout: table('logical-table', input.source, 30), blockExtentPt: 30 };
        },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = {
      sectionOccurrenceId: 'section:0', source: source(2), startType: 'nextPage' as const,
      context: section,
      pageNumbering: { start: null, format: null }, titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: owner,
      sequence: [{
        kind: 'adjacent-table-group', logicalSequenceId: 'logical:0', source: source(0),
        tables: [
          { kind: 'table', source: source(0), rowCount: 1 },
          { kind: 'table', source: source(1), rowCount: 2 },
        ],
      }],
    }, services, { currentDateMs: 0 });

    expect(acquiredInputs).toEqual(['adjacent-table-group']);
    expect(layout.pages[0]!.layers.body.map((node) => node.source.path)).toEqual([[0]]);
  });

  it('admits references from each accepted split-table slice only', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const admissions: Array<{ pageIndex: number; referenceIds: readonly string[] }> = [];
    let measuredPageIndex = -1;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: (request) => {
          measuredPageIndex = request.location.pageIndex;
          return request.cursor
            ? { layout: table('terminal-slice', request.input.source, 30), blockExtentPt: 30 }
            : {
                layout: tableWithFootnote('reference-slice', request.input.source, 30, 'fn-first'),
                blockExtentPt: 30,
                nextCursor: {
                  kind: 'table' as const,
                  cursor: { rowIndex: 1, rowFragmentIndex: 0, cells: [] },
                },
              };
        },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => {
          admissions.push({ pageIndex: measuredPageIndex, referenceIds: [...referenceIds] });
          return referenceIds.length === 0 ? 0 : 5;
        },
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const legacySourceWideInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [{
        kind: 'body-block',
        block: { kind: 'table', source: source(0), footnoteReferenceIds: ['fn-first'] },
      }],
    } as unknown as BodyLayoutInput;

    const layout = paginateBody(legacySourceWideInput, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(admissions).toEqual([
      { pageIndex: 0, referenceIds: ['fn-first'] },
      { pageIndex: 1, referenceIds: [] },
    ]);
  });

  it('does not union source-wide references across an adjacent table group', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const admissions: string[][] = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: ({ input }) => ({
          layout: tableWithFootnote('retained-left', input.source, 30, 'left'), blockExtentPt: 30,
        }),
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => {
          admissions.push([...referenceIds]);
          return 0;
        },
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const legacySourceWideInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [{
        kind: 'adjacent-table-group', logicalSequenceId: 'logical:0', source: source(0),
        tables: [
          { kind: 'table', source: source(0), rowCount: 1, footnoteReferenceIds: ['left'] },
          { kind: 'table', source: source(1), rowCount: 1, footnoteReferenceIds: ['right'] },
        ],
      }],
    } as unknown as BodyLayoutInput;

    paginateBody(legacySourceWideInput, services, { currentDateMs: 0 });

    expect(admissions).toEqual([['left']]);
  });

  it('anchors a continuous section page-number restart to its shared first page', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const heights = [20, 60, 30];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: paragraph(`p${input.source.path[0]}`, input.source, heights[input.source.path[0]!]!),
          blockExtentPt: heights[input.source.path[0]!]!, lineEndBoundaries: [],
        }),
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = (id: string, startType: 'nextPage' | 'continuous', start: number | null) => ({
      sectionOccurrenceId: id, source: source(3), startType, context: section,
      pageNumbering: { start, format: null }, titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    });
    const block = (index: number) => ({
      kind: 'body-block' as const,
      block: {
        kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
        keepLines: false, keepNext: false, widowControl: true,
        spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
      },
    });
    const first = owner('section:0', 'nextPage', null);
    const second = owner('section:1', 'continuous', 50);
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: first,
      sequence: [block(0), { kind: 'begin-section', source: source(3), section: second }, block(1), block(2)],
    }, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.pageNumber.displayNumber)).toEqual([1, 51]);
  });

  it('retains an upright physical table placement independently from its logical flow charge', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: paragraph('follower', input.source, 10),
          blockExtentPt: 10,
          lineEndBoundaries: [],
        }),
        measureTable: ({ input }) => ({
          layout: table('upright', input.source, 30), blockExtentPt: 20,
          placement: {
            coordinateSpace: 'upright-physical', xPt: 40, yPt: 25,
            sectionFlowOwnership: 'host-flow',
          },
        }),
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = {
      sectionOccurrenceId: 'section:0', source: source(1), startType: 'nextPage' as const,
      context: section, pageNumbering: { start: null, format: null },
      titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] }, initialSection: owner,
      sequence: [
        { kind: 'body-block', block: { kind: 'table', source: source(0) } },
        {
          kind: 'body-block',
          block: {
            kind: 'paragraph', source: source(1), pageBreakBefore: false,
            keepLines: false, keepNext: false, widowControl: true,
            spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
          },
        },
      ],
    }, services, { currentDateMs: 0 });

    const body = layout.pages[0]!.layers.body;
    expect(body[0]!.ordinaryFlow).toBe(false);
    expect(body[0]!.advancePt).toBe(30);
    expect(body[0]!.flowBounds).toMatchObject({ xPt: 40, yPt: 25, heightPt: 20 });
    expect(body[1]!.flowBounds.yPt).toBe(30);
    expect(layout.pages[0]!.layers.paintSequence[0]).toMatchObject({ coordinateSpace: 'upright-physical' });
  });

  it('assigns deterministic distinct occurrence ids to one table continued across two columns', () => {
    const twoColumnSection: SectionLayoutContext = {
      ...section,
      columns: [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }],
    };
    const paginate = () => {
      const services = Object.freeze({
        text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
      }) as LayoutServices;
      attachBodyLayoutKernel(services, {
        openBodyLayoutSession: () => ({
          hasPaginationFields: false,
          measureParagraph: () => { throw new Error('unused'); },
          measureTable: ({ input, cursor }) => {
            const retained = table('retained-table', input.source, cursor ? 20 : 80);
            const layout = {
              ...retained,
              flowBounds: { ...retained.flowBounds, widthPt: 80 },
              inkBounds: { ...retained.inkBounds, widthPt: 80 },
              columnWidthsPt: [80],
            };
            return cursor
            ? { layout, blockExtentPt: 20 }
            : {
                layout, blockExtentPt: 80,
                nextCursor: {
                  kind: 'table',
                  cursor: { rowIndex: 1, rowFragmentIndex: 0, cells: [] },
                },
              };
          },
          measureStoryExtent: () => 0,
          measureFootnoteReserve: () => 0,
          measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
          measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
          resetPageAcquisition: () => undefined,
          moveAcquisitionCursor: () => undefined,
          flowRegistrySnapshot: emptyFlowRegistrySnapshot,
          commitFlowRegistryDelta: () => undefined,
        }),
      });
      const owner = {
        sectionOccurrenceId: 'section:two-column', source: source(1), startType: 'nextPage' as const,
        context: twoColumnSection, pageNumbering: { start: null, format: null },
        titlePage: false, evenAndOddHeaders: false,
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        pageBordersAuthored: false, pageBorders: null,
        pageLayout: {
          physicalGeometry: twoColumnSection.geometry,
          columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
          textDirection: 'lrTb', gutterPt: 0, rtlGutter: false, mirrorMargins: false,
          gutterAtTop: false, bookFoldPrinting: false, bookFoldRevPrinting: false,
          printTwoOnOne: false,
        },
      };
      return paginateBody({
        source: { story: 'body', storyInstance: 'body', path: [] }, initialSection: owner,
        sequence: [{ kind: 'body-block', block: { kind: 'table', source: source(0) } }],
      }, services, { currentDateMs: 0 });
    };

    const layout = paginate();
    expect(layout.pages).toHaveLength(1);
    const body = layout.pages[0]!.layers.body;
    expect(body).toHaveLength(2);
    expect(body.map((node) => node.flowDomainId)).toEqual([
      layout.pages[0]!.sectionRegions[0]!.flowDomainIds[0],
      layout.pages[0]!.sectionRegions[0]!.flowDomainIds[1],
    ]);
    expect(new Set(body.map((node) => node.id)).size).toBe(2);
    expect(layout.pages[0]!.readingOrder).toEqual(body.map((node) => node.id));
    expect(paginate().pages[0]!.layers.body.map((node) => node.id))
      .toEqual(body.map((node) => node.id));
  });

  it('aligns an upright physical table by its logical allocation span', () => {
    const centeredSection: SectionLayoutContext = { ...section, verticalAlignment: 'center' };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: ({ input }) => ({
          layout: table('upright', input.source, 30), blockExtentPt: 20,
          placement: {
            coordinateSpace: 'upright-physical', xPt: 40, yPt: 25,
            sectionFlowOwnership: 'host-flow',
          },
        }),
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = {
      sectionOccurrenceId: 'section:centered', source: source(1), startType: 'nextPage' as const,
      context: centeredSection, pageNumbering: { start: null, format: null },
      titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: centeredSection.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] }, initialSection: owner,
      sequence: [{ kind: 'body-block', block: { kind: 'table', source: source(0) } }],
    }, services, { currentDateMs: 0 });
    const retained = layout.pages[0]!.layers.body[0]!;

    expect(retained.advancePt).toBe(30);
    expect(retained.flowBounds.yPt).toBe(55);
    expect(layout.pages[0]!.layers.paintSequence[0]).toEqual({
      layer: 'body', node: retained, coordinateSpace: 'upright-physical',
    });
    expect(layout.pages[0]!.layers.paintSequence[0]!.node).toBe(retained);
  });

  it('moves a feasible keepNext chain together before accepting any chain member', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const heights = [30, 20, 20, 20];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: paragraph(`p${input.source.path[0]}`, input.source, heights[input.source.path[0]!]!),
          blockExtentPt: heights[input.source.path[0]!]!, lineEndBoundaries: [],
        }),
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: ({ input }) => ({
          fullExtentPt: heights[input.source.path[0]!]!,
          leadContentExtentPt: heights[input.source.path[0]!]!,
        }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = {
      sectionOccurrenceId: 'section:0', source: source(3), startType: 'nextPage' as const,
      context: section, pageNumbering: { start: null, format: null },
      titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] }, initialSection: owner,
      sequence: heights.map((_height, index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: index === 1 || index === 2, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      })),
    }, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1, 2, 3]]);
  });

  it('preserves the current equal-height approximation for a non-final multi-column section', () => {
    const balancedSection: SectionLayoutContext = {
      ...section,
      columns: [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }],
    };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const layout = paragraph(`p${input.source.path[0]}`, input.source, 20);
          return {
            layout: {
              ...layout,
              flowBounds: { ...layout.flowBounds, widthPt: 80 },
              inkBounds: { ...layout.inkBounds, widthPt: 80 },
            },
            blockExtentPt: 20, lineEndBoundaries: [],
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 20, leadContentExtentPt: 20 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = (
      id: string,
      context: SectionLayoutContext,
      startType: 'continuous' | 'nextPage',
    ) => ({
      sectionOccurrenceId: id, source: source(7), startType, context,
      pageNumbering: { start: null, format: null }, titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: context.geometry, columns: context === balancedSection
          ? { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] }
          : null,
        textDirection: 'lrTb', gutterPt: 0, rtlGutter: false, mirrorMargins: false,
        gutterAtTop: false, bookFoldPrinting: false, bookFoldRevPrinting: false,
        printTwoOnOne: false,
      },
    });
    const first = owner('section:balanced', balancedSection, 'continuous');
    const final = owner('section:final', section, 'nextPage');
    const blocks = Array.from({ length: 6 }, (_, index) => ({
      kind: 'body-block' as const,
      block: {
        kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
        keepLines: false, keepNext: false, widowControl: true,
        spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
      },
    }));
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: first,
      sequence: [...blocks, { kind: 'begin-section', source: source(6), section: final }],
    }, services, { currentDateMs: 0 });
    const balancedNodes = layout.pages[0]!.layers.body;

    expect(balancedNodes.map((node) => node.flowBounds.xPt)).toEqual([10, 10, 10, 110, 110, 110]);
  });

  it('commits page-owned anchor exclusions before measuring earlier page content', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const events: string[] = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          events.push(`measure:${input.source.path[0]}`);
          return {
            layout: paragraph(`p${input.source.path[0]}`, input.source, 20),
            blockExtentPt: 20, lineEndBoundaries: [],
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 20, leadContentExtentPt: 20 }),
        prescanPageAnchors: ({ anchors, location }) => {
          events.push(`prescan:${anchors.map((anchor) => anchor.paragraphSource.path[0]).join(',')}`);
          return {
            floats: {
              coordinateSpace: 'logical-page-points', flowDomainId: location.flowDomainId,
              baseEntries: [],
              baseNextParagraphId: 0, nextParagraphId: 1,
              entries: [{
                kind: 'shape', occurrenceId: anchors[0]!.occurrenceId, paragraphId: 0,
                bounds: { xPt: 100, yPt: 10, widthPt: 20, heightPt: 20 },
                exclusionBounds: { xPt: 100, yPt: 10, widthPt: 20, heightPt: 20 },
              }],
            },
          };
        },
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => { events.push('commit'); },
      }),
    });
    const owner = {
      sectionOccurrenceId: 'section:0', source: source(2), startType: 'nextPage' as const,
      context: section, pageNumbering: { start: null, format: null },
      titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb', gutterPt: 0,
        rtlGutter: false, mirrorMargins: false, gutterAtTop: false, bookFoldPrinting: false,
        bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const block = (index: number, anchors: readonly string[] = []) => ({
      kind: 'body-block' as const,
      block: {
        kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
        keepLines: false, keepNext: false, widowControl: true, spaceBeforePt: 0, spaceAfterPt: 0,
        contextualSpacing: false, styleId: null,
        ...(anchors.length === 0 ? {} : { pageOwnedAnchorOccurrenceIds: anchors }),
      },
    });

    paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] }, initialSection: owner,
      sequence: [block(0), block(1, ['anchor:1'])],
    }, services, { currentDateMs: 0 });

    expect(events.slice(0, 3)).toEqual(['prescan:1', 'commit', 'measure:0']);
  });
});

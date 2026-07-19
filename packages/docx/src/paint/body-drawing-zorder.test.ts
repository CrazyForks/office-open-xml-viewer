import { describe, expect, it, vi } from 'vitest';
import { stableFingerprint } from '../layout/fingerprint.js';
import { buildPageLayers } from '../layout/page-graph.js';
import type {
  DrawingLayout,
  LayoutPage,
  LayoutRect,
  PaintNode,
  ParagraphLayout,
  ResolvedBorderSegment,
  ResolvedFloatingTablePlacementLayout,
  TableLayout,
  TextBoxLayout,
} from '../layout/types.js';
import { paintLayoutPageContent } from './canvas-page.js';
import type { CanvasPaintContext, PaintCanvas2D } from './types.js';

const bounds = { xPt: 0, yPt: 0, widthPt: 10, heightPt: 10 } as const;
const source = (path: readonly number[]) => ({
  story: 'body' as const,
  storyInstance: 'body',
  path,
});
const fontRoute = {
  familyList: 'Test Sans', scope: 'native', fingerprint: 'test-sans',
} as const;

function drawing(
  id: string,
  fill: string,
  behindDoc: boolean,
  relativeHeight: number,
  sourceOrder: number,
  textBoxIds: readonly string[] = [],
  ownership: Readonly<{
    horizontal: 'host' | 'page';
    vertical: 'host' | 'page';
  }> = { horizontal: 'host', vertical: 'host' },
): DrawingLayout {
  return {
    kind: 'drawing', id, source: source([sourceOrder]), flowDomainId: 'body',
    ordinaryFlow: false, flowBounds: bounds, inkBounds: bounds, advancePt: 0,
    anchorLayer: {
      occurrenceId: `occurrence:${id}`, behindDoc, relativeHeight, sourceOrder,
      horizontalOwnership: ownership.horizontal, verticalOwnership: ownership.vertical,
    },
    textBoxIds,
    commands: [{ kind: 'fill-rect', rect: bounds, fill }],
  };
}

function table(
  id: string,
  blocks: readonly (ParagraphLayout | TableLayout)[],
  options: Readonly<{
    xPt?: number;
    yPt?: number;
    clipBounds?: LayoutRect;
    cellClipBounds?: LayoutRect;
  }> = {},
): TableLayout {
  const xPt = options.xPt ?? 0;
  const yPt = options.yPt ?? 0;
  const tableBounds = { xPt, yPt, widthPt: 80, heightPt: 40 };
  const cellBounds = { xPt, yPt, widthPt: 80, heightPt: 40 };
  return {
    kind: 'table', id, source: source([70]), flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: tableBounds, inkBounds: tableBounds, advancePt: 40,
    ...(options.clipBounds ? { clipBounds: options.clipBounds } : {}),
    columnWidthsPt: [80], borders: [], rows: [{
      kind: 'table-row', id: `${id}:row`, source: source([70, 0]),
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: tableBounds, inkBounds: tableBounds, advancePt: 40,
      heightPt: 40, contentHeightPt: 40,
      cells: [{
        kind: 'table-cell', id: `${id}:cell`, source: source([70, 0, 0]),
        flowDomainId: 'body', ordinaryFlow: true,
        flowBounds: cellBounds, inkBounds: cellBounds, contentBounds: cellBounds,
        advancePt: 40, verticalMerge: 'none', vAlign: 'top',
        ...(options.cellClipBounds ? { clipBounds: options.cellClipBounds } : {}),
        blocks: blocks.map((layout, index) => ({
          layout,
          offsetPt: index * 10,
          advancePt: layout.advancePt,
        })),
      }],
    }],
  };
}

function paragraph(
  id: string,
  text: string,
  drawings: readonly DrawingLayout[] = [],
  textBoxes: readonly TextBoxLayout[] = [],
): ParagraphLayout {
  const placements = text === '' ? [] : [{
    kind: 'text' as const, text, range: { start: 0, end: text.length },
    origin: { xPt: 0, yPt: 8 }, bounds, advancePt: 10,
    clusters: [{
      range: { start: 0, end: text.length }, offset: { xPt: 0, yPt: 0 }, advancePt: 10,
    }],
    paintOps: [{
      text, range: { start: 0, end: text.length }, offset: { xPt: 0, yPt: 0 },
      letterSpacingPt: 0, scaleX: 1, direction: 'ltr' as const,
      kerning: 'auto' as const, writingMode: 'horizontal-tb' as const,
    }],
    color: { kind: 'explicit' as const, color: '#000000' },
    fontRoute, fontSizePt: 10, fontWeight: 400 as const,
    fontStyle: 'normal' as const, direction: 'ltr' as const, decorations: [],
  }];
  return {
    kind: 'paragraph', id, source: source([Number(id.replace(/\D/gu, '')) || 0]),
    flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: bounds, inkBounds: bounds, advancePt: 10,
    spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
    lines: placements.length === 0 ? [] : [{
      range: { start: 0, end: text.length }, bounds, baselinePt: 8,
      advancePt: 10, placements,
    }],
    borders: [], resources: [], drawings, textBoxes, events: [], exclusions: [],
  };
}

function textBox(id: string, ownerText: string, nestedDrawings: readonly DrawingLayout[] = []): TextBoxLayout {
  return {
    kind: 'textbox', id, source: source([90]), flowDomainId: 'body', ordinaryFlow: false,
    flowBounds: bounds, inkBounds: bounds, advancePt: 0,
    story: {
      story: 'textbox',
      flowBounds: bounds,
      inkBounds: bounds,
      blocks: [paragraph(`${id}-paragraph`, ownerText, nestedDrawings)],
      advancePt: 10,
      diagnostics: [],
    },
    transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
    writingMode: 'horizontal-tb',
    insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
  };
}

function resolvedFloatingTable(
  id: string,
  child: TableLayout,
  xPt: number,
  yPt: number,
): ResolvedFloatingTablePlacementLayout {
  return {
    kind: 'resolved-floating-table-placement', id,
    occurrenceId: id, xPt, yPt,
    bounds: { xPt, yPt, widthPt: child.flowBounds.widthPt, heightPt: child.flowBounds.heightPt },
    exclusionBounds: {
      xPt, yPt, widthPt: child.flowBounds.widthPt, heightPt: child.flowBounds.heightPt,
    },
    overlap: 'never', child,
    source: {
      kind: 'floating-table-placement', occurrenceId: id,
      ownership: 'source', physicalPageIndex: 0, displayPageNumber: 1,
      hostCellId: 'nested:cell', sourceBlockIndex: 0, anchorBlockIndex: 0,
      tableId: child.id, overlap: 'never',
      positioning: {
        leftFromTextPt: 0, rightFromTextPt: 0,
        topFromTextPt: 0, bottomFromTextPt: 0,
        horzAnchor: 'text', horzSpecified: true, vertAnchor: 'text',
        xPt: 0, yPt: 0,
      },
      anchorBounds: bounds,
      child,
    },
  } as ResolvedFloatingTablePlacementLayout;
}

function page(
  body: readonly PaintNode[],
  logicalToPhysical = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
): LayoutPage {
  return {
    pageIndex: 0,
    geometry: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 100, contentTopPt: 0, contentBottomPt: 100 },
    flowDomains: [{
      id: 'body', kind: 'body', logicalBounds: bounds, physicalBounds: bounds,
    }],
    section: {} as LayoutPage['section'], sectionOccurrenceId: 'section:0', parityBlank: false,
    bookmarkStarts: [],
    pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
    columnSeparators: [],
    sectionRegions: [{
      id: 'region:0', sectionOccurrenceId: 'section:0', section: {} as LayoutPage['section'],
      coordinateSpace: {
        writingMode: 'horizontal-tb', logicalToPhysical,
        physicalToLogical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      },
      blockStartPt: 0, blockEndPt: 100,
      columnFlowDirection: 'ltr', columnIndexes: [0],
      flowDomainIds: ['body'],
    }],
    pageBorders: null,
    layers: buildPageLayers(
      body.map((node) => ({
        layer: 'body' as const, node, coordinateSpace: 'section-logical' as const,
      })),
    ),
    readingOrder: body.map((node) => node.id),
  };
}

function recordingPaintContext() {
  const operations: string[] = [];
  const matrices: Array<readonly number[]> = [];
  const drawingRecords: Array<Readonly<{
    fill: string;
    matrix: readonly number[];
    clips: readonly Readonly<{ rect: readonly number[]; matrix: readonly number[] }>[];
  }>> = [];
  let matrix = [1, 0, 0, 1, 0, 0];
  let clips: Array<Readonly<{ rect: readonly number[]; matrix: readonly number[] }>> = [];
  let pendingRect: readonly number[] | undefined;
  let currentFill = '';
  const stack: Array<Readonly<{
    matrix: number[];
    clips: Array<Readonly<{ rect: readonly number[]; matrix: readonly number[] }>>;
    pendingRect?: readonly number[];
  }>> = [];
  const apply = (next: readonly number[]): void => {
    const [a, b, c, d, e, f] = matrix;
    const [na, nb, nc, nd, ne, nf] = next;
    matrix = [
      a! * na! + c! * nb!, b! * na! + d! * nb!,
      a! * nc! + c! * nd!, b! * nc! + d! * nd!,
      a! * ne! + c! * nf! + e!, b! * ne! + d! * nf! + f!,
    ];
  };
  const ctx = {
    get fillStyle() { return currentFill; },
    set fillStyle(value: string | CanvasGradient | CanvasPattern) { currentFill = String(value); },
    strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
    textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px',
    fontKerning: 'auto', globalAlpha: 1,
    save() { stack.push({ matrix: [...matrix], clips: [...clips], pendingRect }); },
    restore() {
      const restored = stack.pop();
      if (!restored) return;
      matrix = restored.matrix;
      clips = restored.clips;
      pendingRect = restored.pendingRect;
    },
    transform(a: number, b: number, c: number, d: number, e: number, f: number) {
      apply([a, b, c, d, e, f]);
    },
    translate(x: number, y: number) { apply([1, 0, 0, 1, x, y]); },
    rotate(angle: number) {
      apply([Math.cos(angle), Math.sin(angle), -Math.sin(angle), Math.cos(angle), 0, 0]);
    },
    scale(x: number, y: number) { apply([x, 0, 0, y, 0, 0]); },
    setLineDash() {}, beginPath() {},
    rect(x: number, y: number, width: number, height: number) {
      pendingRect = [x, y, width, height];
    },
    clip() {
      if (pendingRect) clips.push({ rect: pendingRect, matrix: [...matrix] });
    },
    moveTo() {}, lineTo() {}, stroke() { operations.push('stroke'); }, fill() {}, strokeRect() {},
    fillRect() {
      operations.push(`drawing:${currentFill}`);
      matrices.push([...matrix]);
      drawingRecords.push({ fill: currentFill, matrix: [...matrix], clips: [...clips] });
    },
    fillText(text: string) { operations.push(`text:${text}`); },
    measureText: vi.fn(() => { throw new Error('paint must not measure text'); }),
    drawImage() {},
  } as unknown as PaintCanvas2D;
  const context: CanvasPaintContext = {
    ctx, scale: 1, dpr: 1, resources: { paint: vi.fn() },
  };
  return {
    context, operations, matrices, drawingRecords,
    measureText: (ctx as unknown as { measureText: ReturnType<typeof vi.fn> }).measureText,
  };
}

describe('canonical body drawing z-order', () => {
  it('keeps the materialized body drawing context between surrounding page layers', () => {
    const behind = drawing('behind', 'behind', true, 5, 1);
    const front = drawing('front', 'front', false, 10, 2);
    const header = drawing('header', 'header', false, 0, 0);
    const footer = drawing('footer', 'footer', false, 0, 0);
    const body = paragraph('p1', 'BODY', [behind, front]);
    const base = page([body]);
    const layout: LayoutPage = {
      ...base,
      layers: buildPageLayers([
          { layer: 'header', node: header, coordinateSpace: 'section-logical' },
          { layer: 'body', node: body, coordinateSpace: 'section-logical' },
          { layer: 'footer', node: footer, coordinateSpace: 'section-logical' },
      ]),
    };
    const { context, operations } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(operations).toEqual([
      'drawing:header', 'drawing:behind', 'text:BODY', 'drawing:front', 'drawing:footer',
    ]);
  });

  it('paints page-wide behind drawings before text in z-order with stable ties', () => {
    const layout = page([
      paragraph('p1', 'EARLY', [drawing('behind-high', 'behind-high', true, 20, 4)]),
      paragraph('p2', 'LATER', [drawing('behind-tie-1', 'behind-tie-1', true, 10, 2)]),
      paragraph('p3', 'LAST', [drawing('behind-tie-2', 'behind-tie-2', true, 10, 2)]),
    ]);
    const { context, operations } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(operations).toEqual([
      'drawing:behind-tie-1', 'drawing:behind-tie-2', 'drawing:behind-high',
      'text:EARLY', 'text:LATER', 'text:LAST',
    ]);
  });

  it('sorts page-wide front drawings by relativeHeight, source order, then stable encounter order', () => {
    const layout = page([
      paragraph('p1', 'ONE', [drawing('high', 'high', false, 20, 4)]),
      paragraph('p2', 'TWO', [drawing('tie-1', 'tie-1', false, 10, 2)]),
      paragraph('p3', 'THREE', [drawing('tie-2', 'tie-2', false, 10, 2)]),
    ]);
    const { context, operations } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(operations).toEqual([
      'text:ONE', 'text:TWO', 'text:THREE',
      'drawing:tie-1', 'drawing:tie-2', 'drawing:high',
    ]);
  });

  it('keeps an owned text box adjacent to its materialized owner and paints both exactly once', () => {
    const owned = textBox('owned', 'OWNED');
    const owner = drawing('owner', 'owner', false, 10, 1, [owned.id]);
    const layout = page([
      paragraph('p1', 'BODY', [owner], [owned]),
      paragraph('p2', 'LATER'),
    ]);
    const { context, operations } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(operations).toEqual(['text:BODY', 'text:LATER', 'drawing:owner', 'text:OWNED']);
  });

  it('keeps anchors in a non-owned text-box story local when page anchors are materialized', () => {
    const nested = drawing('nested', 'nested', false, 20, 2);
    const local = textBox('local', 'LOCAL', [nested]);
    const pageFront = drawing('page-front', 'page-front', false, 10, 1);
    const layout = page([
      paragraph('p1', 'BODY', [pageFront], [local]),
    ]);
    const { context, operations } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(operations).toEqual([
      'text:BODY', 'text:LOCAL', 'drawing:nested', 'drawing:page-front',
    ]);
  });

  it('paints a behind owner and its owned text box atomically before body ink', () => {
    const owned = textBox('owned', 'OWNED');
    const owner = drawing('owner', 'owner', true, 10, 1, [owned.id]);
    const layout = page([paragraph('p1', 'BODY', [owner], [owned])]);
    const { context, operations } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(operations).toEqual(['drawing:owner', 'text:OWNED', 'text:BODY']);
  });

  it('keeps a nested behind anchor local and adjacent to its behind owner z-position', () => {
    const nested = drawing('nested', 'nested', true, 100, 3);
    const owned = textBox('owned', 'OWNED', [nested]);
    const owner = drawing('owner', 'owner', true, 10, 1, [owned.id]);
    const later = drawing('later', 'later', true, 20, 2);
    const layout = page([
      paragraph('p1', 'BODY', [owner], [owned]),
      paragraph('p2', 'LATER', [later]),
    ]);
    const { context, operations } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(operations).toEqual([
      'drawing:owner', 'drawing:nested', 'text:OWNED', 'drawing:later',
      'text:BODY', 'text:LATER',
    ]);
  });

  it('re-enters the owner entry matrix and prevents nested story replay from re-enqueueing', () => {
    const nested = drawing('nested', 'nested', false, 30, 3);
    const owned = textBox('owned', 'OWNED', [nested]);
    const owner = drawing('owner', 'owner', false, 10, 1, [owned.id]);
    const layout = page(
      [paragraph('p1', 'BODY', [owner], [owned]), paragraph('p2', 'LATER')],
      { a: 0, b: 1, c: -1, d: 0, e: 100, f: 5 },
    );
    const { context, operations, matrices } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(operations).toEqual([
      'text:BODY', 'text:LATER', 'drawing:owner', 'text:OWNED', 'drawing:nested',
    ]);
    expect(matrices).toEqual([
      [0, 1, -1, 0, 100, 5],
      [0, 1, -1, 0, 100, 5],
    ]);
  });

  it('does not measure or mutate retained layout while painting ordered drawings', () => {
    const layout = page([
      paragraph('p1', 'BODY', [
        drawing('behind', 'behind', true, 5, 1),
        drawing('front', 'front', false, 10, 2),
      ]),
    ]);
    const before = stableFingerprint('page', layout);
    const { context, measureText } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(measureText).not.toHaveBeenCalled();
    expect(stableFingerprint('page', layout)).toBe(before);
  });

  it('replays a host-owned front drawing through its non-zero table-cell placement', () => {
    const front = drawing('front', 'front', false, 10, 1);
    const layout = page([table('table', [paragraph('p1', 'CELL', [front])], {
      xPt: 20, yPt: 30,
    })]);
    const { context, drawingRecords } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(drawingRecords.find((record) => record.fill === 'front')?.matrix).toEqual([
      1, 0, 0, 1, 20, 30,
    ]);
  });

  it('replays a page-owned front drawing without retaining its table-cell placement', () => {
    const front = drawing('front', 'front', false, 10, 1, [], {
      horizontal: 'page', vertical: 'page',
    });
    const layout = page([table('table', [paragraph('p1', 'CELL', [front])], {
      xPt: 20, yPt: 30,
    })]);
    const { context, drawingRecords } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(drawingRecords.find((record) => record.fill === 'front')?.matrix).toEqual([
      1, 0, 0, 1, 0, 0,
    ]);
  });

  it('sorts nested-table behind drawings with root behind drawings before body ink', () => {
    const nestedBehind = drawing('nested-behind', 'nested-behind', true, 10, 2);
    const rootBehind = drawing('root-behind', 'root-behind', true, 20, 1);
    const nested = table('nested', [paragraph('p2', 'NESTED', [nestedBehind])]);
    const layout = page([
      paragraph('p1', 'ROOT', [rootBehind]),
      table('outer', [nested], { xPt: 20, yPt: 30 }),
    ]);
    const { context, operations } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(operations).toEqual([
      'drawing:nested-behind', 'drawing:root-behind', 'text:ROOT', 'text:NESTED',
    ]);
  });

  it('discovers and paints a nested table resolved floating child exactly once', () => {
    const behind = drawing('floating-behind', 'floating-behind', true, 5, 1);
    const front = drawing('floating-front', 'floating-front', false, 10, 2);
    const floatingBase = table(
      'floating',
      [paragraph('p3', 'FLOATING', [behind, front])],
      { clipBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 100 } },
    );
    const floatingCell = floatingBase.rows[0]!.cells[0]!;
    const floating: TableLayout = {
      ...floatingBase,
      rows: [{
        ...floatingBase.rows[0]!,
        cells: [{ ...floatingCell, background: { color: 'floating-background' } }],
      }],
      borders: [{
        edge: 'left', from: { xPt: 0, yPt: 0 }, to: { xPt: 0, yPt: 40 },
        color: '#123456', widthPt: 1, authoredStyle: 'single', style: 'solid',
      } as ResolvedBorderSegment],
    };
    const nestedBase = table('nested', [], {
      clipBounds: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 200 },
    });
    const nested: TableLayout = {
      ...nestedBase,
      resolvedFloatingTables: [resolvedFloatingTable('nested:float', floating, 70, 90)],
    };
    const outer = table('outer', [nested], {
      xPt: 20, yPt: 30,
      clipBounds: { xPt: 20, yPt: 30, widthPt: 200, heightPt: 200 },
      cellClipBounds: { xPt: 20, yPt: 30, widthPt: 180, heightPt: 180 },
    });
    const layout = page([outer]);
    const { context, operations, drawingRecords, measureText } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(operations).toEqual([
      'drawing:floating-behind',
      'drawing:floating-background',
      'text:FLOATING',
      'stroke',
      'drawing:floating-front',
    ]);
    for (const fill of ['floating-behind', 'floating-front']) {
      const record = drawingRecords.find((item) => item.fill === fill);
      expect(record?.matrix).toEqual([1, 0, 0, 1, 70, 90]);
      expect(record?.clips).toEqual([
        { rect: [20, 30, 200, 200], matrix: [1, 0, 0, 1, 0, 0] },
        { rect: [20, 30, 180, 180], matrix: [1, 0, 0, 1, 0, 0] },
        { rect: [0, 0, 200, 200], matrix: [1, 0, 0, 1, 20, 30] },
        { rect: [0, 0, 100, 100], matrix: [1, 0, 0, 1, 70, 90] },
      ]);
    }
    expect(measureText).not.toHaveBeenCalled();
  });

  it('replays host- and page-owned front drawings under the retained vertical section matrix', () => {
    const host = drawing('host', 'host', false, 10, 1);
    const pageOwned = drawing('page', 'page', false, 20, 2, [], {
      horizontal: 'page', vertical: 'page',
    });
    const layout = page(
      [table('table', [paragraph('p1', 'CELL', [host, pageOwned])], { xPt: 20, yPt: 30 })],
      { a: 0, b: 1, c: -1, d: 0, e: 100, f: 5 },
    );
    const { context, drawingRecords } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(drawingRecords.find((record) => record.fill === 'host')?.matrix).toEqual([
      0, 1, -1, 0, 70, 25,
    ]);
    expect(drawingRecords.find((record) => record.fill === 'page')?.matrix).toEqual([
      0, 1, -1, 0, 100, 5,
    ]);
  });

  it('re-enters table and cell clip frames while replaying a front drawing', () => {
    const front = drawing('front', 'front', false, 10, 1);
    const tableClip = { xPt: 20, yPt: 30, widthPt: 60, heightPt: 30 } as const;
    const cellClip = { xPt: 22, yPt: 32, widthPt: 50, heightPt: 20 } as const;
    const layout = page([table('table', [paragraph('p1', 'CELL', [front])], {
      xPt: 20, yPt: 30, clipBounds: tableClip, cellClipBounds: cellClip,
    })]);
    const { context, drawingRecords } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(drawingRecords.find((record) => record.fill === 'front')?.clips).toEqual([
      { rect: [20, 30, 60, 30], matrix: [1, 0, 0, 1, 0, 0] },
      { rect: [22, 32, 50, 20], matrix: [1, 0, 0, 1, 0, 0] },
    ]);
  });

  it('preserves arbitrary non-body ordering before one contiguous body run', () => {
    const front = drawing('front', 'front', false, 0, 0);
    const body = paragraph('p1', 'BODY');
    const base = page([body]);
    const layout: LayoutPage = {
      ...base,
      layers: buildPageLayers([
          { layer: 'front', node: front, coordinateSpace: 'section-logical' },
          { layer: 'body', node: body, coordinateSpace: 'section-logical' },
      ]),
    };
    const { context, operations } = recordingPaintContext();

    paintLayoutPageContent(layout, context);

    expect(operations).toEqual(['drawing:front', 'text:BODY']);
  });

  it('relies on the completed contract while preserving a contiguous body run', () => {
    const firstBody = paragraph('p1', 'FIRST');
    const footer = drawing('footer', 'footer', false, 0, 0);
    const secondBody = paragraph('p2', 'SECOND');
    const base = page([firstBody, secondBody]);
    const layout: LayoutPage = {
      ...base,
      layers: buildPageLayers([
          { layer: 'body', node: firstBody, coordinateSpace: 'section-logical' },
          { layer: 'body', node: secondBody, coordinateSpace: 'section-logical' },
          { layer: 'footer', node: footer, coordinateSpace: 'section-logical' },
      ]),
    };
    const { context, operations } = recordingPaintContext();

    paintLayoutPageContent(layout, context);
    expect(operations).toEqual(['text:FIRST', 'text:SECOND', 'drawing:footer']);
  });

  it('fails loudly for unsupported textbox page nodes', () => {
    const kind = 'textbox' as const;
    const unsupported = {
      kind,
      id: `unsupported:${kind}`,
      source: source([99]),
      flowDomainId: 'body', ordinaryFlow: false,
      flowBounds: bounds, inkBounds: bounds, advancePt: 0,
      story: {
        story: 'body',
        flowBounds: bounds,
        inkBounds: bounds,
        blocks: [],
        advancePt: 0,
        diagnostics: [],
      },
      transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      writingMode: 'horizontal-tb',
      insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
    } as PaintNode;
    const layout = page([unsupported]);
    const { context } = recordingPaintContext();

    expect(() => paintLayoutPageContent(layout, context)).toThrow(
      new RegExp(`unsupported page paint node kind: ${kind}`, 'i'),
    );
  });
});

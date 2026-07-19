import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import type {
  DocumentLayout,
  NoteLayout,
  ParagraphLayout,
} from '../layout/types.js';
import { createLayoutPage } from '../layout/page-factory.js';
import { assertDocumentLayout } from '../layout/invariants.js';
import type { PageLayerId } from '../layout/types.js';
import { paintLayoutPage, paintLayoutPageContent } from './canvas-page.js';

const canonicalPageMeta = (section: SectionLayoutContext) => ({
  sectionOccurrenceId: 'section:0',
  parityBlank: false,
  bookmarkStarts: [],
  pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
  columnSeparators: [],
  sectionRegions: [{
    id: 'region:0', sectionOccurrenceId: 'section:0', section,
    coordinateSpace: {
      writingMode: 'horizontal-tb' as const,
      logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      physicalToLogical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
    },
    blockStartPt: 10, blockEndPt: 190,
    columnFlowDirection: 'ltr' as const, columnIndexes: [0],
    flowDomainIds: ['body'],
  }],
  pageBorders: null,
});

describe('paintLayoutPage', () => {
  it('paints section separators after leading retained layers and before body ink', () => {
    const events: string[] = [];
    let fillStyle = '';
    const ctx = {
      save() {}, restore() {}, beginPath() {},
      moveTo() {}, lineTo() {},
      stroke() { events.push('separator'); },
      fillRect() { events.push(fillStyle); },
      get fillStyle() { return fillStyle; },
      set fillStyle(value: string | CanvasGradient | CanvasPattern) { fillStyle = String(value); },
      strokeStyle: '', lineWidth: 1,
    } as unknown as CanvasRenderingContext2D;
    const section: SectionLayoutContext = {
      geometry: {
        pageWidth: 100, pageHeight: 200,
        marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
        headerDistance: 5, footerDistance: 5,
      },
      columns: [{ xPt: 10, wPt: 30 }, { xPt: 60, wPt: 30 }],
      columnSeparator: true,
      grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
      textDirection: 'lrTb', verticalAlignment: 'top',
    };
    const layers: PageLayerId[] = ['background', 'behindText', 'header', 'body'];
    const nodes = layers.map((layer, index) => ({
      kind: 'drawing' as const,
      id: layer,
      source: { story: 'body' as const, storyInstance: 'body', path: [index] },
      flowDomainId: `page:0:region:region%3A0:column:0`,
      ordinaryFlow: layer === 'body',
      flowBounds: { xPt: 10, yPt: 10, widthPt: 1, heightPt: 1 },
      inkBounds: { xPt: 10, yPt: 10, widthPt: 1, heightPt: 1 },
      advancePt: 1,
      commands: [{
        kind: 'fill-rect' as const,
        rect: { xPt: 10, yPt: 10, widthPt: 1, heightPt: 1 },
        fill: layer,
      }],
    }));
    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: { widthPt: 100, heightPt: 200, contentTopPt: 10, contentBottomPt: 190 },
      sectionOccurrenceId: 'section:0', section,
      sectionRegions: [{
        id: 'region:0', sectionOccurrenceId: 'section:0', section,
        writingMode: 'horizontal-tb', blockStartPt: 10, blockEndPt: 190,
        columns: [
          { inlineStartPt: 10, inlineExtentPt: 30 },
          { inlineStartPt: 60, inlineExtentPt: 30 },
        ],
      }],
      paint: nodes.map((node, index) => ({ layer: layers[index]!, node })),
      readingOrder: nodes,
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
    });
    expect(() => assertDocumentLayout({ pages: [page], diagnostics: [] })).not.toThrow();

    paintLayoutPageContent(page, {
      ctx,
      scale: 1,
      dpr: 1,
    } as unknown as Parameters<typeof paintLayoutPageContent>[1]);

    expect(events).toEqual(['background', 'behindText', 'header', 'separator', 'body']);
  });

  it('paints section separators at the body boundary when the page has no body entry', () => {
    const events: string[] = [];
    let fillStyle = '';
    const ctx = {
      save() {}, restore() {}, beginPath() {}, moveTo() {}, lineTo() {},
      stroke() { events.push('separator'); },
      fillRect() { events.push(fillStyle); },
      get fillStyle() { return fillStyle; },
      set fillStyle(value: string | CanvasGradient | CanvasPattern) { fillStyle = String(value); },
      strokeStyle: '', lineWidth: 1,
    } as unknown as CanvasRenderingContext2D;
    const section: SectionLayoutContext = {
      geometry: {
        pageWidth: 100, pageHeight: 200,
        marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
        headerDistance: 5, footerDistance: 5,
      },
      columns: [{ xPt: 10, wPt: 30 }, { xPt: 60, wPt: 30 }],
      columnSeparator: true,
      grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
      textDirection: 'lrTb', verticalAlignment: 'top',
    };
    const layers: PageLayerId[] = ['background', 'header', 'notes', 'front'];
    const nodes = layers.map((layer, index) => ({
      kind: 'drawing' as const,
      id: layer,
      source: { story: 'body' as const, storyInstance: 'body', path: [index] },
      flowDomainId: 'page:0:region:region%3A0:column:0',
      ordinaryFlow: false,
      flowBounds: { xPt: 10, yPt: 10, widthPt: 1, heightPt: 1 },
      inkBounds: { xPt: 10, yPt: 10, widthPt: 1, heightPt: 1 },
      advancePt: 0,
      commands: [{
        kind: 'fill-rect' as const,
        rect: { xPt: 10, yPt: 10, widthPt: 1, heightPt: 1 },
        fill: layer,
      }],
    }));
    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: { widthPt: 100, heightPt: 200, contentTopPt: 10, contentBottomPt: 190 },
      sectionOccurrenceId: 'section:0', section,
      sectionRegions: [{
        id: 'region:0', sectionOccurrenceId: 'section:0', section,
        writingMode: 'horizontal-tb', blockStartPt: 10, blockEndPt: 190,
        columns: [
          { inlineStartPt: 10, inlineExtentPt: 30 },
          { inlineStartPt: 60, inlineExtentPt: 30 },
        ],
      }],
      paint: nodes.map((node, index) => ({ layer: layers[index]!, node })),
      readingOrder: nodes,
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
    });
    expect(() => assertDocumentLayout({ pages: [page], diagnostics: [] })).not.toThrow();

    paintLayoutPageContent(page, {
      ctx, scale: 1, dpr: 1,
    } as unknown as Parameters<typeof paintLayoutPageContent>[1]);

    expect(events).toEqual(['background', 'header', 'separator', 'notes', 'front']);
  });

  it('paints separators before a retained front entry that precedes the body run', () => {
    const events: string[] = [];
    let fillStyle = '';
    const ctx = {
      save() {}, restore() {}, beginPath() {}, moveTo() {}, lineTo() {},
      stroke() { events.push('separator'); },
      fillRect() { events.push(fillStyle); },
      get fillStyle() { return fillStyle; },
      set fillStyle(value: string | CanvasGradient | CanvasPattern) { fillStyle = String(value); },
      strokeStyle: '', lineWidth: 1,
    } as unknown as CanvasRenderingContext2D;
    const section: SectionLayoutContext = {
      geometry: {
        pageWidth: 100, pageHeight: 200,
        marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
        headerDistance: 5, footerDistance: 5,
      },
      columns: [{ xPt: 10, wPt: 30 }, { xPt: 60, wPt: 30 }],
      columnSeparator: true,
      grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
      textDirection: 'lrTb', verticalAlignment: 'top',
    };
    const layers: PageLayerId[] = ['front', 'body'];
    const nodes = layers.map((layer, index) => ({
      kind: 'drawing' as const,
      id: layer,
      source: { story: 'body' as const, storyInstance: 'body', path: [index] },
      flowDomainId: 'page:0:region:region%3A0:column:0',
      ordinaryFlow: layer === 'body',
      flowBounds: { xPt: 10, yPt: 10, widthPt: 1, heightPt: 1 },
      inkBounds: { xPt: 10, yPt: 10, widthPt: 1, heightPt: 1 },
      advancePt: layer === 'body' ? 1 : 0,
      commands: [{
        kind: 'fill-rect' as const,
        rect: { xPt: 10, yPt: 10, widthPt: 1, heightPt: 1 },
        fill: layer,
      }],
    }));
    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: { widthPt: 100, heightPt: 200, contentTopPt: 10, contentBottomPt: 190 },
      sectionOccurrenceId: 'section:0', section,
      sectionRegions: [{
        id: 'region:0', sectionOccurrenceId: 'section:0', section,
        writingMode: 'horizontal-tb', blockStartPt: 10, blockEndPt: 190,
        columns: [
          { inlineStartPt: 10, inlineExtentPt: 30 },
          { inlineStartPt: 60, inlineExtentPt: 30 },
        ],
      }],
      paint: nodes.map((node, index) => ({ layer: layers[index]!, node })),
      readingOrder: nodes,
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
    });
    expect(() => assertDocumentLayout({ pages: [page], diagnostics: [] })).not.toThrow();

    paintLayoutPageContent(page, {
      ctx, scale: 1, dpr: 1,
    } as unknown as Parameters<typeof paintLayoutPageContent>[1]);

    expect(events).toEqual(['separator', 'front', 'body']);
  });

  it('paints a note separator outside the retained story clip', () => {
    const events: string[] = [];
    const ctx = {
      save() { events.push('save'); },
      restore() { events.push('restore'); },
      beginPath() { events.push('begin'); },
      rect(x: number, y: number, width: number, height: number) {
        events.push(`rect:${x},${y},${width},${height}`);
      },
      clip() { events.push('clip'); },
      moveTo() {}, lineTo() {},
      stroke() { events.push('stroke'); },
      setLineDash() {},
      strokeStyle: '', lineWidth: 1,
    } as unknown as CanvasRenderingContext2D;
    const section: SectionLayoutContext = {
      geometry: {
        pageWidth: 100, pageHeight: 200,
        marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
        headerDistance: 5, footerDistance: 5,
      },
      columns: [{ xPt: 10, wPt: 80 }],
      columnSeparator: false,
      grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
      textDirection: 'lrTb', verticalAlignment: 'top',
    };
    const flowDomainId = 'page:0:region:region%3A0:column:0';
    const source = { story: 'footnote' as const, storyInstance: '1', path: [0] };
    const child: ParagraphLayout = {
      kind: 'paragraph', id: 'note-child', source, flowDomainId,
      ordinaryFlow: true,
      flowBounds: { xPt: 10, yPt: 150, widthPt: 80, heightPt: 10 },
      inkBounds: { xPt: 10, yPt: 150, widthPt: 80, heightPt: 10 },
      advancePt: 10, spacing: { beforePt: 0, afterPt: 0 },
      contextualSpacing: false, lines: [], borders: [], resources: [],
      drawings: [], textBoxes: [], events: [], exclusions: [],
    };
    const note: NoteLayout = {
      kind: 'note', id: 'note', source: { ...source, path: [] }, flowDomainId,
      ordinaryFlow: false,
      flowBounds: { xPt: 10, yPt: 144, widthPt: 80, heightPt: 16 },
      inkBounds: { xPt: 10, yPt: 144, widthPt: 80, heightPt: 16 },
      advancePt: 16,
      separator: [{
        edge: 'top', from: { xPt: 10, yPt: 147 }, to: { xPt: 30, yPt: 147 },
        color: '#000000', widthPt: 0.5, authoredStyle: 'single', style: 'solid',
      }],
      story: {
        story: 'footnote',
        flowBounds: child.flowBounds,
        inkBounds: child.inkBounds,
        clipBounds: { xPt: 10, yPt: 150, widthPt: 80, heightPt: 10 },
        blocks: [child],
        advancePt: 10,
      },
    };
    const page = createLayoutPage({
      pageIndex: 0,
      physicalPage: { widthPt: 100, heightPt: 200, contentTopPt: 10, contentBottomPt: 190 },
      sectionOccurrenceId: 'section:0', section,
      sectionRegions: [{
        id: 'region:0', sectionOccurrenceId: 'section:0', section,
        writingMode: 'horizontal-tb', blockStartPt: 10, blockEndPt: 190,
        columns: [{ inlineStartPt: 10, inlineExtentPt: 80 }],
      }],
      paint: [{ layer: 'notes', node: note }],
      readingOrder: [note],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
    });

    paintLayoutPageContent(page, {
      ctx, scale: 1, dpr: 1,
    } as unknown as Parameters<typeof paintLayoutPageContent>[1]);

    expect(events.indexOf('stroke')).toBeLessThan(events.indexOf('rect:10,150,80,10'));
    expect(events).toContain('clip');
  });

  it('paints retained geometry without measuring text', async () => {
    const fills: Array<{ fill: string; args: number[] }> = [];
    let currentFill = '';
    const context = {
      get fillStyle() { return currentFill; },
      set fillStyle(value: string | CanvasGradient | CanvasPattern) { currentFill = String(value); },
      save() {},
      restore() {},
      setTransform() {},
      clearRect() {},
      fillRect(...args: number[]) {
        fills.push({ fill: currentFill, args });
      },
      measureText() {
        throw new Error('paint must not measure text');
      },
    } as unknown as CanvasRenderingContext2D;
    const target = {
      width: 0,
      height: 0,
      getContext: () => context,
    } as unknown as HTMLCanvasElement;
    const node = {
      kind: 'drawing' as const,
      id: 'drawing-1',
      source: { story: 'body' as const, storyInstance: 'body', path: [0] },
      flowBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
      inkBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
      advancePt: 40,
      ordinaryFlow: true,
      flowDomainId: 'body',
      commands: [{
        kind: 'fill-rect' as const,
        rect: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
        fill: '#ff0000',
      }],
    };
    const layout: DocumentLayout = {
      pages: [{
        pageIndex: 0,
        geometry: {
          xPt: 0,
          yPt: 0,
          widthPt: 100,
          heightPt: 200,
          contentTopPt: 10,
          contentBottomPt: 190,
        },
        flowDomains: [{
          id: 'body',
          kind: 'body',
          logicalBounds: { xPt: 10, yPt: 10, widthPt: 80, heightPt: 180 },
          physicalBounds: { xPt: 10, yPt: 10, widthPt: 80, heightPt: 180 },
        }],
        section: {} as SectionLayoutContext,
        ...canonicalPageMeta({} as SectionLayoutContext),
        layers: {
          paintSequence: [{ layer: 'body', node, coordinateSpace: 'section-logical' }],
          background: [],
          behindText: [],
          header: [],
          body: [node],
          notes: [],
          front: [],
          footer: [],
        },
        readingOrder: ['drawing-1'],
      }],
      diagnostics: [],
    };

    await expect(paintLayoutPage(layout, 0, target, { scale: 1, dpr: 1 })).resolves.toBeUndefined();
    expect(fills).toEqual([{ fill: '#ff0000', args: [10, 20, 30, 40] }]);
  });

  it('consumes the completed sequence without dereferencing page layer arrays', async () => {
    const context = {
      save() {}, restore() {}, setTransform() {}, clearRect() {}, fillRect() {},
      fillStyle: '',
    } as unknown as CanvasRenderingContext2D;
    const target = { width: 0, height: 0, getContext: () => context } as unknown as HTMLCanvasElement;
    const node = {
      kind: 'drawing' as const,
      id: 'drawing-1',
      source: { story: 'body' as const, storyInstance: 'body', path: [0] },
      flowBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
      inkBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
      advancePt: 40,
      ordinaryFlow: true,
      flowDomainId: 'body',
      commands: [],
    };
    const page = {
      pageIndex: 0,
      geometry: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 200, contentTopPt: 10, contentBottomPt: 190 },
      flowDomains: [{
        id: 'body', kind: 'body' as const,
        logicalBounds: { xPt: 10, yPt: 10, widthPt: 80, heightPt: 180 },
        physicalBounds: { xPt: 10, yPt: 10, widthPt: 80, heightPt: 180 },
      }],
      section: {} as SectionLayoutContext,
      ...canonicalPageMeta({} as SectionLayoutContext),
      layers: {
        paintSequence: [{
          layer: 'body' as const, node, coordinateSpace: 'section-logical' as const,
        }],
        background: [], behindText: [], header: [],
        get body(): never { throw new Error('paint dereferenced the body layer'); },
        notes: [], front: [], footer: [],
      },
      readingOrder: [node.id],
    };
    const layout: DocumentLayout = { pages: [page], diagnostics: [] };

    await expect(paintLayoutPage(layout, 0, target, { scale: 1, dpr: 1 }))
      .resolves.toBeUndefined();
  });

  it('dispatches retained tables through the canonical page painter', async () => {
    const fills: unknown[] = [];
    let currentFill = '';
    const context = {
      get fillStyle() { return currentFill; },
      set fillStyle(value: string | CanvasGradient | CanvasPattern) { currentFill = String(value); },
      save() {}, restore() {}, setTransform() {}, clearRect() {},
      beginPath() {}, rect() {}, clip() {}, translate() {}, rotate() {}, scale() {},
      fillRect(x: number, y: number, width: number, height: number) {
        fills.push([x, y, width, height, currentFill]);
      },
      strokeRect() {}, setLineDash() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {},
      strokeStyle: '', lineWidth: 1,
    } as unknown as CanvasRenderingContext2D;
    const target = { width: 0, height: 0, getContext: () => context } as unknown as HTMLCanvasElement;
    const bounds = { xPt: 10, yPt: 20, widthPt: 80, heightPt: 16 };
    const cell = {
      kind: 'table-cell', id: 'cell-0',
      source: { story: 'body', storyInstance: 'body', path: [0, 0, 0] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: bounds, inkBounds: bounds, contentBounds: bounds,
      advancePt: 16, verticalMerge: 'none', vAlign: 'top',
      background: { color: '#abcdef' }, blocks: [],
    };
    const row = {
      kind: 'table-row', id: 'row-0',
      source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: bounds, inkBounds: bounds, advancePt: 16, cells: [cell],
    };
    const table = {
      kind: 'table', id: 'table-0',
      source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: bounds, inkBounds: bounds, advancePt: 16,
      columnWidthsPt: [80], rows: [row], borders: [],
    };
    const layout = {
      pages: [{
        pageIndex: 0,
        geometry: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 200, contentTopPt: 10, contentBottomPt: 190 },
        flowDomains: [{
          id: 'body', kind: 'body',
          logicalBounds: { xPt: 10, yPt: 10, widthPt: 80, heightPt: 180 },
          physicalBounds: { xPt: 10, yPt: 10, widthPt: 80, heightPt: 180 },
        }],
        section: {} as SectionLayoutContext,
        ...canonicalPageMeta({} as SectionLayoutContext),
        layers: {
          paintSequence: [{ layer: 'body', node: table, coordinateSpace: 'section-logical' }],
          background: [], behindText: [], header: [], body: [table], notes: [], front: [], footer: [],
        },
        readingOrder: ['table-0'],
      }],
      diagnostics: [],
    } as unknown as DocumentLayout;

    await paintLayoutPage(layout, 0, target, { scale: 1, dpr: 1 });

    expect(fills).toContainEqual([10, 20, 80, 16, '#abcdef']);
  });
});

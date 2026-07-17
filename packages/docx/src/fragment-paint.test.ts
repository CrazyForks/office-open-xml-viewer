import { describe, it, expect, beforeAll } from 'vitest';
import { createLayoutServices, layoutDocument, renderDocumentToCanvas } from './renderer.js';
import { paintLayoutPage } from './paint/canvas-page.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  FramePr,
    SectionProps,
} from './types';
import type { TableFragmentLayout } from './layout/table-pagination.js';

// ─────────────────────────────────────────────────────────────────────────────
// PR 5 Task 13 — body fragment paint purity.
//
// A migrated body paragraph paints from its stored measured fragment
// through the renderer-owned retained adapter. At paint scale 1, stored point
// geometry needs no
// rescale, so the paint pass must draw the paragraph's lines WITHOUT calling
// measureText at all — no line layout, no segment measurement, no remeasurement.
//
// This is proved end-to-end through the real production flow: pages are paginated
// with a normal OffscreenCanvas metric, then painted at scale 1 onto a canvas whose
// measureText THROWS. If any part of the migrated paragraph paint tried to measure,
// the render would throw; instead it completes and draws the paragraph text.
// ─────────────────────────────────────────────────────────────────────────────

interface Call { text: string; x: number; y: number; }

/** Pagination-side canvas with a normal linear glyph metric. */
function makeMeasuringCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const per = p * 0.5;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeMeasuringCtx(); }
  };
});

/** Paint-side recording canvas whose measureText THROWS — any measurement during
 *  paint fails the test loudly. Records every text draw for the content assertion. */
function makeThrowingPaintCanvas(): { canvas: HTMLCanvasElement; calls: Call[]; measured: () => number } {
  let font = '10px serif';
  const calls: Call[] = [];
  let measured = 0;
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (_s: string) => {
      measured++;
      throw new Error('measureText must not be called during fragment paint');
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {}, setTransform() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, x: number, y: number) { calls.push({ text: s, x, y }); },
    strokeText(s: string, x: number, y: number) { calls.push({ text: s, x, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls, measured: () => measured };
}

function makeMeasuringPaintCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
  let font = '10px serif';
  const calls: Call[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const size = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * size * 0.5,
        fontBoundingBoxAscent: size * 0.8, fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8, actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {}, setTransform() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(text: string, x: number, y: number) { calls.push({ text, x, y }); },
    strokeText(text: string, x: number, y: number) { calls.push({ text, x, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

function para(text: string, over: Partial<DocParagraph> = {}): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
    ...over,
  } as unknown as DocParagraph;
}

function frame(over: Partial<FramePr> = {}): FramePr {
  return {
    dropCap: 'none', lines: 1, wrap: 'around',
    hAnchor: 'text', vAnchor: 'text', hRule: 'auto',
    hSpace: 0, vSpace: 0,
    ...over,
  };
}

function doc(body: BodyElement[], pageHeight = 400): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 200, pageHeight,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

// ---- Table builders (PR 6 Task 16) --------------------------------------------
function eb() {
  return { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
}
function tcell(content: CellElement[], over: Partial<DocTableCell> = {}): DocTableCell {
  return {
    content, colSpan: 1, vMerge: null, borders: eb(),
    background: null, vAlign: 'top', widthPt: null, ...over,
  } as unknown as DocTableCell;
}
function textCell(text: string, over: Partial<DocTableCell> = {}): DocTableCell {
  return tcell([{ type: 'paragraph', ...para(text) } as unknown as CellElement], over);
}
function trow(cells: DocTableCell[], over: Partial<DocTableRow> = {}): DocTableRow {
  return { cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false, ...over } as unknown as DocTableRow;
}
function tbl(rows: DocTableRow[], colWidths: number[], over: Partial<DocTable> = {}): DocTable {
  return {
    type: 'table', colWidths, rows, borders: eb(),
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed', ...over,
  } as unknown as DocTable;
}

describe('table fragment paint purity (PR 6 Task 16)', () => {
  it('paints a table with a NESTED table + vMerge at scale 1 without calling measureText', async () => {
    // Both outer and nested tables must be present in the retained layout tree;
    // paint traverses those nodes without reacquiring paragraph metrics.
    const inner = tbl([trow([textCell('inner one'), textCell('inner two')])], [40, 40]);
    const outer = tbl(
      [
        trow([textCell('a', { vMerge: true }), textCell('top')]),
        trow([textCell('', { vMerge: false }), tcell([{ type: 'table', ...inner } as unknown as CellElement])]),
      ],
      [60, 100],
    );
    const model = doc([outer as unknown as BodyElement]);
    const layout = layoutDocument(model);
    const table = layout.pages[0]?.layers.body[0];
    expect(table?.kind).toBe('table');
    if (table?.kind !== 'table') {
      throw new Error('expected retained TableLayout/TableFragmentLayout');
    }
    const nestedLayouts = table.rows.flatMap((tableRow) => tableRow.cells.flatMap((tableCell) =>
      tableCell.blocks.map((block) => block.layout).filter((layout) => layout.kind === 'table')));
    expect(nestedLayouts).toHaveLength(1);
    expect(nestedLayouts[0]!.rows.length).toBeGreaterThan(0);
    const paint = makeThrowingPaintCanvas();
    await expect(
      paintLayoutPage(layout, 0, paint.canvas, { dpr: 1, scale: 1 }),
    ).resolves.not.toThrow();
    expect(paint.measured()).toBe(0);
    // Non-vacuity: outer + inner cell text were drawn.
    expect(paint.calls.some((c) => c.text.includes('top'))).toBe(true);
    expect(paint.calls.some((c) => c.text.includes('inner'))).toBe(true);
  });

  it('paints a page-split table with a repeated header at scale 1, measure-free', async () => {
    const bodyRows = Array.from({ length: 12 }, (_v, i) => trow([textCell(`row ${i}`)]));
    const rows = [trow([textCell('HEADER')], { isHeader: true }), ...bodyRows];
    const model = doc([tbl(rows, [120]) as unknown as BodyElement], 120);
    const layout = layoutDocument(model);
    expect(layout.pages.length).toBeGreaterThan(1);
    const retained = layout.pages.map((page) => {
      const table = page.layers.body.find((node) => node.kind === 'table');
      if (table?.kind !== 'table') {
        throw new Error('expected retained TableFragmentLayout');
      }
      return table as TableFragmentLayout;
    });
    for (const fragment of retained.slice(1)) {
      expect(fragment.rows[0]?.ownership).toBe('repeated-header');
      expect(fragment.rows[0]?.logicalRowIndex).toBe(0);
    }
    for (let p = 0; p < layout.pages.length; p++) {
      const paint = makeThrowingPaintCanvas();
      await expect(
        paintLayoutPage(layout, p, paint.canvas, { dpr: 1, scale: 1 }),
      ).resolves.not.toThrow();
      expect(paint.measured()).toBe(0);
      expect(paint.calls.length).toBeGreaterThan(0);
    }
    // The header text is repeated on the continuation page.
    const paint2 = makeThrowingPaintCanvas();
    await paintLayoutPage(layout, 1, paint2.canvas, { dpr: 1, scale: 1 });
    expect(paint2.calls.some((c) => c.text.includes('HEADER'))).toBe(true);
  });
});

describe('fragment paint purity (PR 5 Task 13)', () => {
  it('keeps header/footer frame paragraphs on the B1 legacy story painter', async () => {
    const model = doc([para('body') as unknown as BodyElement]);
    model.headers.default = {
      body: [{
        type: 'paragraph',
        ...para('legacy header frame', { framePr: frame({ w: 80, hRule: 'auto' }) }),
      } as unknown as BodyElement],
    };
    const paint = makeMeasuringPaintCanvas();

    await renderDocumentToCanvas(model, paint.canvas, 0, { dpr: 1, width: 200 });

    expect(paint.calls.map((call) => call.text).join('')).toContain('legacy header frame');
  });

  it('paints a body text frame from retained geometry without measuring', async () => {
    const framed = para('retained frame text', {
      framePr: frame({ w: 50 }),
    });
    const model = doc([
      framed as unknown as BodyElement,
      para('anchor paragraph') as unknown as BodyElement,
    ]);
    const layout = layoutDocument(model);
    const placed = layout.pages[0]?.layers.body[0];
    if (placed?.kind !== 'paragraph') throw new Error('expected frame paragraph layout');
    expect(placed.advancePt).toBe(0);
    expect(placed.ordinaryFlow).toBe(false);
    expect(placed.lines.length).toBeGreaterThan(0);
    expect(placed.lines.flatMap((line) => line.placements)).not.toHaveLength(0);

    const retainedLines = placed.lines;
    const partition = retainedLines.map((line) => line.range);
    for (const width of [200, 400]) {
      const paint = makeThrowingPaintCanvas();
      await expect(
        paintLayoutPage(layout, 0, paint.canvas, { dpr: 1, scale: width / 200 }),
      ).resolves.not.toThrow();
      expect(paint.measured()).toBe(0);
      expect(paint.calls.map((call) => call.text).join('')).toContain('retained frame');
      expect(placed.lines).toBe(retainedLines);
      expect(placed.lines.map((line) => line.range)).toEqual(partition);
    }
  });

  it('prepares frame metadata even when callers provide custom layout services', () => {
    const model = doc([
      para('custom services frame', { framePr: frame({ w: 50 }) }) as unknown as BodyElement,
      para('anchor') as unknown as BodyElement,
    ]);
    const base = createLayoutServices(model);
    const custom = { text: base.text, images: base.images, math: base.math };
    const placed = layoutDocument(model, custom, { currentDateMs: 0 }).pages[0]?.layers.body[0];

    expect(placed?.kind).toBe('paragraph');
    if (placed?.kind !== 'paragraph') throw new Error('expected frame paragraph layout');
    expect(placed.ordinaryFlow).toBe(false);
  });

  it('retains identical adjacent framePr paragraphs as one stacked frame and one exclusion', () => {
    const shared = frame({ w: 50 });
    const model = doc([
      para('frame first', { framePr: { ...shared } }) as unknown as BodyElement,
      para('frame second', { framePr: { ...shared } }) as unknown as BodyElement,
      para('anchor') as unknown as BodyElement,
    ]);
    const nodes = layoutDocument(model).pages[0]?.layers.body ?? [];
    const [first, second, anchor] = nodes;
    if (first?.kind !== 'paragraph' || second?.kind !== 'paragraph') {
      throw new Error('expected frame paragraph layouts');
    }
    expect(first.advancePt).toBe(0);
    expect(second.advancePt).toBe(0);
    expect(second.flowBounds.yPt)
      .toBeGreaterThanOrEqual(first.flowBounds.yPt + first.flowBounds.heightPt);
    if (anchor?.kind !== 'paragraph') throw new Error('expected anchor paragraph layout');
    expect(anchor.exclusions).toHaveLength(1);
  });

  it('uses final-width reflow to determine an automatic frame height', () => {
    const model = doc([
      para('abcdefghij', { framePr: frame({ w: 20, hRule: 'auto' }) }) as unknown as BodyElement,
      para('anchor') as unknown as BodyElement,
    ]);
    const [framed, anchor] = layoutDocument(model).pages[0]?.layers.body ?? [];
    if (framed?.kind !== 'paragraph' || anchor?.kind !== 'paragraph') throw new Error('expected paragraph layouts');
    expect(framed.lines.length).toBeGreaterThan(1);
    expect(framed.advancePt).toBe(0);
    expect(anchor.exclusions[0]?.bounds.heightPt).toBeCloseTo(framed.flowBounds.heightPt, 6);
  });

  it('uses the larger of authored and final-content height for hRule=atLeast', () => {
    const contentDriven = doc([
      para('abcdefghij', { framePr: frame({ w: 20, hRule: 'atLeast', h: 5 }) }) as unknown as BodyElement,
      para('anchor one') as unknown as BodyElement,
    ]);
    const authoredDriven = doc([
      para('x', { framePr: frame({ w: 20, hRule: 'atLeast', h: 80 }) }) as unknown as BodyElement,
      para('anchor two') as unknown as BodyElement,
    ]);
    const [contentFrame, contentAnchor] = layoutDocument(contentDriven).pages[0]?.layers.body ?? [];
    const authoredAnchor = layoutDocument(authoredDriven).pages[0]?.layers.body[1];
    if (contentFrame?.kind !== 'paragraph'
      || contentAnchor?.kind !== 'paragraph'
      || authoredAnchor?.kind !== 'paragraph') throw new Error('expected paragraph layouts');

    expect(contentFrame.advancePt).toBe(0);
    expect(contentAnchor.exclusions[0]?.bounds.heightPt).toBeCloseTo(contentFrame.flowBounds.heightPt, 6);
    expect(contentAnchor.exclusions[0]?.bounds.heightPt).toBeGreaterThan(5);
    expect(authoredAnchor.exclusions[0]?.bounds.heightPt).toBeCloseTo(80, 6);
  });

  it('retains one authored outer clip on every member of an hRule=exact frame group', () => {
    const shared = frame({ w: 30, hRule: 'exact', h: 25 });
    const model = doc([
      para('first frame member', { framePr: { ...shared } }) as unknown as BodyElement,
      para('second frame member', { framePr: { ...shared } }) as unknown as BodyElement,
      para('anchor') as unknown as BodyElement,
    ]);
    const [first, second, anchor] = layoutDocument(model).pages[0]?.layers.body ?? [];
    if (first?.kind !== 'paragraph'
      || second?.kind !== 'paragraph'
      || anchor?.kind !== 'paragraph') throw new Error('expected paragraph layouts');

    expect(first.clipBounds).toEqual(second.clipBounds);
    expect(first.clipBounds).toEqual({
      xPt: first.clipBounds?.xPt,
      yPt: first.clipBounds?.yPt,
      widthPt: 30,
      heightPt: 25,
    });
    expect(anchor.exclusions[0]?.bounds.heightPt).toBeCloseTo(25, 6);
  });

  it('folds contextual spacing once across a three-paragraph frame group', () => {
    const shared = frame({ w: 50 });
    const model = doc([
      para('one', { framePr: { ...shared }, styleId: 's', spaceBefore: 2, spaceAfter: 6 }) as unknown as BodyElement,
      para('two', { framePr: { ...shared }, styleId: 's', contextualSpacing: true, spaceBefore: 4, spaceAfter: 5 }) as unknown as BodyElement,
      para('three', { framePr: { ...shared }, styleId: 's', spaceBefore: 8, spaceAfter: 3 }) as unknown as BodyElement,
      para('anchor') as unknown as BodyElement,
    ]);
    const [first, second, third, anchor] = layoutDocument(model).pages[0]?.layers.body ?? [];
    if (
      first?.kind !== 'paragraph'
      || second?.kind !== 'paragraph'
      || third?.kind !== 'paragraph'
    ) throw new Error('expected frame paragraph layouts');
    expect(first.advancePt).toBe(0);
    expect(second.advancePt).toBe(0);
    expect(second.flowBounds.yPt - first.flowBounds.yPt).toBeCloseTo(first.flowBounds.heightPt, 6);
    // p2 after=5 is replaced by the contextual gap 3 before p3.
    expect(third.flowBounds.yPt - second.flowBounds.yPt).toBeCloseTo(second.flowBounds.heightPt - 2, 6);
    if (anchor?.kind !== 'paragraph') throw new Error('expected anchor paragraph layout');
    expect(anchor.exclusions[0]?.bounds.heightPt).toBeCloseTo(
      third.flowBounds.yPt + third.flowBounds.heightPt - (first.flowBounds.yPt - 2),
      6,
    );
  });

  it('merges paragraph borders inside one frame and reserves the final bottom edge once', () => {
    const edge = { style: 'single', color: '000000', width: 2, space: 3 };
    const borders = { top: edge, right: edge, bottom: edge, left: edge, between: null };
    const shared = frame({ w: 50 });
    const model = doc([
      para('one', { framePr: { ...shared }, borders, spaceBefore: 0, spaceAfter: 0 }) as unknown as BodyElement,
      para('two', { framePr: { ...shared }, borders, spaceBefore: 0, spaceAfter: 0 }) as unknown as BodyElement,
      para('anchor') as unknown as BodyElement,
    ]);
    const [first, second, anchor] = layoutDocument(model).pages[0]?.layers.body ?? [];
    if (first?.kind !== 'paragraph' || second?.kind !== 'paragraph') {
      throw new Error('expected frame paragraph layouts');
    }
    expect(first.borders.some((border) => border.edge === 'bottom')).toBe(false);
    expect(second.borders.some((border) => border.edge === 'bottom')).toBe(true);
    expect(first.advancePt).toBe(0);
    expect(second.advancePt).toBe(0);
    expect(second.flowBounds.yPt).toBeCloseTo(first.flowBounds.yPt + first.flowBounds.heightPt, 6);
    if (anchor?.kind !== 'paragraph') throw new Error('expected anchor paragraph layout');
    expect(anchor.exclusions[0]?.bounds.heightPt).toBeCloseTo(
      second.flowBounds.yPt + second.flowBounds.heightPt - first.flowBounds.yPt,
      6,
    );
  });

  it('paints a premeasured body paragraph at scale 1 without ever calling measureText', async () => {
    const model = doc([para('hello world one two three') as unknown as BodyElement]);
    const layout = layoutDocument(model); // measured with the normal OffscreenCanvas
    const paint = makeThrowingPaintCanvas();

    // Paint scale 1 (render width == page width). A migrated paragraph draws its
    // stored fragment lines; nothing measures.
    await expect(
      paintLayoutPage(layout, 0, paint.canvas, { dpr: 1, scale: 1 }),
    ).resolves.not.toThrow();

    expect(paint.measured()).toBe(0);
    // Non-vacuity: the paragraph's words were actually drawn.
    const drewText = paint.calls.some((c) => c.text.includes('hello'));
    expect(drewText).toBe(true);
  });

  it('paints a paragraph that SPLITS across pages from fragments, still measure-free', async () => {
    // A long paragraph over a short page splits; each continuation slice paints from
    // the shared measured fragment window without remeasuring.
    const long = Array.from({ length: 120 }, () => 'w').join(' ');
    const model = doc([para(long) as unknown as BodyElement], 60);
    const layout = layoutDocument(model);
    expect(layout.pages.length).toBeGreaterThan(1);
    const split = layout.pages.some((page) => page.layers.body.some((node) => (
      node.kind === 'paragraph' && node.continuation?.continuesOnNext
    )));
    expect(split).toBe(true);

    for (let p = 0; p < layout.pages.length; p++) {
      const paint = makeThrowingPaintCanvas();
      await expect(
        paintLayoutPage(layout, p, paint.canvas, { dpr: 1, scale: 1 }),
      ).resolves.not.toThrow();
      expect(paint.measured()).toBe(0);
      expect(paint.calls.length).toBeGreaterThan(0);
    }
  });
});

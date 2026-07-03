import { describe, it, expect } from 'vitest';
import {
  splitTableAcrossPages,
  __test_computeTableLayout,
  __test_setTableReuseEnabled,
} from './renderer.js';
import type { RenderState } from './renderer.js';
import type {
  DocTable,
  DocTableRow,
  DocTableCell,
  DocParagraph,
  PaginatedBodyElement,
} from './types';

// B2 table stage 1b — compute-once TABLE layout reuse (stamped column widths +
// row heights). These characterization tests pin:
//   1. splitTableAcrossPages stamps each SLICE with its own rows' heights
//      (repeated tblHeader rows prepended on continuations) + the shared column
//      widths + the scale-1 contentWPt gate input.
//   2. computeTableLayout with the reuse ON returns the SAME layout as with reuse
//      OFF (a fresh resolveColumnWidths/resolveTableRowHeights recompute) — the
//      reuse is a pure optimization.
//   3. the reuse returns EXACTLY `stamp × scale` (proving the gate fires, not that
//      it silently fell through to recompute).
//   4. a stamp whose `contentWPt` no longer matches the paint band falls back to
//      recompute (the self-verifying gate rejects a stale stamp).

function emptyBorders() {
  return { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
}

// ---- Minimal recording 2D context (mirrors table-clip-exact.test) --------------
function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {}, strokeRect() {},
    rect() {}, clip() {}, scale() {}, translate() {},
    setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  (ctx as unknown as { canvas: unknown }).canvas = { width: 2000, height: 2000 };
  return ctx as unknown as CanvasRenderingContext2D;
}

function makeState(scale: number): RenderState {
  return {
    ctx: makeCtx(),
    scale,
    dpr: 1,
    contentX: 0,
    contentW: 500 * scale,
    y: 0,
    pageH: 1000,
    defaultColor: '#000000',
    pageIndex: 0,
    totalPages: 1,
    images: new Map(),
    dryRun: true,
    marginLeft: 0,
    marginRight: 0,
    marginTop: 0,
    marginBottom: 0,
    floats: [],
    floatParaSeq: 0,
    docGrid: { type: null, linePitchPt: null, charSpacePt: null },
    docEastAsian: false,
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    kinsoku: { enabled: false, lineStartForbidden: new Set<number>(), lineEndForbidden: new Set<number>() },
    defaultTabPt: 36,
    showTrackChanges: false,
  } as unknown as RenderState;
}

function para(text: string): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text === '' ? [] : [{
      type: 'text', text,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
}

function cell(text: string, widthPt: number): DocTableCell {
  return {
    content: [{ type: 'paragraph', ...para(text) }],
    colSpan: 1,
    vMerge: null,
    borders: emptyBorders(),
    background: null,
    vAlign: 'top',
    widthPt,
  } as unknown as DocTableCell;
}

function twoColRow(a: string, b: string): DocTableRow {
  return {
    cells: [cell(a, 200), cell(b, 300)],
    rowHeight: null,
    rowHeightRule: 'auto',
    isHeader: false,
  } as unknown as DocTableRow;
}

function twoColTable(rows: DocTableRow[]): DocTable {
  return {
    colWidths: [200, 300], // pt grid (sums to the 500pt content band)
    rows,
    borders: emptyBorders(),
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
    layout: 'fixed', // trust the grid so recompute is deterministic
  } as unknown as DocTable;
}

describe('splitTableAcrossPages — table stamp (B2 table stage 1b)', () => {
  const rowsOf = (slice: PaginatedBodyElement) => (slice as unknown as DocTable).rows;

  it('stamps each slice with its own rows heights + shared column widths + contentWPt', () => {
    const t = twoColTable(Array.from({ length: 6 }, (_, i) => twoColRow(`a${i}`, `b${i}`)));
    const rowHs = [10, 20, 30, 40, 50, 60];
    const colWidthsPt = [200, 300];
    const pages: PaginatedBodyElement[][] = [[]];
    const newPage = () => { pages.push([]); };
    // contentH = 65 ⇒ rows of 10,20,30 = 60 fit on page 1; 40 alone on page 2 (40),
    // then 50, then 60 — one row per later page (each > 65-startY residual).
    splitTableAcrossPages(
      t, rowHs, 0, 65, pages, newPage,
      undefined, undefined, undefined, undefined, undefined, undefined,
      { colWidthsPt, contentWPt: 500 },
    );

    // Every slice carries the stamp.
    const slices = pages.map((p) => p[0]).filter(Boolean);
    expect(slices.length).toBeGreaterThan(1); // it really split
    for (const s of slices) {
      expect(s.tableColWidthsPt).toEqual(colWidthsPt);
      expect(s.tableLayoutInputs).toEqual({ scale: 1, contentWPt: 500 });
      // Row-height stamp aligns 1:1 with the slice's rows.
      expect(s.tableRowHeightsPt!.length).toBe(rowsOf(s).length);
    }

    // The stamped heights, concatenated across slices, are the original per-row
    // heights in order (no header repetition here — no isHeader rows).
    const flat = slices.flatMap((s) => s.tableRowHeightsPt as number[]);
    expect(flat).toEqual(rowHs);
  });

  it('prepends the repeated tblHeader rows heights on continuation slices', () => {
    // Row 0 is a header; it repeats at the top of every continuation slice, so its
    // height must be prepended to the continuation stamp (aligning with sliceRows).
    const header: DocTableRow = { ...twoColRow('H0', 'H1'), isHeader: true } as unknown as DocTableRow;
    const body = Array.from({ length: 4 }, (_, i) => twoColRow(`a${i}`, `b${i}`));
    const t = twoColTable([header, ...body]);
    const rowHs = [15, 20, 20, 20, 20]; // header=15, body rows=20 each
    const pages: PaginatedBodyElement[][] = [[]];
    const newPage = () => { pages.push([]); };
    // contentH = 60: page1 = header(15) + 2 body(40) = 55 fit; continuations repeat
    // header(15) + as many body as fit in 60-15=45 ⇒ 2 body (40).
    splitTableAcrossPages(
      t, rowHs, 0, 60, pages, newPage,
      undefined, undefined, undefined, undefined, undefined, undefined,
      { colWidthsPt: [200, 300], contentWPt: 500 },
    );
    const slices = pages.map((p) => p[0]).filter(Boolean);
    expect(slices.length).toBeGreaterThan(1);
    // Page 1 slice: header + body rows, stamp = [15, 20, ...].
    expect(slices[0].tableRowHeightsPt![0]).toBe(15);
    // Continuation slice(s): first stamped height is the REPEATED header height.
    for (const s of slices.slice(1)) {
      expect((s as unknown as DocTable).rows[0].isHeader).toBe(true);
      expect(s.tableRowHeightsPt![0]).toBe(15); // header height prepended
      // Length aligns with sliceRows (header + body).
      expect(s.tableRowHeightsPt!.length).toBe((s as unknown as DocTable).rows.length);
    }
  });

  it('omits the stamp when no tableStamp payload is passed (direct callers)', () => {
    const t = twoColTable(Array.from({ length: 4 }, (_, i) => twoColRow(`a${i}`, `b${i}`)));
    const pages: PaginatedBodyElement[][] = [[]];
    const newPage = () => { pages.push([]); };
    splitTableAcrossPages(t, [30, 30, 30, 30], 0, 65, pages, newPage);
    for (const p of pages) {
      const s = p[0];
      if (!s) continue;
      expect(s.tableColWidthsPt).toBeUndefined();
      expect(s.tableRowHeightsPt).toBeUndefined();
      expect(s.tableLayoutInputs).toBeUndefined();
    }
  });
});

describe('computeTableLayout — stamp reuse (B2 table stage 1b)', () => {
  function stampedTable(): DocTable {
    // Build the table, resolve its true layout once at scale 1 (reuse OFF) and
    // stamp THAT as the paginator would, so the stamp is self-consistent.
    const t = twoColTable([twoColRow('alpha', 'beta gamma'), twoColRow('x', 'y')]);
    const prev = __test_setTableReuseEnabled(false);
    const truth = __test_computeTableLayout(t, 500, makeState(1));
    __test_setTableReuseEnabled(prev);
    const el = t as PaginatedBodyElement;
    el.tableColWidthsPt = truth.colWidths; // scale-1 ⇒ pt
    el.tableRowHeightsPt = truth.rowHeights; // scale-1 ⇒ pt
    el.tableLayoutInputs = { scale: 1, contentWPt: 500 };
    return t;
  }

  it('reuse ON equals reuse OFF (pure optimization) at a non-unit scale', () => {
    const t = stampedTable();
    const scale = 1.5;
    const contentWPx = 500 * scale;

    const prev = __test_setTableReuseEnabled(false);
    const off = __test_computeTableLayout(t, contentWPx, makeState(scale));
    __test_setTableReuseEnabled(prev);

    const on = __test_computeTableLayout(t, contentWPx, makeState(scale));

    expect(on.colWidths).toEqual(off.colWidths);
    expect(on.rowHeights).toEqual(off.rowHeights);
    expect(on.tableW).toBe(off.tableW);
  });

  it('reuse returns exactly stamp × scale (the gate fires, not a silent fallthrough)', () => {
    const t = stampedTable();
    const scale = 2;
    const on = __test_computeTableLayout(t, 500 * scale, makeState(scale));
    const el = t as PaginatedBodyElement;
    expect(on.colWidths).toEqual(el.tableColWidthsPt!.map((w) => w * scale));
    expect(on.rowHeights).toEqual(el.tableRowHeightsPt!.map((h) => h * scale));
  });

  it('rejects a stamp whose contentWPt no longer matches the paint band (falls back to recompute)', () => {
    const t = stampedTable();
    const el = t as PaginatedBodyElement;
    // Corrupt the stamp so it cannot be reused: pretend it was fit to a much
    // narrower band. The gate must reject it and recompute against the real band.
    el.tableRowHeightsPt = el.tableRowHeightsPt!.map(() => 9999);
    el.tableColWidthsPt = [1, 1];
    el.tableLayoutInputs = { scale: 1, contentWPt: 123 };

    const scale = 1;
    const got = __test_computeTableLayout(t, 500 * scale, makeState(scale));

    // A reuse would have returned the corrupt stamp (colWidths [1,1], rows 9999);
    // the recompute returns the real fixed-grid columns (200,300 at scale 1).
    expect(got.colWidths).not.toEqual([1, 1]);
    expect(got.colWidths[0]).toBeCloseTo(200, 6);
    expect(got.colWidths[1]).toBeCloseTo(300, 6);
    expect(got.rowHeights.every((h) => h < 9999)).toBe(true);
  });

  it('reuse OFF (master switch) always recomputes even with a valid stamp', () => {
    const t = stampedTable();
    const el = t as PaginatedBodyElement;
    // Poison the column stamp; with reuse OFF the recompute ignores it.
    el.tableColWidthsPt = [1, 1];
    const prev = __test_setTableReuseEnabled(false);
    const got = __test_computeTableLayout(t, 500, makeState(1));
    __test_setTableReuseEnabled(prev);
    expect(got.colWidths[0]).toBeCloseTo(200, 6);
    expect(got.colWidths[1]).toBeCloseTo(300, 6);
  });
});

import { describe, it, expect } from 'vitest';
import { computePages } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocTable,
  DocTableRow,
  DocxTextRun,
  SectionProps,
  TblpPr,
  PaginatedBodyElement,
} from './types';

// Unit tests for the keep-with-anchor pagination of a page-overflowing FLOATING
// table (ECMA-376 §17.4.57 `<w:tblpPr>`). §17.4.57 pins only the table's
// size/position; keeping an undivided floating table with its anchor context on
// the next page is Word runtime behaviour — the SAME "keep on page" semantics
// already covered for a text frame (§17.3.1.11, frame-keep-with-anchor.test.ts)
// and a paragraph-anchored image float. These assertions guard that behaviour,
// which the private-sample VRT (sample-11's small top-of-page float only, never
// a page-boundary one) cannot cover.
//
// The stub canvas mirrors frame-keep-with-anchor.test.ts / pagination.test.ts:
// glyph advance = charCount × fontPx and the font box = 0.8/0.2 em, so a single
// line is exactly fontPx tall. Table row heights are pinned with
// rowHeightRule="exact" so `tableH` is deterministic (independent of the stub's
// cell measurement), the table analogue of the frame's hRule="exact"/h.

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
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
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

// pageWidth 200 / pageHeight 140, margins 20 ⇒ content band 160×100 (bodyTop 20).
function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 140,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

type DocRun = DocParagraph['runs'][number];

function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocRun;
}

function para(opts: { text?: string; fontSize?: number } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: opts.text ? [textRun(opts.text, fontSize)] : [],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

// Full TblpPr with the spec defaults; callers override only the axis under test.
// The default vertAnchor is 'page' (matching the Rust parser's fill for an
// absent w:vertAnchor), so the vAnchor gate is exercised by overriding it.
function tblp(over: Partial<TblpPr> = {}): TblpPr {
  return {
    leftFromText: 0,
    rightFromText: 0,
    topFromText: 0,
    bottomFromText: 0,
    horzAnchor: 'text',
    horzSpecified: true,
    vertAnchor: 'page',
    tblpX: 0,
    tblpY: 0,
    ...over,
  };
}

/** A single-cell row of the given exact pt height (`rowHeightRule="exact"` short-
 *  circuits content measurement, so `tableH` == `rowHeight` regardless of the
 *  stub canvas — the table analogue of the frame's hRule="exact"/h). */
function row(heightPt: number): DocTableRow {
  return {
    cells: [
      {
        content: [],
        colSpan: 1,
        vMerge: null,
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        background: null,
        vAlign: 'top',
        widthPt: null,
      },
    ],
    rowHeight: heightPt,
    rowHeightRule: 'exact',
    isHeader: false,
  };
}

/** A floating table (`w:tblpPr`) of total height `tableHPt`, laid out as one
 *  exact-height row so its measured extent is deterministic. */
function floatTable(tp: TblpPr, tableHPt: number): BodyElement {
  const t: DocTable = {
    colWidths: [80],
    rows: [row(tableHPt)],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
    tblpPr: tp,
  };
  return { type: 'table', ...t } as unknown as BodyElement;
}

/** True when an element is a floating table (identified by its tblpPr). */
const isFloatTable = (el: PaginatedBodyElement): boolean =>
  el.type === 'table' && (el as unknown as DocTable).tblpPr != null;

/** True when a page holds a floating table. */
const hasFloatTable = (page: PaginatedBodyElement[]): boolean => page.some(isFloatTable);

/** Text of a paragraph element (joins its text runs). */
const textOf = (el: PaginatedBodyElement): string =>
  el.type === 'paragraph'
    ? (el as unknown as DocParagraph).runs
        .filter((r) => r.type === 'text')
        .map((r) => (r as DocxTextRun).text)
        .join('')
    : '';

/** True when a page holds the anchor paragraph (matched by its text). */
const hasAnchorText = (page: PaginatedBodyElement[], text: string): boolean =>
  page.some((el) => textOf(el) === text);

/** The newspaper column an element landed in. */
const colOf = (el: PaginatedBodyElement): number | undefined => el.colIndex;

/** Find the (first) floating-table element on a page. */
const floatTableEl = (page: PaginatedBodyElement[]): PaginatedBodyElement | undefined =>
  page.find(isFloatTable);

describe('computePages — floating-table page-fit / keep-with-anchor (§17.4.57)', () => {
  // Content band 160×100, bodyTop 20. Table: vertAnchor="text", one exact row of
  // height H ⇒ table body box [paraTop, paraTop+H]. With N leading 20pt lines the
  // table's in-flow top is at y=20N; the table overflows the 100pt content area
  // once 20N + H > 100.

  it('relocates a text-anchored floating table + its anchor text to the next page when it overflows the bottom', () => {
    // 3 leading lines (y advances 20→40→60), then the floating table (60 tall:
    // 60+60 > 100 ⇒ overflow), then the anchor text paragraph.
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      floatTable(tblp({ vertAnchor: 'text', tblpY: 0 }), 60),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    // The table does NOT stay on page 1 (would overflow); it moves to page 2…
    expect(hasFloatTable(pages[0])).toBe(false);
    expect(hasFloatTable(pages[1])).toBe(true);
    // …and its trailing anchor text follows it onto page 2 (kept together).
    expect(hasAnchorText(pages[1], 'anchor')).toBe(true);
    expect(hasAnchorText(pages[0], 'anchor')).toBe(false);
    // Page 1 keeps only the three leading lines (no float band leaked over).
    expect(pages[0].map(textOf)).toEqual(['a', 'b', 'c']);
  });

  it('relocates an overflowing floating table to the NEXT COLUMN (not a new page) in a multi-column section', () => {
    // 2 equal columns: colW = (160-20)/2 = 70; content height 100 (5 × 20pt per
    // column). Column 0 gets 3 leading lines (y 20→40→60); the table body box
    // (60 tall) then overflows column 0's bottom (60+60 > 100). With a column
    // still available on the page, it relocates to column 1 — NOT a new page.
    const twoCol = section({ columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] } });
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      floatTable(tblp({ vertAnchor: 'text', horzAnchor: 'text', tblpY: 0 }), 60),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, twoCol, makeCtx());
    // Still a single page (column 1 absorbed the table + anchor).
    expect(pages.length).toBe(1);
    const f = floatTableEl(pages[0]);
    expect(f).toBeDefined();
    expect(colOf(f as PaginatedBodyElement)).toBe(1); // moved into column 1
    // The three leading lines (a/b/c) stayed in column 0.
    const leadCols = pages[0]
      .filter((el) => ['a', 'b', 'c'].includes(textOf(el)))
      .map(colOf);
    expect(leadCols).toEqual([0, 0, 0]);
    // The trailing anchor text follows the table into column 1.
    const anchor = pages[0].find((el) => textOf(el) === 'anchor');
    expect(anchor).toBeDefined();
    expect(colOf(anchor as PaginatedBodyElement)).toBe(1);
  });

  it('keeps a text-anchored floating table in place when it fits (near the top of the page)', () => {
    // Only 1 leading line (y=20). Table body box [20,80] fits within [0,100] ⇒
    // no relocation. Everything stays on page 1. (sample-11 shape: a small
    // vertAnchor="text" tblpY=1 float near the page top must NOT be sent.)
    const body = [
      para({ text: 'a' }),
      floatTable(tblp({ vertAnchor: 'text', tblpY: 1 }), 60),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(hasFloatTable(pages[0])).toBe(true);
    expect(hasAnchorText(pages[0], 'anchor')).toBe(true);
  });

  it('does NOT relocate (or loop) a floating table taller than the page content area', () => {
    // 150 tall > content height 100: the table can never fit on any page, so
    // relocating would loop forever. It is left in place and allowed to overflow.
    // The real assertion is that this terminates (no timeout / infinite paging)
    // AND that the floating table is NOT row-split (Word keeps it undivided).
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      floatTable(tblp({ vertAnchor: 'text', tblpY: 0 }), 150),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    // Table stays on page 1 with its three leading lines (no relocation).
    expect(hasFloatTable(pages[0])).toBe(true);
    // A floating table adds no flow height, so the anchor stays on the same page.
    expect(hasAnchorText(pages[0], 'anchor')).toBe(true);
    expect(pages.length).toBe(1);
    // Exactly ONE floating-table element exists across all pages (not split into
    // per-page slices like a block table).
    const floatCount = pages.reduce((s, p) => s + p.filter(isFloatTable).length, 0);
    expect(floatCount).toBe(1);
  });

  it('does NOT relocate a page-anchored floating table that overflows (absolute y is honored in place)', () => {
    // vertAnchor="page", tblpY=90: the table is pinned at page-y 90 with H=60 ⇒
    // bottom 150, past the 140 page edge. An absolute page position is the SAME
    // on any page, so relocation cannot help — Word draws it there (overflowing).
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      floatTable(tblp({ vertAnchor: 'page', tblpY: 90 }), 60),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(hasFloatTable(pages[0])).toBe(true);
    expect(hasAnchorText(pages[0], 'anchor')).toBe(true);
  });

  it('does NOT relocate a margin-anchored floating table that overflows (absolute y is honored in place)', () => {
    // vertAnchor="margin", tblpY=70: table at margin-top(20)+70 = 90, H=60 ⇒
    // bottom 150, past the bottom margin. Absolute ⇒ left in place (mirrors
    // vertAnchor="page").
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      floatTable(tblp({ vertAnchor: 'margin', tblpY: 70 }), 60),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(hasFloatTable(pages[0])).toBe(true);
  });
});

import { describe, it, expect } from 'vitest';
import {
  layoutBidiTabStops,
  nextTabStopRtl,
  type BidiTabItem,
} from './line-layout.js';
import { computeLineVisualOrder } from './bidi-line.js';
import { renderDocumentToCanvas } from './renderer.js';
import type { DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// ECMA-376 §17.3.1.6 (base RTL) + §17.3.1.37 (tabs) + §17.15.1.25 (default tab)
// + §17.18.84 (ST_TabJc start/end logical edges) — a bidi paragraph's tab stops
// are anchored at the LEADING (right) text edge and its cells reorder visually.
// Issue #820: the Arabic TOC's page numbers and dot/underscore leaders (and the
// footer's tab-aligned fields) landed on the wrong visual side (or wrapped to a
// new line) because tabs were resolved in LTR pen coordinates. These tests pin
// the mirrored resolution independently of any font metrics.

describe('nextTabStopRtl (§17.3.1.37 / §17.15.1.25, leading = right edge)', () => {
  const stops = [
    { pos: 100, alignment: 'left' as const, leader: 'none' as const },
    { pos: 300, alignment: 'right' as const, leader: 'underscore' as const },
  ];
  it('advances leftward to the nearest custom stop past the pen', () => {
    // Pen 50 from the right edge → next stop further left is pos 100.
    expect(nextTabStopRtl(50, stops, 36)?.pos).toBe(100);
    // Pen 100 (on the first stop) → advances to 300.
    expect(nextTabStopRtl(100, stops, 36)?.pos).toBe(300);
  });
  it('falls onto the §17.15.1.25 automatic grid AFTER all custom stops', () => {
    // Past the last custom stop (300); the grid is anchored at the leading edge
    // with interval 36, so the next multiple past 300 is 324.
    const s = nextTabStopRtl(300, stops, 36);
    expect(s?.pos).toBe(324);
    expect(s?.alignment).toBe('left');
    expect(s?.leader).toBeUndefined();
  });
  it('returns null when no interval and the pen is past every custom stop', () => {
    expect(nextTabStopRtl(400, stops, 0)).toBeNull();
  });
});

describe('computeLineVisualOrder treats a tab as a segment separator (UAX#9 S)', () => {
  it('reorders tab-delimited cells in mirrored order under an RTL base', () => {
    // Logical order: [chapNum][space][title] TAB [pageNum] (all rtl-marked).
    const segs = [
      { text: '1.1', rtl: true }, { text: ' ', rtl: true }, { text: 'TITLE', rtl: true },
      { isTab: true },
      { text: '4', rtl: true },
    ];
    const { order } = computeLineVisualOrder(segs as unknown[], true);
    // Visual L→R must be: pageNum, TAB, title, space, chapNum — the page number
    // ends up on the visual LEFT and the chapter number on the visual RIGHT, with
    // the tab (leader region) between the cells. Without the S classification the
    // whole line reversed as one run and the page number landed mid-line.
    expect(order.map((i) => (('isTab' in segs[i]) ? 'TAB' : (segs[i] as { text: string }).text)))
      .toEqual(['4', 'TAB', 'TITLE', ' ', '1.1']);
  });
});

describe('layoutBidiTabStops (§17.3.1.37 mirror — reading-frame layout)', () => {
  const avail = 400; // content width px (0 = left edge, 400 = leading/right edge)

  it('places an end (right) tab so its page number trails at the mirrored stop', () => {
    // TOC row logical order: chapNum(20) space(4) title(80) TAB pageNum(8).
    // The tab is the style right/underscore leader stop at pos 300 from the right
    // edge. The page number is trailing-aligned: its LEFT edge sits on the stop
    // (visual x = avail-300 = 100).
    const items: BidiTabItem[] = [
      { isTab: false, width: 20 }, { isTab: false, width: 4 }, { isTab: false, width: 80 },
      { isTab: true, width: 0 },
      { isTab: false, width: 8 },
    ];
    const stops = [{ pos: 300, alignment: 'right' as const, leader: 'underscore' as const }];
    const res = layoutBidiTabStops(items, stops, avail, 36);
    expect(res[4].visualX).toBeCloseTo(100, 1);
    // Chapter number (first logical, leading) sits at the RIGHT edge: its right
    // edge = avail.
    expect(res[0].visualX + 20).toBeCloseTo(avail, 1);
    // The tab carries the underscore leader and fills the visible gap.
    expect(res[3].leader).toBe('underscore');
    expect(res[3].width).toBeGreaterThan(0);
  });

  it('pins a page number to the left margin when the stop is past it', () => {
    // A right/leader stop at pos 420 (past avail=400) would place the page
    // number's left edge left of the margin; it pins so its far (left) edge is on
    // the margin (visual x 0), the page number spanning [0, 8].
    const items: BidiTabItem[] = [
      { isTab: false, width: 20 }, { isTab: false, width: 80 },
      { isTab: true, width: 0 },
      { isTab: false, width: 8 },
    ];
    const stops = [{ pos: 420, alignment: 'right' as const, leader: 'underscore' as const }];
    const res = layoutBidiTabStops(items, stops, avail, 36);
    expect(res[3].visualX).toBeCloseTo(0, 1);
  });

  it('flips physical left (leading) so its content ends at the mirrored stop', () => {
    // Single leading (left) tab at pos 150 from the right edge → visual x
    // avail-150 = 250. Following content (width 30) has its LEADING (right) edge
    // there, so it spans [220, 250].
    const items: BidiTabItem[] = [
      { isTab: false, width: 40 }, // leading content
      { isTab: true, width: 0 },
      { isTab: false, width: 30 }, // follows the tab
    ];
    const stops = [{ pos: 150, alignment: 'left' as const, leader: 'none' as const }];
    const res = layoutBidiTabStops(items, stops, avail, 1000 /* no auto grid */);
    // Following content's RIGHT edge at the stop (visual x 250).
    expect(res[2].visualX + 30).toBeCloseTo(250, 1);
    expect(res[1].leader ?? 'none').toBe('none');
  });

  it('centers content around a mirrored center stop', () => {
    const items: BidiTabItem[] = [
      { isTab: false, width: 40 },
      { isTab: true, width: 0 },
      { isTab: false, width: 20 },
    ];
    const stops = [{ pos: 200, alignment: 'center' as const, leader: 'none' as const }];
    const res = layoutBidiTabStops(items, stops, avail, 1000);
    // Center stop mirrors to visual x avail-200 = 200; the width-20 content is
    // centered on it → spans [190, 210], midpoint 200.
    expect(res[2].visualX + 10).toBeCloseTo(200, 1);
  });

  it('is a no-op shape for a line with no tabs (widths unchanged)', () => {
    const items: BidiTabItem[] = [{ isTab: false, width: 10 }, { isTab: false, width: 20 }];
    const res = layoutBidiTabStops(items, [], avail, 36);
    expect(res.map((r) => r.width)).toEqual([10, 20]);
  });
});

// ── End-to-end: render a synthetic bidi TOC row + footer row through
// renderDocumentToCanvas with a recording canvas (fixed-width glyphs), so the
// full layout→reorder→draw path is exercised (not just the pure helper). The
// mock glyph is `fontSize` px wide; positions are therefore exact and font-free.
interface FillCall { text: string; x: number; }
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[]; leaderXs: number[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fills: FillCall[] = [];
  const leaderXs: number[] = [];
  let dir: CanvasDirection = 'ltr';
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    letterSpacing: '0px',
    get direction() { return dir; }, set direction(v: CanvasDirection) { dir = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, rect() {}, clip() {}, scale() {},
    translate() {}, setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; }, drawImage() {},
    fillText(text: string, x: number) {
      if (text === '_' || text === '.' || text === '·' || text === '-') leaderXs.push(x);
      else fills.push({ text, x });
    },
    strokeText() {}, fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, globalAlpha: 1,
    lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, fills, leaderXs };
}

function bidiPara(runs: unknown[], tabStops: unknown[], opts: Partial<DocParagraph> = {}): DocParagraph {
  return {
    alignment: 'left', bidi: true,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
    tabStops, runs,
    defaultFontSize: 10, defaultFontFamily: 'Arial', widowControl: false,
    ...opts,
  } as unknown as DocParagraph;
}
function txt(text: string, rtl = false) {
  return {
    type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Arial', fontFamilyEastAsia: 'Arial',
    isLink: false, background: null, vertAlign: null, hyperlink: null, rtl: rtl || undefined,
  };
}
function docOf(paras: DocParagraph[], width = 400): DocxDocumentModel {
  return {
    section: {
      pageWidth: width, pageHeight: 400, marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    settings: { defaultTabStop: 36 },
    body: paras.map((p) => ({ type: 'paragraph', ...p })),
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { Arial: 'swiss' },
  } as unknown as DocxDocumentModel;
}

describe('bidi TOC / footer rows render on one line, mirrored (issue #820)', () => {
  it('draws the TOC page number at the left, chapter at the right, with a leader', async () => {
    const { canvas, fills, leaderXs } = makeRecordingCanvas();
    // pageWidth 400, no margins ⇒ scale 1, content [0,400]. Right/underscore
    // stop at 380pt (near the leading/right edge distance), so the page number
    // trails at 400-380 = 20.
    const row = bidiPara(
      [txt('AB', true), txt(' ', true), txt('TITLE', true), txt('\t', true), txt('9', true)],
      [{ pos: 380, alignment: 'right', leader: 'underscore' }],
    );
    await renderDocumentToCanvas(docOf([row]), canvas, 0, { dpr: 1, width: 400 });

    const pageNum = fills.find((f) => f.text === '9');
    const chapter = fills.find((f) => f.text === 'AB');
    expect(pageNum, 'page number drawn').toBeDefined();
    expect(chapter, 'chapter number drawn').toBeDefined();
    // Page number on the visual LEFT (near x=20), chapter number on the visual
    // RIGHT (its 2 glyphs = 20px end at the right edge 400 ⇒ starts near 380).
    expect(pageNum!.x).toBeCloseTo(20, 0);
    expect(chapter!.x).toBeGreaterThan(pageNum!.x + 100);
    // A continuous underscore leader fills the gap between them.
    expect(leaderXs.length).toBeGreaterThan(3);
    expect(Math.min(...leaderXs)).toBeGreaterThan(pageNum!.x);
    expect(Math.max(...leaderXs)).toBeLessThan(chapter!.x);
    // Everything on ONE line (no wrap): the page number and chapter share a row.
    // A single-line paragraph draws each token exactly once.
    expect(fills.filter((f) => f.text === '9')).toHaveLength(1);
  });

  it('right-aligns a footer field row (Page N tab of M) to the leading edge', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    // Footer with a trailing right tab at 380: the "of" cell right-aligns near
    // the right edge. Verify the last token ends at/near the right margin.
    const row = bidiPara(
      [txt('P', true), txt('\t', true), txt('N', true)],
      [{ pos: 380, alignment: 'right', leader: 'none' }],
    );
    await renderDocumentToCanvas(docOf([row]), canvas, 0, { dpr: 1, width: 400 });
    const p = fills.find((f) => f.text === 'P');
    const n = fills.find((f) => f.text === 'N');
    expect(p).toBeDefined();
    expect(n).toBeDefined();
    // "P" is the leading (logical-first) token → visual RIGHT edge; "N" trails at
    // the mirrored stop (400-380 = 20) on the visual LEFT.
    expect(n!.x).toBeCloseTo(20, 0);
    expect(p!.x).toBeGreaterThan(n!.x);
  });
});

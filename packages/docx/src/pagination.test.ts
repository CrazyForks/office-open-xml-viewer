import { describe, it, expect } from 'vitest';
import { computePages } from './renderer.js';
import type { BodyElement, DocParagraph, DocxTextRun, ShapeRun, SectionProps, PaginatedBodyElement } from './types';

// Unit tests for computePages pagination behaviour that the renderer-path VRT
// (local-only, private samples) cannot guard in CI. A deterministic stub canvas
// makes line wrapping and line heights predictable: glyph advance = charCount ×
// fontPx, and the font box = 0.8/0.2 em (so a single line is exactly fontPx tall
// with no spacing/grid). CJK characters break between any two glyphs, so a run of
// N of them wraps to ceil(N / charsPerLine) lines.

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

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 140,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocRun;
}

type DocRun = DocParagraph['runs'][number];

function para(opts: { text?: string; fontSize?: number; widowControl?: boolean } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: opts.text ? [textRun(opts.text, fontSize)] : [],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
    widowControl: opts.widowControl,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

/** A minimal anchored text-box shape (wps:wsp under wp:anchor). `wrapMode`
 *  defaults to topAndBottom so it reserves a full-width float band. */
function shapeRun(opts: {
  widthPt: number;
  heightPt: number;
  anchorYPt?: number;
  anchorYFromPara?: boolean;
  anchorXFromMargin?: boolean;
  wrapMode?: string | null;
  wrapSide?: string | null;
  distTop?: number;
  distBottom?: number;
  distLeft?: number;
  distRight?: number;
}): DocRun {
  const s: ShapeRun = {
    widthPt: opts.widthPt,
    heightPt: opts.heightPt,
    anchorXPt: 0,
    anchorYPt: opts.anchorYPt ?? 0,
    anchorXFromMargin: opts.anchorXFromMargin ?? true,
    anchorYFromPara: opts.anchorYFromPara ?? true,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: { fillType: 'solid', color: 'FFFFFF' },
    stroke: null,
    wrapMode: opts.wrapMode === undefined ? 'topAndBottom' : opts.wrapMode,
    wrapSide: opts.wrapSide ?? null,
    distTop: opts.distTop ?? 0,
    distBottom: opts.distBottom ?? 0,
    distLeft: opts.distLeft ?? 0,
    distRight: opts.distRight ?? 0,
  };
  return { type: 'shape', ...s } as DocRun;
}

/** A paragraph carrying an explicit run list (e.g. an anchored shape, optionally
 *  followed by inline text). */
function paraWith(runs: DocRun[], opts: { fontSize?: number } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

const sliceOf = (el: PaginatedBodyElement) =>
  (el as { lineSlice?: { start: number; end: number } }).lineSlice;

describe('computePages — empty-paragraph relocation (C2: §17.3.1.29)', () => {
  it('moves an unsplittable mark-only paragraph to the next page instead of overflowing the bottom margin', () => {
    // content height = 140 - 40 = 100; each empty mark = 20px → exactly 5 per page.
    const body = Array.from({ length: 7 }, () => para()); // 7 empty paragraphs
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    expect(pages[0].length).toBe(5); // page 1 fills exactly
    expect(pages[1].length).toBe(2); // overflow relocated, NOT clipped onto page 1
    // no page holds more than its 5-line capacity (would mean an overflow)
    for (const p of pages) expect(p.length).toBeLessThanOrEqual(5);
  });
});

describe('computePages — line-boundary splitting + widowControl (C1: §17.3.1.44)', () => {
  // contentW = 160, glyph advance = fontPx; at 20px → 8 chars/line. 48 chars → 6 lines.
  // content height = 100 → 5 lines (100px) fit per page.
  const sixLineText = 'あ'.repeat(48);

  it('avoids a widow: a single trailing line is not stranded on the next page (default widowControl on)', () => {
    const pages = computePages([para({ text: sixLineText })], section(), makeCtx());
    expect(pages.length).toBe(2);
    // Greedy fit is 5 lines on page 1; widowControl pulls one down so ≥2 carry over.
    expect(sliceOf(pages[0][0])).toEqual({ start: 0, end: 4 });
    expect(sliceOf(pages[1][0])).toEqual({ start: 4, end: 6 });
  });

  it('honors w:widowControl="off": the trailing single line is allowed (matches sample-9)', () => {
    const pages = computePages([para({ text: sixLineText, widowControl: false })], section(), makeCtx());
    expect(pages.length).toBe(2);
    expect(sliceOf(pages[0][0])).toEqual({ start: 0, end: 5 }); // greedy 5 lines
    expect(sliceOf(pages[1][0])).toEqual({ start: 5, end: 6 }); // lone widow line allowed
  });
});

describe('computePages — anchored wrap-shape float exclusion (B: §20.4.2.16)', () => {
  // A topAndBottom anchored text-box SHAPE must reserve a float band exactly
  // like an anchored image does, so body text in following paragraphs flows
  // BELOW it instead of overlapping. Geometry: page content height =
  // 140 - 20 - 20 = 100pt. The shape is anchored at the first paragraph's top
  // (≈ marginTop = 20pt) with height 50 ⇒ band y∈[20,70]. The following text
  // paragraph is 2 lines × 20pt = 40pt.
  //
  // Without the shape registering a float (the bug): the shape paragraph's mark
  // (20pt: y20→40) + the 2-line text (40pt: y40→80) all fit ⇒ 1 page.
  // With the float (the fix): the mark flows below the band (y70→90) and the
  // text is pushed under the band, so its 2 lines (≥ y90 → ≈130) overflow the
  // 100pt content area ⇒ 2 pages.
  it('pushes following text below a topAndBottom shape (shape registers a float)', () => {
    const body = [
      paraWith([shapeRun({ widthPt: 160, heightPt: 50, wrapMode: 'topAndBottom' })]),
      para({ text: 'あ'.repeat(16), fontSize: 20 }), // 160/20 = 8 chars/line → 2 lines
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    // The text paragraph (with its float-displaced first line) is pushed to the
    // second page; page 1 holds only the shape-anchoring paragraph.
    const onPage2 = pages[1].some(
      (el) => el.type === 'paragraph' &&
        (el as unknown as DocParagraph).runs.some((r) => r.type === 'text'),
    );
    expect(onPage2).toBe(true);
  });

  it('does NOT reserve a band for a wrapNone shape (no float, text stays on one page)', () => {
    // Same geometry but wrapMode:'none' ⇒ no exclusion rect; everything fits on
    // one page. Guards against over-registering floats for non-wrapping shapes.
    const body = [
      paraWith([shapeRun({ widthPt: 160, heightPt: 50, wrapMode: 'none' })]),
      para({ text: 'あ'.repeat(16), fontSize: 20 }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
  });
});

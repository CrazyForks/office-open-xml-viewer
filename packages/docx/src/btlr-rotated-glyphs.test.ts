import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement, DocParagraph, DocxTextRun, SectionProps, DocxDocumentModel,
} from './types';

// ECMA-376 §17.6.20 + Part 4 §14.11.7 — section-level `btLr` glyph treatment
// (issue #988 re-adjudication, correcting adjudication ①). Word GT (raster-
// proven on asymmetric glyphs — the dakuten of 「び」 lands bottom-right):
// a `btLr` section is the horizontal layout rotated +90° CW WHOLESALE. It
// shares the `tbRl` page frame (quarter-turned logical geometry, +90° page
// paint, columns right→left, advance top→bottom) but — unlike `tbRl` — CJK
// glyphs are NOT counter-rotated upright, vertical punctuation forms are NOT
// substituted, and 縦中横 (§17.3.2.10) is NOT grouped: every glyph rotates
// with the page. These tests pin the glyph mode split with a recording canvas:
//   - tbRl page: the +π/2 page rotation PLUS per-glyph −π/2 counter-rotations.
//   - btLr page: ONLY the +π/2 page rotation — runs draw as ordinary whole-run
//     `fillText` calls (the horizontal draw path inside the rotated frame).

// Deterministic stub canvas that records rotate() and fillText() calls.
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  rotations: number[];
  fills: string[];
} {
  let font = '10px serif';
  const rotations: number[] = [];
  const fills: string[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    rotate(r: number) { rotations.push(r); },
    setTransform() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(t: string) { fills.push(t); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, rotations, fills };
}

// `renderDocumentToCanvas` paginates via `new OffscreenCanvas(...)` when no
// prebuilt pages are passed — polyfill with the same deterministic stub.
(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() {
    return makeRecordingCanvas().canvas.getContext('2d');
  }
};

type DocRun = DocParagraph['runs'][number];
function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null, ...extra,
  } as unknown as DocxTextRun;
  return { type: 'text', ...run } as DocRun;
}
function paraOf(runs: DocRun[]): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs, defaultFontSize: 20, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph;
  return { type: 'paragraph', ...p } as BodyElement;
}

const EMPTY_HF = { default: null, first: null, even: null };

function verticalDoc(textDirection: 'btLr' | 'tbRl', runs: DocRun[]): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 612, pageHeight: 792,
    marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
    headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
    textDirection,
  } as SectionProps;
  return {
    section, body: [paraOf(runs)],
    headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

const PAGE_ROT = Math.PI / 2;
const GLYPH_ROT = -Math.PI / 2;
const near = (a: number, b: number) => Math.abs(a - b) < 1e-9;

describe('section btLr glyph mode (§17.6.20 + #988 re-adjudication) — all glyphs rotate with the page', () => {
  it('tbRl control: CJK glyphs are counter-rotated upright (−π/2 per glyph)', async () => {
    const { canvas, rotations } = makeRecordingCanvas();
    await renderDocumentToCanvas(verticalDoc('tbRl', [textRun('縦横テスト')]), canvas, 0, { dpr: 1 });
    // One +π/2 page rotation…
    expect(rotations.filter((r) => near(r, PAGE_ROT)).length).toBe(1);
    // …and one −π/2 counter-rotation per upright CJK glyph (5 here).
    expect(rotations.filter((r) => near(r, GLYPH_ROT)).length).toBe(5);
  });

  it('btLr: the page rotates +π/2 ONCE and NO glyph is counter-rotated', async () => {
    const { canvas, rotations, fills } = makeRecordingCanvas();
    await renderDocumentToCanvas(verticalDoc('btLr', [textRun('縦横テスト')]), canvas, 0, { dpr: 1 });
    // The page frame is still the rotated vertical one (frame unchanged from
    // tbRl — #988 ruling ①'s geometry findings stand)…
    expect(rotations.filter((r) => near(r, PAGE_ROT)).length).toBe(1);
    // …but the glyphs ride the page rotation: no upright counter-rotation.
    expect(rotations.filter((r) => near(r, GLYPH_ROT)).length).toBe(0);
    // The run is drawn as ONE contextual whole-run fillText (the horizontal
    // draw path), not per-glyph pieces — the raster equals the horizontal
    // rendering rotated +90°.
    expect(fills).toContain('縦横テスト');
  });

  it('btLr: 縦中横 (§17.3.2.10 eastAsianVert) is NOT grouped — the run stays on the horizontal path', async () => {
    const runs = [textRun('令和'), textRun('２９', { eastAsianVert: true } as Partial<DocxTextRun>)];
    const { canvas, rotations, fills } = makeRecordingCanvas();
    await renderDocumentToCanvas(verticalDoc('btLr', runs), canvas, 0, { dpr: 1 });
    // No counter-rotation at all — neither upright CJK cells nor a 縦中横 cell
    // (drawTateChuYokoRun would record a −π/2 too).
    expect(rotations.filter((r) => near(r, GLYPH_ROT)).length).toBe(0);
    expect(fills).toContain('令和');
    expect(fills).toContain('２９');

    // tbRl control: the same runs DO counter-rotate (upright cells + 縦中横).
    const { canvas: c2, rotations: r2 } = makeRecordingCanvas();
    await renderDocumentToCanvas(verticalDoc('tbRl', runs), c2, 0, { dpr: 1 });
    expect(r2.filter((r) => near(r, GLYPH_ROT)).length).toBeGreaterThan(0);
  });

  it('btLr: the numbering marker also rides the page rotation (no upright draw)', async () => {
    const p: DocParagraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
      numbering: { numId: 1, level: 0, format: 'japaneseCounting', text: '一、', indentLeft: 40, tab: 40, suff: 'tab' },
      tabStops: [],
      runs: [textRun('項目')], defaultFontSize: 20, defaultFontFamily: 'NotInMetrics', widowControl: false,
    } as unknown as DocParagraph;
    const doc = {
      ...verticalDoc('btLr', [textRun('x')]),
      body: [{ type: 'paragraph', ...p } as BodyElement],
    } as DocxDocumentModel;
    const { canvas, rotations, fills } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, canvas, 0, { dpr: 1 });
    // The marker WAS painted (whole-string fillText — non-vacuity guard)…
    expect(fills).toContain('一、');
    // …and rode the page rotation: no upright counter-rotation anywhere.
    expect(rotations.filter((r) => near(r, GLYPH_ROT)).length).toBe(0);
  });
});

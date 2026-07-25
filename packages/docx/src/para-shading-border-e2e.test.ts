import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  ParagraphBorders,
  SectionProps,
} from './types';

// ECMA-376 §17.3.1.7 / §17.3.1.31 — paragraph shading fills the retained border
// box. This integration test exercises the canonical layout-and-paint path and
// records the resulting Canvas operations: a multi-line paragraph's shading
// edges must meet its retained top and bottom border segments.

interface FillRectCall { x: number; y: number; w: number; h: number; fillStyle: string; }
interface HStroke { y: number; strokeStyle: string; }

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillRects: FillRectCall[];
  hStrokes: HStroke[];
  textBaselines: number[];
} {
  let font = '10px serif';
  let fillStyle = '#000';
  let strokeStyle = '#000';
  const fillRects: FillRectCall[] = [];
  const hStrokes: HStroke[] = [];
  const textBaselines: number[] = [];
  // Pending path points for the current beginPath..stroke cycle. We only care
  // about axis-aligned HORIZONTAL segments (y1 === y2), which is how a paragraph
  // bottom border materializes.
  let path: { x: number; y: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    get strokeStyle() { return strokeStyle; },
    set strokeStyle(v: string) { strokeStyle = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {},
    beginPath() { path = []; },
    closePath() {},
    moveTo(x: number, y: number) { path.push({ x, y }); },
    lineTo(x: number, y: number) { path.push({ x, y }); },
    stroke() {
      // Record every horizontal segment (consecutive points sharing a y).
      for (let i = 1; i < path.length; i++) {
        if (path[i].y === path[i - 1].y) {
          hStrokes.push({ y: path[i].y, strokeStyle });
        }
      }
    },
    fill() {},
    fillRect(x: number, y: number, w: number, h: number) {
      fillRects.push({ x, y, w, h, fillStyle });
    },
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    // fillText(text, x, y): y is the baseline of the drawn glyph run.
    fillText(_t: string, _x: number, y: number) { textBaselines.push(y); },
    strokeText() {},
    lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillRects, hStrokes, textBaselines };
}

// A single edge: `single` style, distinct color, and space 0 so the shading rect
// is NOT extended past the text box (paraShadingRect leaves each edge in place).
const BORDER_COLOR = '112233';
const edge = (): NonNullable<ParagraphBorders['bottom']> =>
  ({ style: 'single', color: BORDER_COLOR, width: 1, space: 0 } as NonNullable<ParagraphBorders['bottom']>);

const allBorders = (): ParagraphBorders => ({
  top: edge(), bottom: edge(), left: edge(), right: edge(), between: null,
});

const SHADING = 'ffcc00';

function borderedShadedPara(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    shading: SHADING,
    borders: allBorders(),
    runs: [
      {
        type: 'text', text, bold: false, italic: false, underline: false,
        strikethrough: false, fontSize: 11, color: null, fontFamily: 'Times New Roman',
        fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null,
        hyperlink: null,
      } as DocParagraph['runs'][number],
    ],
    defaultFontSize: 11, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

// Narrow page (margins 0) so the space-separated words wrap onto several lines.
const PAGE_WIDTH = 120;

function docOf(...paras: DocParagraph[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: PAGE_WIDTH, pageHeight: 2000,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: paras.map((p) => p as unknown as BodyElement),
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('paragraph shading meets the bottom border, end-to-end (§17.3.1.7)', () => {
  it('the drawn shading rect bottom coincides with the drawn bottom border (multi-line)', async () => {
    const { canvas, fillRects, hStrokes, textBaselines } = makeRecordingCanvas();
    // Many short words wrap to several retained lines at PAGE_WIDTH=120.
    const text = 'aa bb cc dd ee ff gg hh ii jj kk ll mm nn oo pp qq rr ss tt';
    await renderDocumentToCanvas(docOf(borderedShadedPara(text)), canvas, 0, {
      dpr: 1, width: PAGE_WIDTH, // scale = width / pageWidthPt(=120) = 1 px per pt
      fetchImage: async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    });

    // The shading fill is the wide, tall rect painted in the shading color and
    // spanning the full paragraph width (paraW === PAGE_WIDTH at scale 1, margins 0).
    const shadingRects = fillRects.filter(
      (r) => r.fillStyle === `#${SHADING}` && Math.abs(r.w - PAGE_WIDTH) < 0.5,
    );
    expect(shadingRects.length).toBe(1);
    const shading = shadingRects[0];
    expect(shading.h).toBeGreaterThan(0);

    // The bottom border is the bottom-most horizontal stroke painted in the
    // border color (top border is at the same-color stroke with the SMALLEST y).
    const borderStrokes = hStrokes.filter((s) => s.strokeStyle === `#${BORDER_COLOR}`);
    expect(borderStrokes.length).toBeGreaterThanOrEqual(2); // at least top + bottom
    const bottomBorderY = Math.max(...borderStrokes.map((s) => s.y));
    const topBorderY = Math.min(...borderStrokes.map((s) => s.y));

    // Axis-aligned strokes can receive a sub-pixel device snap, so allow less
    // than one pixel while requiring both shading edges to meet their borders.
    const NUDGE = 0.75;
    expect(Math.abs(shading.y + shading.h - bottomBorderY)).toBeLessThan(NUDGE);
    // The shading top also meets the top border (sanity: no vertical drift).
    expect(Math.abs(shading.y - topBorderY)).toBeLessThan(NUDGE);

    // Distinct baselines prove this is a multi-line retained paragraph.
    const distinctBaselines = new Set(textBaselines.map((y) => Math.round(y))).size;
    expect(distinctBaselines).toBeGreaterThanOrEqual(3);

    // And the shading height therefore spans multiple lines: it is well over one
    // line's height (each line ≈ 15.8 px for 11 pt here). Guards against a single
    // tall line coincidentally satisfying the coupling above.
    const boxH = bottomBorderY - topBorderY;
    const singleLineMaxH = 40; // generous upper bound for one 11 pt line box (px)
    expect(boxH).toBeGreaterThan(singleLineMaxH);
  });
});

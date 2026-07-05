import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  ParagraphBorders,
  SectionProps,
} from './types';

// ECMA-376 §17.3.1.7 — a paragraph BOTTOM border is drawn `w:space` points below
// the text ("the space after the bottom of the text … before this border is
// drawn"), and §17.3.4 gives the border its own width (`w:sz`, eighths of a point).
// drawParaBorders strokes the line CENTERED on `textBottom + space`, so its outer
// (bottom) edge is at `textBottom + space + width/2`. Word reserves that whole
// extent in the vertical flow — a bottom-bordered paragraph pushes the FOLLOWING
// paragraph BELOW the border rather than letting its first line box overlap the
// rule. The spec is silent on the flow reservation; this is Word's observed layout
// (sample-14: the reference-list rule sat ~1.75 pt too high — half a border-width
// plus its space — so "Further examples…" nearly touched the rule).
//
// The renderer used to draw the bottom border PAST `state.y` without advancing the
// flow by that extent, so the next paragraph overlapped it. This test measures the
// flow delta a bottom border introduces: the follower's baseline must drop by
// exactly `space + width/2` versus an identical layout whose leading paragraph has
// NO border.

const BORDER_COLOR = 'aa00bb';
const SPACE_PT = 1;
const WIDTH_PT = 1.5; // sz12 = 12 eighths of a point → 1.5 pt
const EXPECTED_EXTENT = SPACE_PT + WIDTH_PT / 2; // 1.75 pt

interface HStroke { y: number; strokeStyle: string; }

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  hStrokes: HStroke[];
  textBaselines: number[];
} {
  let font = '10px serif';
  let strokeStyle = '#000';
  const hStrokes: HStroke[] = [];
  const textBaselines: number[] = [];
  let path: { x: number; y: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    fillStyle: '#000',
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
      for (let i = 1; i < path.length; i++) {
        if (path[i].y === path[i - 1].y) hStrokes.push({ y: path[i].y, strokeStyle });
      }
    },
    fill() {},
    fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(_t: string, _x: number, y: number) { textBaselines.push(y); },
    strokeText() {},
    lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, hStrokes, textBaselines };
}

function bottomBorderOnly(): ParagraphBorders {
  return {
    top: null,
    bottom: { style: 'single', color: BORDER_COLOR, width: WIDTH_PT, space: SPACE_PT } as NonNullable<ParagraphBorders['bottom']>,
    left: null, right: null, between: null,
  };
}

function para(text: string, borders: ParagraphBorders | null): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    borders,
    runs: text
      ? [{
          type: 'text', text, bold: false, italic: false, underline: false,
          strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
          fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null,
          hyperlink: null,
        } as DocParagraph['runs'][number]]
      : [],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

const PAGE_WIDTH = 400;
function docOf(...paras: DocParagraph[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: PAGE_WIDTH, pageHeight: 4000,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: paras.map((p) => p as unknown as BodyElement),
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function followerBaseline(leadHasBorder: boolean): Promise<{ baseline: number; borderY: number | null }> {
  const { canvas, hStrokes, textBaselines } = makeRecordingCanvas();
  await renderDocumentToCanvas(
    docOf(para('', leadHasBorder ? bottomBorderOnly() : null), para('Follower', null)),
    canvas, 0, {
      dpr: 1, width: PAGE_WIDTH, // scale = 1 px per pt
      fetchImage: async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    });
  const borderStrokes = hStrokes.filter((s) => s.strokeStyle === `#${BORDER_COLOR}`);
  return {
    baseline: Math.min(...textBaselines),
    borderY: borderStrokes.length ? Math.max(...borderStrokes.map((s) => s.y)) : null,
  };
}

describe('a bottom paragraph border reserves flow so following content clears it (§17.3.1.7)', () => {
  it('a bottom border drops the following paragraph by exactly space + width/2', async () => {
    // Baseline layouts differ ONLY in whether the leading empty paragraph carries a
    // bottom border. The border must push the follower down by its outer extent
    // (space + half the stroke width) so the follower's line box clears the rule.
    const withBorder = await followerBaseline(true);
    const noBorder = await followerBaseline(false);

    expect(withBorder.borderY).not.toBeNull();
    const delta = withBorder.baseline - noBorder.baseline;
    // Exact reservation (a sub-pixel crispness nudge on the stroke does not move the
    // baseline, which is placed by the flow cursor, so no tolerance is needed here —
    // allow a hair for float arithmetic).
    expect(delta).toBeCloseTo(EXPECTED_EXTENT, 3);

    // And the follower's line box now clears the border's OUTER edge: with a
    // 10 pt font (ascent 0.8 em = 8 pt here) the box top = baseline − 8 sits at or
    // below the border edge = borderY + width/2.
    const boxTop = withBorder.baseline - 8;
    expect(boxTop).toBeGreaterThanOrEqual((withBorder.borderY as number) + WIDTH_PT / 2 - 1e-6);
  });
});

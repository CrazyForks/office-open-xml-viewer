import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  ParagraphBorders,
  SectionProps,
} from './types';

// ECMA-376 §17.3.1.7 / §17.3.1.31 — the paragraph shading fill and the paragraph
// BORDER box must have the SAME height. renderParagraph paints the shading BEFORE
// the draw loop (it is the background) using `paintedParagraphHeight`, which
// REPLAYS the loop's per-line `state.y` advancement; it then draws the bottom
// border AFTER the loop at `state.y − textAreaTopY`. Equality therefore holds
// BY CONSTRUCTION — but only as long as `paintedParagraphHeight` replays EXACTLY
// the same `state.y` mutations the draw loop performs. Today `drawParagraphLine`
// mutates `state.y` in exactly two places (the `line.topY` float-clearance
// max-jump, then `state.y += lineHForLine(line)`), and `paintedParagraphHeight`
// mirrors both.
//
// The existing unit tests in `para-shading-border.test.ts` pin
// `paintedParagraphHeight` against a HAND-REPLAY of that SAME formula, so they
// would NOT notice if someone added a THIRD `state.y` mutation to the real draw
// loop (the replay would silently drift from the loop). This end-to-end test
// closes that gap: it renders a real bordered + shaded MULTI-LINE paragraph
// through renderDocumentToCanvas → renderParagraph → drawParagraphLine and asserts
// the drawn shading rect and the drawn bottom border coincide. A new state.y
// mutation in the draw loop would move the bottom border relative to the
// pre-computed shading height and fail this test.
//
// Assertion approach chosen: DIRECT BORDER CAPTURE. With `space:0` on every edge
// and a single-style bottom border, `paraShadingRect` does NOT extend the fill
// box (so the shading rect bottom == textAreaTopY + paintedH) and the bottom
// border is a `single` edge — which `drawBorderLine` renders via
// `strokeCrispSegment`'s beginPath/moveTo/lineTo/stroke. We record the horizontal
// stroked segments and take the bottom-most one as the bottom border edge y, then
// assert it equals the shading rect bottom. This is exact and does not depend on a
// trailing paragraph's flow position. (The crispness nudge in strokeCrispSegment
// is sub-pixel; at dpr:1 with integer-ish coordinates we allow a tiny tolerance.)
//
// Scope: this covers the NORMAL-FLOW multi-line case only. That already catches a
// third `state.y` mutation on the common draw path (verified: injecting an extra
// `state.y += 1` per line drifts the bottom border ~one px/line past the shading
// bottom and fails the coupling assert). The FLOAT-CLEARANCE `topY` max-jump branch
// is DELIBERATELY OMITTED here: reproducing an in-paragraph `line.topY` jump end-to-
// end needs a preceding float whose band overlaps only some of the shaded
// paragraph's lines — fragile geometry (image sizing, wrap band overlap) with high
// flake risk for little added coverage. That branch is already pinned by the pure
// `paintedParagraphHeight` float test in `para-shading-border.test.ts`, and both
// mutations a 3rd would sit beside are on the path this test exercises.

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
    // Many short words → wraps to several lines at PAGE_WIDTH=120, exercising the
    // `state.y += lineHForLine` summation (not a single-line shortcut).
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

    // Coupling: shading rect bottom (== textAreaTopY + paintedH) equals the drawn
    // bottom border edge (== textAreaTopY + textH). If a future change adds a third
    // state.y mutation to the draw loop, paintedH drifts from textH and this fails.
    //
    // `strokeCrispSegment` nudges a thin (odd device-width) axis-aligned stroke by
    // up to 0.5 px perpendicular to snap it onto a crisp device row (crispOffset);
    // that nudge is a sub-pixel rendering detail orthogonal to the height coupling,
    // so we allow a 0.75 px slack (< 1 px, safely below a full line). Both edges get
    // the SAME nudge, so the border-box height itself is exact regardless.
    const NUDGE = 0.75;
    expect(Math.abs(shading.y + shading.h - bottomBorderY)).toBeLessThan(NUDGE);
    // The shading top also meets the top border (sanity: no vertical drift).
    expect(Math.abs(shading.y - topBorderY)).toBeLessThan(NUDGE);

    // Multi-line proof: the text drew on at least THREE distinct baselines, so the
    // `state.y += lineHForLine` summation ran over multiple lines (not a single-line
    // shortcut that would happen to line up top and bottom borders too). Distinct
    // baselines are the direct witness of how many lines the draw loop advanced past.
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

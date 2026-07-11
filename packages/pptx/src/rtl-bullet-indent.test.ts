import { describe, it, expect } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { TextBody, Paragraph, Bullet } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// ECMA-376 §21.1.2.2.7 (a:pPr@rtl) + §21.1.2.4.x (a:buChar / a:buAutoNum) — a list
// MARKER in a right-to-left paragraph must lead at the RIGHT (the reading start),
// with the text following CONTIGUOUSLY to its left. Issue #930 (same reading-frame
// class as the RTL tab-stop fix #831/#913 and docx #830): the marker draw mirrored
// the LTR hanging gutter about the text's right edge —
// `textX + textMaxW + (textX − bulletX) − markerW`, i.e. it added the whole
// `|indent|` — which pushed the marker |indent| PAST the leading edge into the
// right margin (the "far-frame over-indent"). PowerPoint instead seats the
// marker's RIGHT edge on the line's leading (right) edge and right-aligns the text
// to `leadingEdge − markerAdvance`, so the words render adjacent to the marker.
//
// The bug was most visible for a NARROW bullet (its gap to the text equalled the
// full hanging indent) but applied to autoNum markers too; the fix corrects both.
//
// The mock ctx measures every glyph at the active font's px size, so widths and
// x are exact. It also records the font at each fillText so the marker's advance
// (and hence its right edge) is recoverable.

const FONT_PX = 20;
const SCALE = 1 / 12700; // emuToPx(emu) = emu·SCALE; 12700 EMU = 1 pt → 1 px

interface Fill { text: string; x: number; direction: CanvasDirection; fontPx: number }

function mockCtx(): { ctx: CanvasRenderingContext2D; fills: Fill[] } {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  let fillStyle = '';
  let direction: CanvasDirection = 'ltr';
  const px = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fills: Fill[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; }, set fillStyle(v: string) { fillStyle = v; },
    get direction() { return direction; }, set direction(v: CanvasDirection) { direction = v; },
    get letterSpacing() { return letterSpacing; }, set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    fillText: (t: string, x: number) => fills.push({ text: t, x, direction, fontPx: px() }),
    strokeText: () => {}, fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    setLineDash: () => {}, closePath: () => {}, arc: () => {},
    strokeStyle: '#000', lineWidth: 1, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fills };
}

function run(text: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: FONT_PX, color: '000000', fontFamily: 'Serif',
  } as TextRunData;
}

// marL / indent = ±MARK_EMU (a hanging gutter); at SCALE, MARK_EMU/12700 px.
const MARK_EMU = 342900; // 27 px at this SCALE
const MARK_PX = MARK_EMU * SCALE; // 27

function body(bullet: Bullet, text: string, rtl: boolean): TextBody {
  const para: Paragraph = {
    alignment: rtl ? 'r' : 'l',
    marL: MARK_EMU, marR: 0, indent: -MARK_EMU,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: [], rtl, runs: [run(text)],
  } as unknown as Paragraph;
  return {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: FONT_PX,
    defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert: 'horz', autoFit: 'none',
  } as unknown as TextBody;
}

interface RunInfo { text: string; inShapeX: number; w: number }
const BOX_W = 600;

function render(b: TextBody): { fills: Fill[]; runs: RunInfo[] } {
  const { ctx, fills } = mockCtx();
  const runs: RunInfo[] = [];
  renderTextBody(
    ctx, b, 0, 0, BOX_W, 400, SCALE,
    null, 0, false, false, '#000000', undefined,
    { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
    (r) => runs.push({ text: r.text, inShapeX: r.inShapeX, w: r.w }),
  );
  return { fills, runs };
}

// Geometry: box 600px, no insets, marR=0 ⇒ the RTL leading edge (where a plain
// RTL paragraph's text right-aligns) is textX + textMaxW = 600.
const LEADING_EDGE = BOX_W; // 600
const charBullet: Bullet = { type: 'char', char: '•', color: null, sizePct: null, fontFamily: null };
const autoNumBullet: Bullet = { type: 'autoNum', numType: 'arabicPeriod', startAt: null } as unknown as Bullet;

/** The marker fill (a non-text glyph run drawn before the paragraph text). */
function markerFill(fills: Fill[], text: string, runText: string): Fill | undefined {
  return fills.find((f) => f.text !== runText && f.text.trim() !== '' && (text ? f.text.includes(text) : true));
}

describe('§21.1.2.4 RTL list marker leads at the right edge (issue #930)', () => {
  it('char bullet: marker RIGHT edge sits on the leading edge, text reserved to its left', () => {
    const { fills, runs } = render(body(charBullet, 'ابج', true));
    const marker = fills.find((f) => f.text.includes('•'))!;
    expect(marker, 'bullet marker drawn').toBeTruthy();
    expect(marker.direction).toBe('rtl');
    const markerW = [...marker.text].length * marker.fontPx;
    const markerRight = marker.x + markerW;
    // (1) marker's RIGHT edge is on the leading (right) edge — NOT |indent| past it.
    //     Before the fix markerRight = LEADING_EDGE + |indent| (27 px over).
    expect(markerRight).toBeCloseTo(LEADING_EDGE, 3);
    // (2) the text is right-aligned to leadingEdge − markerW (contiguous with the
    //     marker's left edge). Before the fix the text stayed at the leading edge
    //     and the marker floated in the right margin.
    const textRight = Math.max(...runs.map((r) => r.inShapeX + r.w));
    expect(textRight).toBeCloseTo(marker.x, 2);
    expect(textRight).toBeLessThan(LEADING_EDGE);
  });

  it('autoNum marker: also leads at the right edge (not pushed into the margin)', () => {
    const { fills } = render(body(autoNumBullet, 'ابج', true));
    // The autoNum label is the only non-text, non-empty fill (e.g. "1.").
    const marker = fills.find((f) => /\d/.test(f.text))!;
    expect(marker, 'autoNum marker drawn').toBeTruthy();
    const markerW = [...marker.text].length * marker.fontPx;
    expect(marker.x + markerW).toBeCloseTo(LEADING_EDGE, 3);
  });

  it('LTR control: a char bullet still hangs at the left gutter (byte-identical)', () => {
    const { fills } = render(body(charBullet, 'abc', false));
    const marker = fills.find((f) => f.text.includes('•'))!;
    // bulletX = textX + indent = marL − |indent| = 0 (the left gutter edge).
    expect(marker.x).toBeCloseTo(0, 3);
    expect(marker.direction).not.toBe('rtl');
  });
});

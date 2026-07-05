import { describe, it, expect } from 'vitest';
import { drawWatermarkTextPath } from './renderer';
import type { TextPath } from './types';

/**
 * ECMA-376 Part 4 §19.1.2.23 `<v:textpath>` text watermark rendering.
 *
 * `drawWatermarkTextPath` draws the watermark string filling the shape box
 * (`fitshape`, §19.1.2.23), rotated by the shape's `rotation` (§19.1.2.19,
 * clockwise) about the box centre, filled with the shape's `fillcolor` at the
 * `<v:fill opacity>` alpha (§19.1.2.5). These tests drive it with a recording
 * canvas context and assert the geometry numerically — the transform sequence
 * (translate to box centre → rotate → non-uniform scale to the box), the alpha,
 * the fill colour, and the drawn text — rather than eyeballing pixels.
 */

interface Op {
  op: string;
  args: number[];
}

function makeRecordingCtx(): {
  ctx: CanvasRenderingContext2D;
  ops: Op[];
  fillTextCalls: { text: string; x: number; y: number; alpha: number; fillStyle: string }[];
} {
  let font = '10px serif';
  let fillStyle = '#000';
  let globalAlpha = 1;
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ops: Op[] = [];
  const fillTextCalls: { text: string; x: number; y: number; alpha: number; fillStyle: string }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    get globalAlpha() { return globalAlpha; },
    set globalAlpha(v: number) { globalAlpha = v; },
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    measureText: (s: string) => {
      const p = px();
      // Deterministic: advance = charCount × 0.6 × em; box = 0.8/0.2 em.
      return {
        width: [...s].length * p * 0.6,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { ops.push({ op: 'save', args: [] }); },
    restore() { ops.push({ op: 'restore', args: [] }); },
    translate(x: number, y: number) { ops.push({ op: 'translate', args: [x, y] }); },
    rotate(a: number) { ops.push({ op: 'rotate', args: [a] }); },
    scale(sx: number, sy: number) { ops.push({ op: 'scale', args: [sx, sy] }); },
    beginPath() {}, moveTo() {}, lineTo() {}, stroke() {}, fillRect() {},
    fillText(text: string, x: number, y: number) {
      fillTextCalls.push({ text, x, y, alpha: globalAlpha, fillStyle });
      ops.push({ op: 'fillText', args: [x, y] });
    },
    strokeText() {},
    strokeStyle: '#000', lineWidth: 1,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, ops, fillTextCalls };
}

const draft = (): TextPath => ({ string: 'DRAFT', fontFamily: 'Calibri' });

describe('drawWatermarkTextPath (§19.1.2.23)', () => {
  it('translates to the box centre, rotates, then non-uniformly scales to fill the box', () => {
    const { ctx, ops } = makeRecordingCtx();
    // Box (x=100, y=200, w=400, h=200). Rotation 315° (Word watermark default).
    drawWatermarkTextPath(ctx, draft(), 100, 200, 400, 200, 315, '#c0c0c0', 0.5);

    const names = ops.map((o) => o.op);
    // Order: save → translate(centre) → rotate → scale → fillText → restore.
    expect(names).toEqual(['save', 'translate', 'rotate', 'scale', 'fillText', 'restore']);

    const translate = ops.find((o) => o.op === 'translate')!;
    // Centre = (100 + 400/2, 200 + 200/2) = (300, 300).
    expect(translate.args[0]).toBeCloseTo(300, 6);
    expect(translate.args[1]).toBeCloseTo(300, 6);

    const rotate = ops.find((o) => o.op === 'rotate')!;
    // 315° clockwise = 315 × π/180 rad.
    expect(rotate.args[0]).toBeCloseTo((315 * Math.PI) / 180, 6);

    // Scale must map the natural text box onto the shape box (fitshape). With
    // REF=100px, natW = 5 chars × 100 × 0.6 = 300; natH = 100×(0.8+0.2)=100.
    const scale = ops.find((o) => o.op === 'scale')!;
    expect(scale.args[0]).toBeCloseTo(400 / 300, 4); // w / natW
    expect(scale.args[1]).toBeCloseTo(200 / 100, 4); // h / natH
  });

  it('draws the text centred at the transformed origin with the fill colour and opacity', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    drawWatermarkTextPath(ctx, draft(), 0, 0, 300, 150, 0, '#808080', 0.35);
    expect(fillTextCalls).toHaveLength(1);
    const call = fillTextCalls[0];
    expect(call.text).toBe('DRAFT');
    // Centred: the transform put the origin at the box centre, so text draws at (0,0).
    expect(call.x).toBeCloseTo(0, 6);
    expect(call.y).toBeCloseTo(0, 6);
    expect(call.alpha).toBeCloseTo(0.35, 6);
    expect(call.fillStyle).toBe('#808080');
  });

  it('sets textAlign/textBaseline to centre so the string is centred in its box', () => {
    const { ctx } = makeRecordingCtx();
    drawWatermarkTextPath(ctx, draft(), 0, 0, 200, 100, 0, '#c0c0c0', 1);
    expect(ctx.textAlign).toBe('center');
    expect(ctx.textBaseline).toBe('middle');
  });

  it('clamps opacity into [0,1] and falls back to a default fill colour when null', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    drawWatermarkTextPath(ctx, draft(), 0, 0, 200, 100, 0, null, 5);
    expect(fillTextCalls[0].alpha).toBe(1); // clamped from 5
    expect(fillTextCalls[0].fillStyle).toBe('#c0c0c0'); // default when color=null
  });

  it('draws nothing for an empty string or a degenerate box', () => {
    const { ctx: c1, fillTextCalls: f1 } = makeRecordingCtx();
    drawWatermarkTextPath(c1, { string: '' }, 0, 0, 200, 100, 0, '#000', 1);
    expect(f1).toHaveLength(0);

    const { ctx: c2, fillTextCalls: f2 } = makeRecordingCtx();
    drawWatermarkTextPath(c2, draft(), 0, 0, 0, 100, 0, '#000', 1);
    expect(f2).toHaveLength(0);
  });
});

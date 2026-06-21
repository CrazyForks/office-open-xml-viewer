import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { preloadImages, renderShapeText } from './renderer';
import type { DocxDocumentModel, ShapeRun, ShapeText } from './types';

/**
 * Inline images living INSIDE a DOCX text box (`<wps:txbx>`) ride on the
 * shape's `textBlocks[i].imagePath` rather than a top-level `image` run. Two
 * things must hold end-to-end:
 *   1. `collectImagePairs` (exercised through `preloadImages`, exactly as
 *      renderer.image.test.ts drives it) must surface those textbox images so
 *      their bytes reach the decode pipeline (WMF decoder included).
 *   2. `renderShapeText` must draw the decoded bitmap fitted to the inner
 *      width, and must NOT throw / draw when the bitmap is missing.
 */

// Recording mock canvas context (extends pagination.test.ts's makeCtx with a
// drawImage spy and a measureText stub: glyph advance = charCount × fontPx).
interface DrawImageCall {
  bmp: unknown;
  x: number;
  y: number;
  w: number;
  h: number;
}
function makeRecordingCtx(): {
  ctx: CanvasRenderingContext2D;
  drawImageCalls: DrawImageCall[];
  fillTextCalls: { text: string; x: number; y: number }[];
} {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const drawImageCalls: DrawImageCall[] = [];
  const fillTextCalls: { text: string; x: number; y: number }[] = [];
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
    save() {}, restore() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {},
    fillText(text: string, x: number, y: number) { fillTextCalls.push({ text, x, y }); },
    strokeText() {},
    drawImage(bmp: unknown, x: number, y: number, w: number, h: number) {
      drawImageCalls.push({ bmp, x, y, w, h });
    },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, drawImageCalls, fillTextCalls };
}

/** A text box (ShapeRun) whose first text block is an inline image and whose
 *  second is a caption — the sample-10 Fig.1 layout. Insets default to 0 so the
 *  fit math is easy to assert. */
function textboxWithImage(overrides: Partial<ShapeText> = {}): ShapeRun {
  const imageBlock: ShapeText = {
    text: '',
    fontSizePt: 10,
    alignment: 'center',
    imagePath: 'word/media/image1.emf',
    mimeType: 'image/x-wmf',
    imageWidthPt: 100,
    imageHeightPt: 50,
    ...overrides,
  };
  const captionBlock: ShapeText = {
    text: 'Fig. 1: A sample figure.',
    fontSizePt: 10,
    alignment: 'center',
  };
  return {
    type: 'shape',
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: null,
    stroke: null,
    textBlocks: [imageBlock, captionBlock],
    textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
  } as unknown as ShapeRun;
}

describe('textbox inline images — collection', () => {
  beforeEach(() => {
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_src: unknown) => ({ width: 4, height: 2, close: () => {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('collectImagePairs (via preloadImages) surfaces a textbox shape image', async () => {
    const fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1, 2, 3])], { type: mime }),
    );
    // A shape run (not an image run) carrying an inline image on its text block.
    const doc = {
      body: [
        { type: 'paragraph', runs: [textboxWithImage()] },
      ],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;

    const map = await preloadImages(doc, fetchImage);

    // The textbox image must have been fetched + decoded and keyed by its path.
    expect(fetchImage).toHaveBeenCalledWith('word/media/image1.emf', 'image/x-wmf');
    expect(map.has('word/media/image1.emf')).toBe(true);
  });
});

describe('textbox inline images — rendering', () => {
  // A bitmap-like stub: drawImage only needs width/height-bearing source.
  const fakeBmp = { width: 200, height: 100, close: () => {} } as unknown as ImageBitmap;

  it('renderShapeText draws the bitmap fitted to the inner width', () => {
    const { ctx, drawImageCalls } = makeRecordingCtx();
    const shape = textboxWithImage(); // natural 100×50 pt
    const images = new Map<string, DecodedImage>([['word/media/image1.emf', fakeBmp]]);

    // Box: 80pt wide so the 100pt-wide image must scale DOWN to innerW=80,
    // height 50 × (80/100) = 40. scale=1 → px == pt.
    const scale = 1;
    renderShapeText(shape, /*x*/ 0, /*y*/ 0, /*w*/ 80, /*h*/ 200, ctx, scale, {}, images);

    expect(drawImageCalls).toHaveLength(1);
    const call = drawImageCalls[0];
    expect(call.bmp).toBe(fakeBmp);
    expect(call.w).toBeCloseTo(80, 5);   // scaled to innerW
    expect(call.h).toBeCloseTo(40, 5);   // aspect preserved
    // innerW == fitW ⇒ centered draw sits flush at x=0.
    expect(call.x).toBeCloseTo(0, 5);
    // Image is the first block ⇒ drawn at the top of the inner box (anchor 't').
    expect(call.y).toBeCloseTo(0, 5);
  });

  it('renderShapeText draws an image smaller than innerW at natural size, centered', () => {
    const { ctx, drawImageCalls } = makeRecordingCtx();
    const shape = textboxWithImage(); // natural 100×50 pt
    const images = new Map<string, DecodedImage>([['word/media/image1.emf', fakeBmp]]);

    // innerW=200 > natural 100 ⇒ keep natural 100×50, centered ⇒ x=(200-100)/2=50.
    renderShapeText(shape, 0, 0, 200, 200, ctx, 1, {}, images);

    expect(drawImageCalls).toHaveLength(1);
    expect(drawImageCalls[0].w).toBeCloseTo(100, 5);
    expect(drawImageCalls[0].h).toBeCloseTo(50, 5);
    expect(drawImageCalls[0].x).toBeCloseTo(50, 5);
  });

  it('renderShapeText with a missing bitmap draws nothing and does not throw', () => {
    const { ctx, drawImageCalls, fillTextCalls } = makeRecordingCtx();
    const shape = textboxWithImage();
    const images = new Map<string, DecodedImage>(); // bitmap NOT present

    expect(() => renderShapeText(shape, 0, 0, 80, 200, ctx, 1, {}, images)).not.toThrow();
    // No image drawn …
    expect(drawImageCalls).toHaveLength(0);
    // … but the caption text block still renders (height was still reserved).
    expect(fillTextCalls.some((c) => c.text === 'Fig. 1: A sample figure.')).toBe(true);
  });
});

// Local alias matching renderer.ts's internal DecodedImage union.
type DecodedImage = ImageBitmap | HTMLImageElement;

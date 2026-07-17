import { describe, expect, it } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  SectionProps,
} from './types.js';

type Matrix = Readonly<{ a: number; b: number; c: number; d: number; e: number; f: number }>;

const identity = (): Matrix => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });

function multiply(outer: Matrix, inner: Matrix): Matrix {
  return {
    a: outer.a * inner.a + outer.c * inner.b,
    b: outer.b * inner.a + outer.d * inner.b,
    c: outer.a * inner.c + outer.c * inner.d,
    d: outer.b * inner.c + outer.d * inner.d,
    e: outer.a * inner.e + outer.c * inner.f + outer.e,
    f: outer.b * inner.e + outer.d * inner.f + outer.f,
  };
}

function recordingCanvas(): {
  canvas: HTMLCanvasElement;
  rotations: number[];
  rotationsAtText: number[];
  regionTransforms: Matrix[];
  textTransforms: Matrix[];
} {
  let matrix = identity();
  const stack: Matrix[] = [];
  let font = '10px serif';
  const rotations: number[] = [];
  const rotationsAtText: number[] = [];
  const regionTransforms: Matrix[] = [];
  const textTransforms: Matrix[] = [];
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, globalAlpha: 1,
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    letterSpacing: '0px', fontKerning: 'auto' as CanvasFontKerning,
    lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
    save() { stack.push(matrix); },
    restore() { matrix = stack.pop() ?? identity(); },
    setTransform(a: number, b: number, c: number, d: number, e: number, f: number) {
      matrix = { a, b, c, d, e, f };
    },
    transform(a: number, b: number, c: number, d: number, e: number, f: number) {
      const next = { a, b, c, d, e, f };
      regionTransforms.push(next);
      matrix = multiply(matrix, next);
    },
    translate(x: number, y: number) {
      matrix = multiply(matrix, { a: 1, b: 0, c: 0, d: 1, e: x, f: y });
    },
    scale(x: number, y: number) {
      matrix = multiply(matrix, { a: x, b: 0, c: 0, d: y, e: 0, f: 0 });
    },
    rotate(angle: number) {
      rotations.push(angle);
      const cos = Math.cos(angle);
      const sin = Math.sin(angle);
      matrix = multiply(matrix, { a: cos, b: sin, c: -sin, d: cos, e: 0, f: 0 });
    },
    measureText(text: string) {
      const size = Number(/([0-9.]+)px/.exec(font)?.[1] ?? 10);
      return {
        width: [...text].length * size * .5,
        actualBoundingBoxLeft: 0,
        actualBoundingBoxRight: [...text].length * size * .5,
        actualBoundingBoxAscent: size * .8,
        actualBoundingBoxDescent: size * .2,
        fontBoundingBoxAscent: size * .8,
        fontBoundingBoxDescent: size * .2,
      } as TextMetrics;
    },
    fillText() {
      textTransforms.push(matrix);
      rotationsAtText.push(rotations.length);
    },
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, rect() {}, clip() {},
    fill() {}, stroke() {}, fillRect() {}, strokeRect() {}, clearRect() {}, setLineDash() {},
    arc() {}, quadraticCurveTo() {}, bezierCurveTo() {}, drawImage() {}, strokeText() {},
    createLinearGradient() { return { addColorStop() {} }; },
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return {
    canvas: canvas as unknown as HTMLCanvasElement,
    rotations,
    rotationsAtText,
    regionTransforms,
    textTransforms,
  };
}

function paragraph(): BodyElement {
  const value = {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text: 'body', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Test Sans',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'Test Sans', widowControl: false,
  } as unknown as DocParagraph;
  return value as BodyElement;
}

function document(textDirection: 'btLr' | 'lrTb'): DocxDocumentModel {
  const section = {
    pageWidth: 200, pageHeight: 300,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 10, footerDistance: 10,
    titlePage: false, evenAndOddHeaders: false, textDirection,
  } as SectionProps;
  return {
    section,
    body: [paragraph()],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

describe('canonical body coordinate-space ownership', () => {
  it('projects a vertical body through its canonical region matrix only once', async () => {
    const recorded = recordingCanvas();

    await renderDocumentToCanvas(document('btLr'), recorded.canvas, 0, { dpr: 1, width: 200 });

    expect(recorded.textTransforms).not.toHaveLength(0);
    expect(recorded.textTransforms[0]).toMatchObject({ a: 0, b: 1, c: -1, d: 0 });
    expect(recorded.regionTransforms).toContainEqual({ a: 0, b: 1, c: -1, d: 0, e: 200, f: 0 });
    expect(recorded.rotationsAtText[0]).toBe(0);
  });

  it('keeps a horizontal body in the unrotated canonical page frame', async () => {
    const recorded = recordingCanvas();

    await renderDocumentToCanvas(document('lrTb'), recorded.canvas, 0, { dpr: 1, width: 200 });

    expect(recorded.textTransforms).not.toHaveLength(0);
    expect(recorded.textTransforms[0]).toMatchObject({ a: 1, b: 0, c: 0, d: 1 });
    expect(recorded.regionTransforms).toEqual([]);
    expect(recorded.rotations).toEqual([]);
  });
});

import { describe, expect, it } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocxDocumentModel, SectionProps } from './types.js';

const FONT_PT = 10;

function recordingCanvas(): {
  canvas: HTMLCanvasElement;
  calls: Array<{ text: string; x: number; y: number }>;
} {
  let font = `${FONT_PT}px serif`;
  const calls: Array<{ text: string; x: number; y: number }> = [];
  const fontSize = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? String(FONT_PT));
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto' as CanvasFontKerning,
    measureText(text: string) {
      const size = fontSize();
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
    fillText(text: string, x: number, y: number) { calls.push({ text, x, y }); },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  return {
    canvas: {
      width: 0,
      height: 0,
      style: {} as Record<string, string>,
      getContext: () => ctx,
    } as unknown as HTMLCanvasElement,
    calls,
  };
}

function paragraph(text: string | readonly string[]): BodyElement {
  const pieces = typeof text === 'string' ? [text] : text;
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: { value: FONT_PT, rule: 'exact', explicit: true },
    numbering: null,
    tabStops: [],
    runs: pieces.map((piece) => ({
      type: 'text',
      text: piece,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      fontSize: FONT_PT,
      color: null,
      fontFamily: 'Synthetic CJK',
      isLink: false,
      background: null,
      vertAlign: null,
      hyperlink: null,
    })),
    defaultFontSize: FONT_PT,
    defaultFontFamily: 'Synthetic CJK',
    widowControl: false,
  } as unknown as BodyElement;
}

function document(
  characterSpacingControl?: string,
  text: string | readonly string[] = 'あ。い',
  pageWidth = 25,
): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth,
    pageHeight: 100,
    marginTop: 0,
    marginRight: 0,
    marginBottom: 0,
    marginLeft: 0,
    headerDistance: 0,
    footerDistance: 0,
    titlePage: false,
    evenAndOddHeaders: false,
  };
  return {
    section,
    body: [paragraph(text)],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    ...(characterSpacingControl ? { settings: { characterSpacingControl } } : {}),
  } as unknown as DocxDocumentModel;
}

async function paint(
  characterSpacingControl?: string,
  text?: string | readonly string[],
  pageWidth?: number,
) {
  const { canvas, calls } = recordingCanvas();
  await renderDocumentToCanvas(document(characterSpacingControl, text, pageWidth), canvas, 0, {
    width: pageWidth ?? 25,
    dpr: 1,
  });
  return calls.filter((call) => call.text.length > 0);
}

describe('ECMA-376 §17.15.1.18 characterSpacingControl', () => {
  it('applies compressPunctuation before line fit and paints the retained compressed geometry', async () => {
    const uncompressed = await paint();
    const compressed = await paint('compressPunctuation');

    expect(new Set(uncompressed.map((call) => call.y)).size).toBe(2);
    expect(new Set(compressed.map((call) => call.y)).size).toBe(1);

    const glyphs = compressed.filter((call) => /[あ。い]/u.test(call.text));
    expect(glyphs.map((call) => call.text)).toEqual(['あ。', 'い']);
    expect(glyphs.map((call) => call.x)).toEqual([0, 15]);
  });

  it('keeps explicit doNotCompress on the uncompressed line partition', async () => {
    const calls = await paint('doNotCompress');
    expect(new Set(calls.map((call) => call.y)).size).toBe(2);
  });

  it('uses the controlled advance while choosing a CJK overflow prefix', async () => {
    // The second run remains wider than the 15pt left on line 1, so this
    // exercises the CJK prefix search rather than the whole-segment fit check.
    // Its `い。` prefix fits exactly only after punctuation compression.
    const calls = await paint('compressPunctuation', ['あ', 'い。う'], 25);
    const glyphs = calls.filter((call) => /[あい。う]/u.test(call.text));

    expect(glyphs.map((call) => call.text).join('')).toBe('あい。う');
    expect(glyphs[1]?.y).toBe(glyphs[0]?.y);
    expect(glyphs.at(-1)?.y).not.toBe(glyphs[0]?.y);
  });
});

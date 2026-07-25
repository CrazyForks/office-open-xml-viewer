import { describe, expect, it } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types.js';

function recordingCanvas(): HTMLCanvasElement {
  let font = '16px sans-serif';
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => ({
      width: [...text].length * 8,
      fontBoundingBoxAscent: 12,
      fontBoundingBoxDescent: 4,
      actualBoundingBoxAscent: 12,
      actualBoundingBoxDescent: 4,
    }) as TextMetrics,
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, drawImage() {}, fillText() {}, strokeText() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1,
    lineCap: 'butt' as CanvasLineCap,
    lineJoin: 'miter' as CanvasLineJoin,
  };
  return {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  } as unknown as HTMLCanvasElement;
}

function textRun(text: string): DocxTextRun {
  return {
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 16,
    color: null,
    fontFamily: 'Identity Sans',
    isLink: false,
    background: null,
    vertAlign: null,
    hyperlink: null,
  };
}

function paragraph(text: string, paragraphId?: string): BodyElement {
  const value: DocParagraph = {
    ...(paragraphId ? { paragraphId } : {}),
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [{ type: 'text', ...textRun(text) } as DocParagraph['runs'][number]],
    defaultFontSize: 16,
    defaultFontFamily: 'Identity Sans',
    widowControl: false,
  };
  return { type: 'paragraph', ...value };
}

function document(body: BodyElement[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: 400,
      pageHeight: 400,
      marginTop: 0,
      marginRight: 0,
      marginBottom: 0,
      marginLeft: 0,
      headerDistance: 0,
      footerDistance: 0,
      titlePage: false,
      evenAndOddHeaders: false,
    } as SectionProps,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as DocxDocumentModel;
}

describe('DOCX text-run paragraph identity', () => {
  it('carries an authored w14:paraId through immutable layout to every callback run', async () => {
    const runs: DocxTextRunInfo[] = [];
    await renderDocumentToCanvas(document([
      paragraph('identified', '1A2B3C4D'),
      paragraph('positional only'),
    ]), recordingCanvas(), 0, {
      width: 400,
      dpr: 1,
      onTextRun: (run) => runs.push(run),
    });

    const identified = runs.filter((run) => run.text === 'identified');
    const positionalOnly = runs.filter((run) => run.text !== 'identified');
    expect(identified).not.toHaveLength(0);
    expect(identified.every((run) => run.paragraphId === '1A2B3C4D')).toBe(true);
    expect(positionalOnly.map((run) => run.text).join('')).toBe('positional only');
    expect(positionalOnly.every((run) => run.paragraphId === undefined)).toBe(true);
  });
});

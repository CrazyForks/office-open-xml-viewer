import { describe, expect, it } from 'vitest';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutLines, type LayoutSeg } from '../line-layout.js';
import type { DocxDocumentModel } from '../types.js';
import { verticalGlyphMeasurementServiceOf } from './runtime-state.js';

const model = {
  section: {},
  body: [],
  headers: { default: null, first: null, even: null },
  footers: { default: null, first: null, even: null },
  footnotes: [],
} as unknown as DocxDocumentModel;

interface MeasureCall { readonly font: string; readonly text: string }

function measurementContext(
  canvas: object,
  calls: MeasureCall[] = [],
): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    canvas,
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto' as CanvasFontKerning,
    measureText(text: string) {
      calls.push({ font, text });
      return {
        width: text.length * 5,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

describe('production measurement capability composition', () => {
  it('binds vertical glyph probing to the context selected by the layout path', () => {
    const calls: MeasureCall[] = [];
    const context = measurementContext({}, calls);
    const services = createLayoutServices(model, { measureContext: context });
    const segment = {
      text: '、。',
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      fontSize: 23,
      color: null,
      fontFamily: 'serif',
      vertAlign: null,
      measuredWidth: 0,
      verticalRun: true,
    } as LayoutSeg;

    layoutLines(
      context,
      [segment],
      200,
      0,
      1,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      verticalGlyphMeasurementServiceOf(services),
    );

    const runCall = calls.find(({ text }) => text === '、。');
    const glyphCall = calls.find(({ text }) => text === '、');
    expect(runCall?.font).toContain('23px');
    expect(glyphCall?.font).toBe(runCall?.font);
  });

  it('fingerprints DOM vertical probing separately from worker-safe measurement', () => {
    const canvasDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'HTMLCanvasElement');
    const documentDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'document');
    class TestHtmlCanvasElement {}

    try {
      Object.defineProperty(globalThis, 'HTMLCanvasElement', {
        configurable: true,
        value: TestHtmlCanvasElement,
      });
      Object.defineProperty(globalThis, 'document', {
        configurable: true,
        value: {},
      });

      const dom = createLayoutServices(model, {
        measureContext: measurementContext(new TestHtmlCanvasElement()),
      });
      const worker = createLayoutServices(model, {
        measureContext: measurementContext({}),
      });

      expect(verticalGlyphMeasurementServiceOf(dom).fingerprint)
        .toBe('vertical-glyph-measurement:dom-vert-probe-v1');
      expect(verticalGlyphMeasurementServiceOf(worker).fingerprint)
        .toBe('vertical-glyph-measurement:no-dom-vert-probe-v1');
      expect(dom.text.fingerprint).not.toBe(worker.text.fingerprint);
    } finally {
      if (canvasDescriptor) {
        Object.defineProperty(globalThis, 'HTMLCanvasElement', canvasDescriptor);
      } else {
        Reflect.deleteProperty(globalThis, 'HTMLCanvasElement');
      }
      if (documentDescriptor) {
        Object.defineProperty(globalThis, 'document', documentDescriptor);
      } else {
        Reflect.deleteProperty(globalThis, 'document');
      }
    }
  });
});

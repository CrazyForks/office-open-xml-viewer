import { describe, expect, it } from 'vitest';
import { createLayoutServices } from '../renderer.js';
import type { DocxDocumentModel } from '../types.js';
import { verticalGlyphMeasurementServiceOf } from './runtime-state.js';

const model = {
  section: {},
  body: [],
  headers: { default: null, first: null, even: null },
  footers: { default: null, first: null, even: null },
  footnotes: [],
} as unknown as DocxDocumentModel;

function measurementContext(canvas: object): CanvasRenderingContext2D {
  return {
    canvas,
    font: '10px serif',
    letterSpacing: '0px',
    fontKerning: 'auto' as CanvasFontKerning,
    measureText(text: string) {
      return {
        width: text.length * 5,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

describe('production measurement capability composition', () => {
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

import { describe, expect, it } from 'vitest';
import { buildSegments, layoutLines, rescaleLayoutLines, type LineLayoutEnvironment } from '../line-layout.js';
import { createLayoutServices } from '../renderer.js';
import type { DocRun, DocxDocumentModel } from '../types.js';
import type { TextLayoutService } from './text.js';
import { mathAstResourceKey } from './resources.js';
import { privateResourceLookupOf } from './runtime-state.js';

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '',
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 8,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
    }),
  } as unknown as CanvasRenderingContext2D;
}

function model(overrides: Partial<DocxDocumentModel> = {}): DocxDocumentModel {
  return {
    section: {
      pageWidth: 612, pageHeight: 792,
      marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    ...overrides,
  };
}

function textRun(text: string, extra: Record<string, unknown> = {}): DocRun {
  return {
    type: 'text', text,
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Authored Sans',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
    ...extra,
  } as DocRun;
}

describe('production layout service integration', () => {
  it('routes every normal run segmentation and measurement through the injected text service', () => {
    const base = createLayoutServices(model(), { measureContext: measureContext() });
    let calls = 0;
    const countingText: TextLayoutService = Object.freeze({
      ...base.text,
      shape(request: Parameters<TextLayoutService['shape']>[0]) {
        calls += 1;
        return base.text.shape(request);
      },
    });
    const services = Object.freeze({ ...base, text: countingText });
    const environment: LineLayoutEnvironment = { pageIndex: 0, totalPages: 1, layoutServices: services };
    const segments = buildSegments([textRun('first'), textRun('second')], environment);
    const afterSegmentation = calls;
    const lines = layoutLines(measureContext(), segments, 300, 0, 1);

    expect(afterSegmentation).toBeGreaterThanOrEqual(2);
    expect(calls).toBeGreaterThan(afterSegmentation);
    const afterLayout = calls;
    rescaleLayoutLines(lines, 2, measureContext(), {}, 0);
    expect(calls).toBeGreaterThan(afterLayout);
  });

  it('carries the w:kern threshold through the text service measure adapter', () => {
    let fontKerning: CanvasFontKerning = 'auto';
    const states: CanvasFontKerning[] = [];
    const ctx = {
      font: '',
      letterSpacing: '0px',
      get fontKerning() { return fontKerning; },
      set fontKerning(value: CanvasFontKerning) { fontKerning = value; },
      measureText(text: string) {
        states.push(fontKerning);
        const width = fontKerning === 'normal' ? 40 : fontKerning === 'none' ? 30 : 20;
        return {
          width: text ? width : 0,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const services = createLayoutServices(model(), { measureContext: ctx });
    const environment: LineLayoutEnvironment = { pageIndex: 0, totalPages: 1, layoutServices: services };
    const measure = (kerning: number) => {
      const segments = buildSegments([textRun('AV', { fontSize: 10, kerning })], environment);
      return layoutLines(ctx, segments, 300, 0, 1)[0].segments[0].measuredWidth;
    };

    expect(measure(5)).toBe(40);
    expect(measure(20)).toBe(30);
    expect(states).toContain('normal');
    expect(states).toContain('none');
    expect(ctx.fontKerning).toBe('auto');
  });

  it('inventories only successfully registered faces and labels Office replacements as substitutions', () => {
    const doc = model({
      majorFont: 'Calibri',
      embeddedFonts: [{ fontName: 'Broken Embedded', partPath: 'word/fonts/missing.odttf', fontKey: '', style: 'regular' }],
    });
    const failed = createLayoutServices(doc, {
      measureContext: measureContext(),
      embeddedFaces: [],
      googleFaces: [],
      useGoogleFonts: true,
    });
    const missingEmbedded = failed.text.shape({ text: 'x', fontSizePt: 10, fonts: { ascii: 'Broken Embedded' } });
    const missingGoogle = failed.text.shape({ text: 'x', fontSizePt: 10, fonts: { ascii: 'Calibri' } });
    expect(missingEmbedded.spans[0]?.font.source).toBe('generic');
    expect(missingEmbedded.diagnostics[0]?.message).toMatch(/unavailable/i);
    expect(missingGoogle.spans[0]?.font.source).toBe('generic');

    const carlito = { family: 'Carlito', weight: '400', style: 'normal', status: 'loaded' } as FontFace;
    const loaded = createLayoutServices(doc, {
      measureContext: measureContext(),
      embeddedFaces: [],
      googleFaces: [carlito],
      useGoogleFonts: true,
    });
    const substituted = loaded.text.shape({ text: 'x', fontSizePt: 10, fonts: { ascii: 'Calibri' } });
    expect(substituted.spans[0]?.font).toMatchObject({ source: 'substitute', resolvedFamily: 'Carlito' });
    expect(substituted.diagnostics[0]?.message).toMatch(/implementation-dependent/i);
  });

  it('requires loaded status and an exact family/weight/style match for every face', () => {
    const doc = model({
      embeddedFonts: [
        { fontName: 'Partial Embedded', partPath: 'word/fonts/regular.odttf', fontKey: '', style: 'regular' },
        { fontName: 'Partial Embedded', partPath: 'word/fonts/bold.odttf', fontKey: '', style: 'bold' },
      ],
    });
    const services = createLayoutServices(doc, {
      measureContext: measureContext(),
      embeddedFaces: [
        { family: '"Partial Embedded"', weight: '400', style: 'normal', status: 'loaded' },
        { family: 'Partial Embedded', weight: '700', style: 'normal', status: 'error' },
        { family: 'Timed Out', weight: '400', style: 'normal', status: 'loading' },
      ] as FontFace[],
    });
    const shape = (family: string, weight: number, style: 'normal' | 'italic' = 'normal') =>
      services.text.shape({ text: 'x', fontSizePt: 10, weight, style, fonts: { ascii: family } });

    expect(shape('Partial Embedded', 400).spans[0]?.font)
      .toMatchObject({ source: 'embedded', resolvedFamily: 'Partial Embedded' });
    expect(shape('Partial Embedded', 700).spans[0]?.font.source).toBe('generic');
    expect(shape('Partial Embedded', 400, 'italic').spans[0]?.font.source).toBe('generic');
    expect(shape('Timed Out', 400).spans[0]?.font.source).toBe('generic');
  });

  it('collects every currently representable math story, including nested tables', () => {
    const math = (value: string) => ({
      type: 'math', nodes: [{ type: 'text', text: value }], display: false, fontSize: 10,
    });
    const paragraph = (value: string) => ({ type: 'paragraph', runs: [math(value)] });
    const table = (value: string) => ({
      type: 'table', rows: [{ cells: [{ content: [paragraph(value)] }] }],
    });
    const doc = model({
      body: [paragraph('body')],
      headers: { default: { body: [table('header')] }, first: null, even: null },
      footers: { default: null, first: { body: [paragraph('footer')] }, even: null },
      footnotes: [{ id: '1', content: [table('footnote')] }],
      endnotes: [{ id: '2', content: [paragraph('endnote')] }],
    } as unknown as Partial<DocxDocumentModel>);
    const services = createLayoutServices(doc, { measureContext: measureContext() });

    for (const value of ['body', 'header', 'footer', 'footnote', 'endnote']) {
      const lookupKey = mathAstResourceKey({ nodes: [{ type: 'text', text: value }], display: false });
      expect(() => services.math.resolve(lookupKey), value).not.toThrow();
    }
  });

  it('requires runtime math handles to match available metadata exactly', () => {
    const available = {
      resourceKey: 'math:available', widthEm: 1, ascentEm: 0.8, descentEm: 0.2, diagnostics: [],
    };
    const unavailable = {
      resourceKey: 'math:unavailable', widthEm: 0, ascentEm: 0, descentEm: 0,
      available: false as const, diagnostics: [],
    };
    const drawable = {} as CanvasImageSource;

    expect(() => createLayoutServices(model(), { mathResources: [available], mathDrawables: new Map() }))
      .toThrow(/math.*membership|missing/i);
    expect(() => createLayoutServices(model(), {
      mathResources: [unavailable],
      mathDrawables: new Map([['math:unavailable', drawable]]),
    })).toThrow(/math.*membership|extra/i);

    const services = createLayoutServices(model(), {
      mathResources: [available, unavailable],
      mathDrawables: new Map([['math:available', drawable]]),
    });
    expect(privateResourceLookupOf(services)?.keys).toEqual(['math:available']);
  });

  it('gives main and worker factories identical fingerprints for identical successful snapshots', () => {
    const embedded = { family: 'Embedded', weight: '700', style: 'italic', status: 'loaded' } as FontFace;
    const options = {
      measureContext: measureContext(),
      embeddedFaces: [embedded],
      googleFaces: [] as FontFace[],
      localMetrics: { authored: { family: '__local_authored', lineHeightRatio: 1.25 } },
    };
    const main = createLayoutServices(model(), options);
    const worker = createLayoutServices(model(), { ...options, measureContext: measureContext() });

    expect(main.text.fingerprint).toBe(worker.text.fingerprint);
    expect(main.images.fingerprint).toBe(worker.images.fingerprint);
    expect(main.math.fingerprint).toBe(worker.math.fingerprint);
  });
});

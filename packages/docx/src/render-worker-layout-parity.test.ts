import { describe, expect, it } from 'vitest';
import type { DocxDocumentModel } from './types.js';
import {
  attachDocumentLayoutVariants,
  selectDocumentLayoutPage,
} from './layout/document-layout-variants.js';
import type { DocumentLayout, LayoutServices } from './layout/types.js';
import type { RenderWorkerRequest, RenderWorkerResponse } from './worker-protocol.js';
import { createLayoutServices, layoutDocument } from './renderer.js';
import { DocxDocument } from './document.js';

function services(): LayoutServices {
  return Object.freeze({
    text: { fingerprint: 'text:equal' },
    images: { fingerprint: 'images:equal' },
    math: { fingerprint: 'math:equal' },
  }) as LayoutServices;
}

function layout(currentDateMs: number): DocumentLayout {
  const variant = currentDateMs === 10 ? 'default' : 'dated';
  const count = variant === 'default' ? 1 : 2;
  return {
    pages: Array.from({ length: count }, (_, pageIndex) => ({
      pageIndex,
      geometry: {
        widthPt: variant === 'default' ? 612 : 792,
        heightPt: variant === 'default' ? 792 : 612,
      },
      variant,
    }) as never),
    diagnostics: [],
  };
}

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '', letterSpacing: '0px', fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 8,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
    }),
  } as unknown as CanvasRenderingContext2D;
}

function realModel(): DocxDocumentModel {
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
  };
}

describe('render worker canonical layout parity', () => {
  it('dispatches variant page validation to the worker for render and collect', async () => {
    const requests: RenderWorkerRequest[] = [];
    const bitmap = {} as ImageBitmap;
    const document = Object.create(DocxDocument.prototype) as DocxDocument;
    Object.assign(document, {
      _mode: 'worker',
      _meta: {
        pageCount: 1, comments: [], footnotes: [], endnotes: [],
        pageSizes: [{ widthPt: 612, heightPt: 792 }], bookmarkPages: [],
      },
      _bridge: {
        request: async (factory: (id: number) => RenderWorkerRequest) => {
          const request = factory(9);
          requests.push(request);
          return request.type === 'renderPage'
            ? { type: 'pageRendered', id: 9, bitmap, runs: [] }
            : { type: 'runsCollected', id: 9, runs: [] };
        },
      },
    });

    await expect(document.renderPageToBitmap(1, { currentDate: 20 })).resolves.toBe(bitmap);
    await expect(document.collectPageRuns(1, { currentDate: 20 })).resolves.toEqual([]);
    expect(requests.map(({ type }) => type)).toEqual(['renderPage', 'collectRuns']);
    expect(requests.map((request) => 'pageIndex' in request ? request.pageIndex : null))
      .toEqual([1, 1]);
  });

  it('uses the real canonical builder with equal main and worker service snapshots', () => {
    const model = realModel();
    const mainServices = createLayoutServices(model, { measureContext: measureContext() });
    const workerServices = createLayoutServices(model, { measureContext: measureContext() });
    attachDocumentLayoutVariants({
      model, services: mainServices, defaultCurrentDateMs: 10,
      buildLayout: (options) => layoutDocument(model, mainServices, options),
    });
    attachDocumentLayoutVariants({
      model, services: workerServices, defaultCurrentDateMs: 10,
      buildLayout: (options) => layoutDocument(model, workerServices, options),
    });

    const main = selectDocumentLayoutPage(mainServices, {
      currentDate: undefined, defaultCurrentDateMs: 10,
    }, 0);
    const worker = selectDocumentLayoutPage(workerServices, {
      currentDate: undefined, defaultCurrentDateMs: 10,
    }, 0);

    expect(worker.key).toBe(main.key);
    expect(worker.options).toEqual(main.options);
    expect(worker.layout.pages.length).toBe(main.layout.pages.length);
    expect(worker.page.geometry).toEqual(main.page.geometry);
    expect(workerServices.text.fingerprint).toBe(mainServices.text.fingerprint);
    expect(workerServices.images.fingerprint).toBe(mainServices.images.fingerprint);
    expect(workerServices.math.fingerprint).toBe(mainServices.math.fingerprint);
  });

  it('selects equal keys, fingerprints, page counts, sizes, and variants from equal normalized inputs', () => {
    const model = {
      body: [], section: { pageWidth: 612, pageHeight: 792 },
    } as unknown as DocxDocumentModel;
    const mainServices = services();
    const workerServices = services();
    let mainBuilds = 0;
    let workerBuilds = 0;
    attachDocumentLayoutVariants({
      model, services: mainServices, defaultCurrentDateMs: 10,
      buildLayout: (options) => { mainBuilds += 1; return layout(options.currentDateMs); },
    });
    attachDocumentLayoutVariants({
      model, services: workerServices, defaultCurrentDateMs: 10,
      buildLayout: (options) => { workerBuilds += 1; return layout(options.currentDateMs); },
    });

    const main = selectDocumentLayoutPage(mainServices, {
      currentDate: new Date(20), defaultCurrentDateMs: 10,
    }, 1);
    const worker = selectDocumentLayoutPage(workerServices, {
      currentDate: 20, defaultCurrentDateMs: 10,
    }, 1);

    expect(worker.options).toEqual(main.options);
    expect(worker.key).toBe(main.key);
    expect(worker.layout.pages.length).toBe(main.layout.pages.length);
    expect(worker.page.geometry).toEqual(main.page.geometry);
    expect((worker.page as unknown as { variant: string }).variant).toBe(
      (main.page as unknown as { variant: string }).variant,
    );
    expect(main.layout).not.toBe(worker.layout);
    expect([mainBuilds, workerBuilds]).toEqual([1, 1]);

    expect(selectDocumentLayoutPage(mainServices, {
      currentDate: 10, defaultCurrentDateMs: 10,
    }, 0).key).not.toBe(main.key);
    expect([mainBuilds, workerBuilds]).toEqual([2, 1]);
  });

  it('keeps worker request and response protocol shapes unchanged', () => {
    const parse = {
      type: 'parse', id: 1, data: new ArrayBuffer(0), useGoogleFonts: false,
      defaultCurrentDateMs: 10,
    } satisfies RenderWorkerRequest;
    const render = {
      type: 'renderPage', id: 2, pageIndex: 0, opts: { currentDate: 20 },
    } satisfies RenderWorkerRequest;
    const collect = {
      type: 'collectRuns', id: 3, pageIndex: 0, opts: { currentDate: 20 },
    } satisfies RenderWorkerRequest;
    const parsed = {
      type: 'parsedMeta', id: 1,
      meta: { pageCount: 1, comments: [], footnotes: [], endnotes: [], pageSizes: [], bookmarkPages: [] },
    } satisfies RenderWorkerResponse;
    const pageRendered = {
      type: 'pageRendered', id: 2, bitmap: {} as ImageBitmap, runs: [],
    } satisfies RenderWorkerResponse;
    const runsCollected = {
      type: 'runsCollected', id: 3, runs: [],
    } satisfies RenderWorkerResponse;

    expect(Object.keys(parse).sort()).toEqual([
      'data', 'defaultCurrentDateMs', 'id', 'type', 'useGoogleFonts',
    ]);
    expect(Object.keys(render).sort()).toEqual(['id', 'opts', 'pageIndex', 'type']);
    expect(Object.keys(collect).sort()).toEqual(['id', 'opts', 'pageIndex', 'type']);
    expect(Object.keys(parsed).sort()).toEqual(['id', 'meta', 'type']);
    expect(Object.keys(pageRendered).sort()).toEqual(['bitmap', 'id', 'runs', 'type']);
    expect(Object.keys(runsCollected).sort()).toEqual(['id', 'runs', 'type']);
  });
});

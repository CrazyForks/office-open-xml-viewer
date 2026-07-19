import { describe, expect, it } from 'vitest';
import type { BodyElement, DocxDocumentModel, SectionProps } from './types.js';
import {
  attachDocumentLayoutVariants,
  selectDocumentLayoutPage,
} from './layout/document-layout-variants.js';
import { layoutVariantStoreOf } from './layout/runtime-state.js';
import type { DeepReadonly, DocumentLayout, LayoutServices } from './layout/types.js';
import type {
  DocumentMeta,
  RenderWorkerRequest,
  RenderWorkerResponse,
} from './worker-protocol.js';
import { createLayoutServices, layoutDocument } from './renderer.js';
import { DocxDocument } from './document.js';
import { retainRenderWorkerDocumentLayout } from './render-worker-layout.js';
import { stableFingerprint } from './layout/fingerprint.js';
import { buildBookmarkPageMap } from './bookmark-nav.js';

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
    body: [{
      type: 'paragraph',
      alignment: 'left',
      indentLeft: 0,
      indentRight: 0,
      indentFirst: 0,
      spaceBefore: 0,
      spaceAfter: 0,
      lineSpacing: null,
      numbering: null,
      tabStops: [],
      defaultFontSize: 10,
      defaultFontFamily: 'Times New Roman',
      widowControl: false,
      runs: [{
        type: 'field',
        fieldType: 'date',
        instruction: ' DATE \\@ "yyyy" ',
        fallbackText: 'cached',
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 10,
        color: null,
        fontFamily: 'Times New Roman',
        background: null,
        vertAlign: null,
      }],
    }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  };
}

interface PrivateSectionPlacementWire {
  readonly sectionId: string;
  readonly sectionBidi: boolean;
}

type PrivateSectionProps = SectionProps & {
  readonly __sectionPlacement: PrivateSectionPlacementWire;
};

function finalNextColumnRtlModel(): DocxDocumentModel {
  const base = realModel();
  const columns = {
    count: 2,
    spacePt: 20,
    equalWidth: true,
    sep: false,
    cols: [],
  };
  const endingSection = {
    type: 'sectionBreak',
    kind: 'nextPage',
    geom: base.section,
    columns,
    textDirection: null,
    pageNumType: null,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    titlePage: false,
    __sectionPlacement: {
      sectionId: 'section:outgoing',
      sectionBidi: true,
    },
  } as unknown as BodyElement;
  const finalSection = {
    ...base.section,
    sectionStart: 'nextColumn',
    columns,
    __sectionPlacement: {
      sectionId: 'section:final',
      sectionBidi: true,
    },
  } as PrivateSectionProps;

  return {
    ...base,
    body: [base.body[0]!, endingSection, structuredClone(base.body[0]!)],
    section: finalSection,
  };
}

function metadataForDefaultLayout(
  model: DocxDocumentModel,
  layout: DeepReadonly<DocumentLayout>,
): DocumentMeta {
  return {
    pageCount: layout.pages.length,
    comments: model.comments ?? [],
    footnotes: model.footnotes ?? [],
    endnotes: model.endnotes ?? [],
    pageSizes: layout.pages.map((page) => ({
      widthPt: page.geometry.widthPt,
      heightPt: page.geometry.heightPt,
    })),
    bookmarkPages: [...buildBookmarkPageMap(layout)],
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

  it('retains complete default and dated layouts through the production worker wiring seam', () => {
    const model = realModel();
    const mainServices = createLayoutServices(model, { measureContext: measureContext() });
    const mainVariants = attachDocumentLayoutVariants({
      model,
      services: mainServices,
      defaultCurrentDateMs: 10,
      buildLayout: (options) => layoutDocument(model, mainServices, options),
    });
    const workerServices = createLayoutServices(model, { measureContext: measureContext() });
    const workerState = retainRenderWorkerDocumentLayout(
      model,
      workerServices,
      10,
    );

    expect(Object.keys(workerState).sort()).toEqual([
      'defaultCurrentDateMs',
      'layoutServices',
      'layoutVariants',
      'model',
    ]);
    expect(workerState.model).toBe(model);
    expect(workerState.layoutServices).toBe(workerServices);
    expect(layoutVariantStoreOf(workerServices)).toBe(workerState.layoutVariants);

    const fingerprints = (currentDate: number | undefined) => {
      const mainInput = { currentDate, defaultCurrentDateMs: 10 };
      const workerInput = {
        currentDate,
        defaultCurrentDateMs: workerState.defaultCurrentDateMs,
      };
      const main = selectDocumentLayoutPage(mainServices, mainInput, 0);
      const worker = selectDocumentLayoutPage(
        workerState.layoutServices,
        workerInput,
        0,
      );
      expect(worker.key).toBe(main.key);
      expect(worker.options).toEqual(main.options);
      expect(worker.layout.pages[0]?.layers.body.length).toBeGreaterThan(0);
      return {
        key: worker.key,
        main: stableFingerprint('document-layout', main.layout),
        worker: stableFingerprint('document-layout', worker.layout),
      };
    };

    const defaultVariant = fingerprints(undefined);
    const datedVariant = fingerprints(Date.UTC(2222, 0, 1));
    expect(selectDocumentLayoutPage(workerState.layoutServices, {
      currentDate: undefined,
      defaultCurrentDateMs: workerState.defaultCurrentDateMs,
    }, 0).layout).toBe(workerState.layoutVariants.defaultLayout);
    expect(stableFingerprint(
      'document-metadata',
      metadataForDefaultLayout(model, workerState.layoutVariants.defaultLayout),
    )).toBe(stableFingerprint(
      'document-metadata',
      metadataForDefaultLayout(model, mainVariants.store.defaultLayout),
    ));
    expect(defaultVariant.worker).toBe(defaultVariant.main);
    expect(datedVariant.worker).toBe(datedVariant.main);
    expect(datedVariant.key).not.toBe(defaultVariant.key);
  });

  it('retains parser-private final RTL nextColumn layout identically in main and worker', () => {
    const model = finalNextColumnRtlModel();
    const mainServices = createLayoutServices(model, { measureContext: measureContext() });
    attachDocumentLayoutVariants({
      model,
      services: mainServices,
      defaultCurrentDateMs: 10,
      buildLayout: (options) => layoutDocument(model, mainServices, options),
    });
    const workerState = retainRenderWorkerDocumentLayout(
      model,
      createLayoutServices(model, { measureContext: measureContext() }),
      10,
    );
    const main = selectDocumentLayoutPage(mainServices, {
      currentDate: undefined,
      defaultCurrentDateMs: 10,
    }, 0);
    const worker = selectDocumentLayoutPage(workerState.layoutServices, {
      currentDate: undefined,
      defaultCurrentDateMs: workerState.defaultCurrentDateMs,
    }, 0);

    expect(stableFingerprint('document-layout', worker.layout)).toBe(
      stableFingerprint('document-layout', main.layout),
    );
    expect(main.layout.pages).toHaveLength(1);
    expect(main.page.sectionRegions.map((region) => ({
      section: region.sectionOccurrenceId,
      direction: region.columnFlowDirection,
      columns: region.columnIndexes,
    }))).toEqual([
      { section: 'section:outgoing', direction: 'rtl', columns: [1] },
      { section: 'section:final', direction: 'rtl', columns: [0] },
    ]);
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

  it('pins worker request and response protocol shapes including structured failures', () => {
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
    const error = {
      type: 'error', id: 4, message: 'unsupported transition',
      errorName: 'UnsupportedPageFlowTransitionError',
      code: 'NEXT_COLUMN_DESTINATION_UNAVAILABLE',
      reason: 'physical-column',
      outgoingColumnIndex: 0,
      outgoingColumnCount: 3,
      incomingColumnCount: 1,
    } satisfies RenderWorkerResponse;

    expect(Object.keys(parse).sort()).toEqual([
      'data', 'defaultCurrentDateMs', 'id', 'type', 'useGoogleFonts',
    ]);
    expect(Object.keys(render).sort()).toEqual(['id', 'opts', 'pageIndex', 'type']);
    expect(Object.keys(collect).sort()).toEqual(['id', 'opts', 'pageIndex', 'type']);
    expect(Object.keys(parsed).sort()).toEqual(['id', 'meta', 'type']);
    expect(Object.keys(pageRendered).sort()).toEqual(['bitmap', 'id', 'runs', 'type']);
    expect(Object.keys(runsCollected).sort()).toEqual(['id', 'runs', 'type']);
    expect(Object.keys(error).sort()).toEqual([
      'code',
      'errorName',
      'id',
      'incomingColumnCount',
      'message',
      'outgoingColumnCount',
      'outgoingColumnIndex',
      'reason',
      'type',
    ]);
  });
});

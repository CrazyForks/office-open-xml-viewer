import { afterEach, describe, expect, it, vi } from 'vitest';
import { buildPageLayers } from '../layout/page-graph.js';
import type { DocumentLayout, LayoutPage, PaintResourceRegistry } from '../layout/types.js';
import { renderSelectedDocumentPage } from './canvas-document.js';

class RecordingContext {
  readonly operations: string[] = [];
  fillStyle = '';
  globalAlpha = 1;

  scale(): void { this.operations.push('scale'); }
  fillRect(): void { this.operations.push('fillRect'); }
  save(): void { this.operations.push('save'); }
  restore(): void { this.operations.push('restore'); }
  drawImage(): void { this.operations.push('drawImage'); }
}

class ElementCanvas {
  width = 1;
  height = 1;
  readonly style: Record<string, string> = {};
  readonly context = new RecordingContext();

  getContext(): RecordingContext { return this.context; }
}

class WorkerCanvas {
  width = 1;
  height = 1;
  readonly context = new RecordingContext();

  getContext(): RecordingContext { return this.context; }
}

const section: LayoutPage['section'] = {
  geometry: {
    pageWidth: 200, pageHeight: 100,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0,
  },
  columns: [{ xPt: 0, wPt: 200 }],
  columnSeparator: false,
  grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
  textDirection: 'tbRl', verticalAlignment: 'top',
};

const page: LayoutPage = {
  pageIndex: 0,
  geometry: {
    xPt: 0, yPt: 0, widthPt: 200, heightPt: 100,
    contentTopPt: 0, contentBottomPt: 100,
  },
  flowDomains: [], section, sectionOccurrenceId: 'section:0', parityBlank: false,
  bookmarkStarts: [],
  pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
  sectionRegions: [], columnSeparators: [], pageBorder: null,
  layers: {
    ...buildPageLayers([]),
    capabilities: { requiresElementBackedVerticalGlyphPaint: true },
  },
  readingOrder: [],
};

const layout: DocumentLayout = { pages: [page], diagnostics: [] };
const registry: PaintResourceRegistry = {
  keys: [], descriptors: [],
  resolve() { throw new Error('empty registry'); },
};

afterEach(() => vi.unstubAllGlobals());

describe('vertical OpenType paint target projection', () => {
  it('paints into an element-backed surface before copying to an OffscreenCanvas target', async () => {
    const created: ElementCanvas[] = [];
    vi.stubGlobal('HTMLCanvasElement', ElementCanvas);
    vi.stubGlobal('document', {
      createElement(tag: string) {
        if (tag !== 'canvas') throw new Error(`unexpected element ${tag}`);
        const canvas = new ElementCanvas();
        created.push(canvas);
        return canvas;
      },
    });
    const target = new WorkerCanvas();

    await renderSelectedDocumentPage(
      layout,
      page,
      target as unknown as OffscreenCanvas,
      {
        dpr: 1, parseError: false, registry, textRuns: [],
      },
    );

    expect(created).toHaveLength(1);
    expect(created[0]!.context.operations).toContain('fillRect');
    expect(target.context.operations).toEqual(['drawImage']);
    expect({ width: target.width, height: target.height }).toEqual({
      width: created[0]!.width,
      height: created[0]!.height,
    });
  });
});

import { describe, expect, it, vi } from 'vitest';
import { paintLayoutPageContent } from './canvas-page.js';
import type { CanvasPaintContext } from './types.js';
import type { LayoutPage } from '../layout/types.js';

describe('retained canonical body paint', () => {
  it('does not initialize the surface or measure while painting an empty retained body', () => {
    const rawContext = {
      measureText: vi.fn(() => { throw new Error('paint measured text'); }),
      clearRect: vi.fn(),
      setTransform: vi.fn(),
    };
    const ctx = rawContext as unknown as CanvasPaintContext['ctx'];
    const page = {
      pageIndex: 0,
      geometry: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 100, contentTopPt: 10, contentBottomPt: 90 },
      flowDomains: [],
      section: {
        geometry: { pageWidth: 100, pageHeight: 100, marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10, headerDistance: 5, footerDistance: 5 },
        columns: [{ xPt: 10, wPt: 80 }], grid: { kind: 'none', linePitchPt: null, charSpacePt: null }, textDirection: 'lrTb', verticalAlignment: 'top',
      },
      sectionOccurrenceId: 'section:0', parityBlank: false, bookmarkStarts: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
      sectionRegions: [],
      pageBorders: null,
      layers: { paintSequence: [], background: [], behindText: [], header: [], body: [], notes: [], front: [], footer: [] },
      readingOrder: [],
    } satisfies LayoutPage;

    paintLayoutPageContent(page, { ctx, scale: 1, dpr: 1, resources: { paint: vi.fn() } });

    expect(rawContext.measureText).not.toHaveBeenCalled();
    expect(rawContext.clearRect).not.toHaveBeenCalled();
    expect(rawContext.setTransform).not.toHaveBeenCalled();
  });
});

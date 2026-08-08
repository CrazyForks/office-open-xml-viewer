import { afterEach, describe, expect, it, vi } from 'vitest';
import type { PptxElementSelectionContext } from './element-selection.js';
import {
  FakePptxEngine,
  installDom,
  makeBorrowedPptxScrollViewer,
  makeContainer,
  type FakeEl,
} from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function context(): PptxElementSelectionContext {
  return {
    format: 'pptx', kind: 'element', slideIndex: 0, elementIndex: 0,
    origin: 'layout', elementType: 'shape', point: { x: 0, y: 0 },
    bounds: { x: 0, y: 0, width: 100, height: 50, rotation: 0, flipH: false, flipV: false },
    shapeId: '9', geometry: 'rect', truncated: false, truncationReasons: [],
    textCharacters: 0, maxTextCharacters: 16_384,
  };
}

describe('PptxScrollViewer selection context', () => {
  it('identifies the clicked mounted slide and clears focus on the desk', async () => {
    installDom();
    const container = makeContainer(960, 720);
    const engine = new FakePptxEngine(1, 9_144_000, 6_858_000);
    engine.elementContext = context();
    const onSelectionContextChange = vi.fn();
    const viewer = makeBorrowedPptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableElementSelection: true,
      onSelectionContextChange,
    });
    const internals = viewer as unknown as {
      _scrollHost: FakeEl;
      _slots: Map<number, { wrapper: FakeEl; canvas: FakeEl }>;
    };
    const slot = internals._slots.get(0) as { wrapper: FakeEl; canvas: FakeEl };
    slot.canvas.clientWidth = 960;
    slot.canvas.clientHeight = 720;

    internals._scrollHost.dispatch('click', {
      target: slot.canvas, button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
    });
    await Promise.resolve();

    expect(engine.elementContextCalls).toEqual([{
      slideIndex: 0,
      point: { x: 9_144_000 / 2, y: 6_858_000 / 2 },
      options: { tolerance: 6 / 960 * 9_144_000, maxTextCharacters: 65_536 },
    }]);
    expect(viewer.getSelectionContext()).toMatchObject({
      kind: 'element', shapeId: '9', origin: 'layout',
    });

    internals._scrollHost.dispatch('click', {
      target: internals._scrollHost, button: 0, clientX: 0, clientY: 0, defaultPrevented: false,
    });
    expect(viewer.getSelectionContext()).toBeNull();
    expect(onSelectionContextChange).toHaveBeenLastCalledWith(null);

    viewer.destroy();
    expect(() => viewer.getSelectionContext()).toThrow('PptxScrollViewer is destroyed');
  });
});

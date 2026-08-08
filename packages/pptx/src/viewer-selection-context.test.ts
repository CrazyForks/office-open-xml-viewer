import { afterEach, describe, expect, it, vi } from 'vitest';
import { PptxPresentation } from './presentation.js';
import { PptxViewer } from './viewer.js';
import type { PptxElementSelectionContext } from './element-selection.js';
import { FakePptxEngine, installDom, makeEl, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SLIDE_WIDTH = 9_144_000;
const SLIDE_HEIGHT = 6_858_000;

function elementContext(shapeId: string): PptxElementSelectionContext {
  return {
    format: 'pptx', kind: 'element', slideIndex: 0, elementIndex: 0,
    origin: 'slide', elementType: 'shape', point: { x: 0, y: 0 },
    bounds: {
      x: 0, y: 0, width: 1_000_000, height: 1_000_000,
      rotation: 0, flipH: false, flipV: false,
    },
    shapeId, geometry: 'rect', truncated: false, truncationReasons: [],
    textCharacters: 0, maxTextCharacters: 16_384,
  };
}

async function mount(mode: 'main' | 'worker' = 'main') {
  installDom();
  const canvas = makeEl('canvas');
  canvas.clientWidth = 960;
  canvas.clientHeight = 720;
  const engine = new FakePptxEngine(1, SLIDE_WIDTH, SLIDE_HEIGHT, mode);
  const onSelectionContextChange = vi.fn();
  vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
  const viewer = new PptxViewer(canvas as unknown as HTMLCanvasElement, {
    mode,
    enableElementSelection: true,
    onSelectionContextChange,
  });
  await viewer.load('deck.pptx');
  return {
    canvas,
    engine,
    onSelectionContextChange,
    viewer,
    wrapper: canvas.parentElement as FakeEl,
  };
}

describe('PptxViewer selection context', () => {
  it('validates the configurable CSS-pixel line tolerance', () => {
    installDom();
    const canvas = makeEl('canvas');
    expect(() => new PptxViewer(canvas as unknown as HTMLCanvasElement, {
      enableElementSelection: true,
      elementHitTolerance: Number.NaN,
    })).toThrow(/elementHitTolerance/);
  });

  it.each(['main', 'worker'] as const)(
    'maps a click to slide EMU and emits a detached compact context in %s mode',
    async (mode) => {
      const mounted = await mount(mode);
      mounted.engine.elementContext = elementContext('7');

      mounted.wrapper.dispatch('click', {
        button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
      });
      await Promise.resolve();

      expect(mounted.engine.elementContextCalls).toEqual([{
        slideIndex: 0,
        point: { x: SLIDE_WIDTH / 2, y: SLIDE_HEIGHT / 2 },
        options: { tolerance: 6 / 960 * SLIDE_WIDTH, maxTextCharacters: 65_536 },
      }]);
      expect(mounted.onSelectionContextChange).toHaveBeenCalledWith(
        expect.objectContaining({ format: 'pptx', kind: 'element', shapeId: '7' }),
      );
      const snapshot = mounted.viewer.getSelectionContext();
      expect(snapshot).toMatchObject({
        shapeId: '7', point: { x: SLIDE_WIDTH / 2, y: SLIDE_HEIGHT / 2 },
      });
      expect(snapshot).not.toBe(mounted.engine.elementContext);
      mounted.viewer.destroy();
    },
  );

  it('does not let a stale hit-test response overwrite a later click', async () => {
    const mounted = await mount();
    const resolvers: Array<(value: PptxElementSelectionContext | null) => void> = [];
    mounted.engine.getElementContextAt = vi.fn(
      (): Promise<PptxElementSelectionContext | null> =>
        new Promise((resolve) => resolvers.push(resolve)),
    );

    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 100, clientY: 100, defaultPrevented: false,
    });
    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 200, clientY: 200, defaultPrevented: false,
    });
    resolvers[1](elementContext('new'));
    await Promise.resolve();
    resolvers[0](elementContext('stale'));
    await Promise.resolve();

    expect(mounted.viewer.getSelectionContext()).toMatchObject({ shapeId: 'new' });
    expect(mounted.onSelectionContextChange).toHaveBeenCalledTimes(1);
    mounted.viewer.destroy();
  });

  it('ignores prevented clicks and closes the query/listener surface on destroy', async () => {
    const mounted = await mount();
    mounted.engine.elementContext = elementContext('7');
    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 480, clientY: 360, defaultPrevented: true,
    });
    await Promise.resolve();
    expect(mounted.engine.elementContextCalls).toEqual([]);

    mounted.viewer.destroy();
    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
    });
    expect(mounted.engine.elementContextCalls).toEqual([]);
    expect(() => mounted.viewer.getSelectionContext()).toThrow('PptxViewer is destroyed');
  });
});

import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxScrollViewer } from './scroll-viewer.js';
import { installDom, makeContainer, FakePptxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('PptxScrollViewer — skeleton (T1)', () => {
  it('builds the wrapper → scrollHost → spacer DOM inside the container', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakePptxEngine(3, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { presentation: engine.asPres() });
    // container → wrapper
    const wrapper = container.children[0];
    expect(wrapper.tag).toBe('div');
    expect(wrapper.style.position).toBe('relative');
    // wrapper → scrollHost
    const scrollHost = wrapper.children[0];
    expect(scrollHost.style.overflow).toBe('auto');
    // scrollHost → spacer
    const spacer = scrollHost.children[0];
    expect(spacer.style.position).toBe('absolute');
    v.destroy();
  });

  it('exposes slideCount from the injected engine', () => {
    installDom();
    const engine = new FakePptxEngine(5, 1000, 600);
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, { presentation: engine.asPres() });
    expect(v.slideCount).toBe(5);
    v.destroy();
  });

  it('load() is unsupported when an engine is injected', async () => {
    installDom();
    const engine = new FakePptxEngine(1, 1000, 600);
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, { presentation: engine.asPres() });
    await expect(v.load('x.pptx')).rejects.toThrow(/injected/i);
    v.destroy();
  });

  it('destroy() removes the DOM and does NOT destroy an injected engine', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakePptxEngine(1, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { presentation: engine.asPres() });
    expect(container.children.length).toBe(1); // wrapper mounted
    v.destroy();
    expect(container.children.length).toBe(0); // wrapper removed
    expect(engine.destroyed).toBe(false); // injected engine preserved (caller owns it)
  });

  it('slideCount is 0 before load resolves (no injected engine)', () => {
    installDom();
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {});
    expect(v.slideCount).toBe(0);
    v.destroy();
  });

  // O1 (design §11): an injected engine's own `mode` is authoritative. An
  // EXPLICITLY conflicting opts.mode is a mis-configuration rejected at
  // construction; a matching or absent opts.mode constructs fine.
  it('throws when opts.mode conflicts with an injected worker-mode engine', () => {
    installDom();
    const engine = new FakePptxEngine(1, 1000, 600, 'worker');
    expect(
      () =>
        new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {
          presentation: engine.asPres(),
          mode: 'main',
        }),
    ).toThrow(/mode/i);
  });

  it('does NOT throw when opts.mode matches an injected worker-mode engine', () => {
    installDom();
    const engine = new FakePptxEngine(1, 1000, 600, 'worker');
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {
      presentation: engine.asPres(),
      mode: 'worker',
    });
    expect(v.slideCount).toBe(1);
    v.destroy();
    // Injected engine is caller-owned even in the worker case: destroy() leaves it intact.
    expect(engine.destroyed).toBe(false);
  });

  it('constructs a default-main injected engine with absent opts.mode (load still rejects; destroy preserves engine)', async () => {
    installDom();
    // Default mode is 'main'; opts.mode is absent ⇒ no conflict, resolved path is main.
    const engine = new FakePptxEngine(2, 1000, 600);
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {
      presentation: engine.asPres(),
    });
    expect(v.slideCount).toBe(2);
    await expect(v.load('x.pptx')).rejects.toThrow(/injected/i);
    v.destroy();
    // Injected engine is caller-owned: destroy() must not tear it down.
    expect(engine.destroyed).toBe(false);
  });
});

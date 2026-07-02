import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { installDom, makeContainer, type FakeDocument, type FakeEl } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * XlsxViewer builds its whole UI subtree (a wrapper div holding the canvas
 * area + sheet-tab bar) inside the caller's container, and injects a `<style>`
 * into `document.head`. destroy() must (1) remove that subtree so the container
 * returns to the empty state it had before construction, and (2) NOT leak a new
 * <style> per instance. These tests pin both, plus removal of the document-level
 * keydown listener.
 */
describe('XlsxViewer.destroy() — subtree + listeners + style', () => {
  it('empties the container (removes the wrapper subtree)', () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    // Construction mounted exactly one wrapper subtree.
    expect(container.childNodes.length).toBe(1);
    v.destroy();
    expect(container.childNodes.length).toBe(0);
  });

  it('injects the viewer <style> once across 3 mount/unmount cycles (module-level, tagged)', () => {
    const doc = installDom() as FakeDocument;
    const container = makeContainer();
    for (let i = 0; i < 3; i++) {
      const v = new XlsxViewer(container as unknown as HTMLElement);
      v.destroy();
    }
    // Exactly one tagged stylesheet survives in <head> — not three (and not zero:
    // it is a class-constant sheet, kept after destroy for any live instances).
    const styles = doc.head.children.filter(
      (c: FakeEl) => c.tag === 'style' && c.hasAttribute('data-xlsx-viewer-styles'),
    );
    expect(styles.length).toBe(1);
    // And it is still present after the last destroy (destroy must NOT remove it).
    expect(doc.head.querySelector('style[data-xlsx-viewer-styles]')).not.toBeNull();
  });

  it('keeps a single tagged stylesheet even while several viewers are alive at once', () => {
    const doc = installDom() as FakeDocument;
    const a = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const b = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const c = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const count = () =>
      doc.head.children.filter(
        (e: FakeEl) => e.tag === 'style' && e.hasAttribute('data-xlsx-viewer-styles'),
      ).length;
    expect(count()).toBe(1);
    a.destroy();
    // b and c are still alive — the shared sheet must remain.
    expect(count()).toBe(1);
    expect(doc.head.querySelector('style[data-xlsx-viewer-styles]')).not.toBeNull();
    b.destroy();
    c.destroy();
  });

  it('removes the document-level keydown listener on destroy', () => {
    const doc = installDom() as FakeDocument;
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    // The constructor registered exactly one keydown handler on document.
    expect(doc.listenerCount('keydown')).toBe(1);
    v.destroy();
    // destroy() detached it — dispatching now reaches no viewer handler.
    expect(doc.listenerCount('keydown')).toBe(0);
    // Dispatching a keydown after destroy must not throw (no live handler).
    expect(() => doc.dispatchEvent('keydown', { key: 'c', ctrlKey: true })).not.toThrow();
  });

  it('is safe to call destroy() twice', () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    v.destroy();
    expect(() => v.destroy()).not.toThrow();
    expect(container.childNodes.length).toBe(0);
  });
});

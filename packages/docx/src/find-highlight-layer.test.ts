import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  buildDocxHighlightLayer,
  DEFAULT_FIND_HIGHLIGHT,
  DEFAULT_FIND_ACTIVE_HIGHLIGHT,
  type DocxHighlightMatch,
} from './find-highlight-layer.js';
import type { DocxTextRunInfo } from './renderer';

/**
 * IX2 docx highlight overlay. Same node-env fake-DOM stub as the text-layer
 * tests (no jsdom in this workspace). buildDocxHighlightLayer draws one absolute
 * box per matched run-slice, positioned from the run's x/y/h plus the slice's
 * measured horizontal extent, with the active match in the emphasis colour.
 */
interface FakeEl {
  tag: string;
  innerHTML: string;
  style: Record<string, string> & { cssText: string };
  children: FakeEl[];
  appendChild(c: FakeEl): void;
}

function makeEl(tag: string): FakeEl {
  const style: Record<string, string> = {};
  const el: FakeEl = {
    tag,
    innerHTML: '',
    children: [],
    style: new Proxy(style as Record<string, string> & { cssText: string }, {
      set(target, prop: string, value: string) {
        if (prop === 'cssText') {
          for (const decl of value.split(';')) {
            const idx = decl.indexOf(':');
            if (idx > 0) target[decl.slice(0, idx).trim()] = decl.slice(idx + 1).trim();
          }
          target.cssText = value;
        } else {
          target[prop] = value;
        }
        return true;
      },
    }),
    appendChild(c: FakeEl) {
      this.children.push(c);
    },
  };
  return el;
}

afterEach(() => vi.unstubAllGlobals());

function run(partial: Partial<DocxTextRunInfo>): DocxTextRunInfo {
  return { text: 'X', x: 0, y: 0, w: 10, h: 12, fontSize: 12, font: '12px serif', ...partial };
}

// Monospace measurer: each char is 7px wide.
const W = 7;
const measureForFont = () => (s: string) => s.length * W;

describe('buildDocxHighlightLayer', () => {
  it('draws one box per slice, positioned from run x/y + measured extent', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [run({ text: 'the quick', x: 100, y: 50, h: 16 })];
    // Match "quick" = slice [4, 9).
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 4, end: 9 }], active: false },
    ];
    buildDocxHighlightLayer(
      layer as unknown as HTMLDivElement,
      runs,
      matches,
      '700px',
      '900px',
      measureForFont,
    );
    expect(layer.children).toHaveLength(1);
    const box = layer.children[0];
    // left = run.x (100) + prefix "the " width (4*7=28) = 128.
    expect(box.style.left).toBe('128px');
    expect(box.style.top).toBe('50px');
    // width = "quick" (5*7=35).
    expect(box.style.width).toBe('35px');
    expect(box.style.height).toBe('16px');
    expect(box.style.background).toBe(DEFAULT_FIND_HIGHLIGHT);
    expect(box.style['pointer-events']).toBe('none');
  });

  it('uses the active colour for the active match', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [run({ text: 'abc' })];
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 0, end: 3 }], active: true },
    ];
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, '1px', '1px', measureForFont);
    expect(layer.children[0].style.background).toBe(DEFAULT_FIND_ACTIVE_HIGHLIGHT);
  });

  it('draws a box per run for a cross-run match', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [run({ text: 'Hel', x: 0 }), run({ text: 'lo', x: 21 })];
    const matches: DocxHighlightMatch[] = [
      {
        slices: [
          { runIndex: 0, start: 0, end: 3 },
          { runIndex: 1, start: 0, end: 2 },
        ],
        active: false,
      },
    ];
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, '1px', '1px', measureForFont);
    expect(layer.children).toHaveLength(2);
    expect(layer.children[0].style.left).toBe('0px'); // run 0 origin
    expect(layer.children[1].style.left).toBe('21px'); // run 1 origin
  });

  it('carries the run transform for a vertical (tbRl) page', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [run({ text: 'abc', transform: 'rotate(90deg)' })];
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 0, end: 3 }], active: false },
    ];
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, '1px', '1px', measureForFont);
    expect(layer.children[0].style.transform).toBe('rotate(90deg)');
  });

  it('clears the layer and skips zero-width slices', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    layer.innerHTML = 'stale';
    const runs = [run({ text: 'abc' })];
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 1, end: 1 }], active: false }, // degenerate
    ];
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, runs, matches, '1px', '1px', measureForFont);
    expect(layer.innerHTML).toBe('');
    expect(layer.children).toHaveLength(0);
  });
});

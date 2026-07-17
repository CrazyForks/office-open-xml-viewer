import { describe, expect, it } from 'vitest';
import type { DrawingLayout, LayoutPage, PageLayerId } from './types.js';
import {
  createPageLayers,
  orderedPagePaintNodes,
  replacePageLayerNodes,
} from './page-graph.js';

const bounds = { xPt: 0, yPt: 0, widthPt: 10, heightPt: 10 } as const;

function drawing(id: string): DrawingLayout {
  return {
    kind: 'drawing', id,
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'body', ordinaryFlow: false,
    flowBounds: bounds, inkBounds: bounds, advancePt: 0, commands: [],
  };
}

function page(entries: readonly Readonly<{ layer: PageLayerId; node: DrawingLayout }>[]): LayoutPage {
  const byLayer = (layer: PageLayerId) => entries
    .filter((entry) => entry.layer === layer)
    .map((entry) => entry.node);
  return {
    layers: {
      paintSequence: entries.map((entry) => ({
        ...entry, coordinateSpace: 'section-logical' as const,
      })),
      background: byLayer('background'),
      behindText: byLayer('behindText'),
      header: byLayer('header'),
      body: byLayer('body'),
      notes: byLayer('notes'),
      front: byLayer('front'),
      footer: byLayer('footer'),
    },
  } as unknown as LayoutPage;
}

describe('orderedPagePaintNodes body run', () => {
  it('creates a frozen completed sequence with a concrete default coordinate space', () => {
    const body = drawing('body');
    const layers = createPageLayers([{ layer: 'body', node: body }]);

    expect(layers.paintSequence).toEqual([
      { layer: 'body', node: body, coordinateSpace: 'section-logical' },
    ]);
    expect(layers.paintSequence[0]!.node).toBe(layers.body[0]);
    expect(Object.isFrozen(layers)).toBe(true);
    expect(Object.isFrozen(layers.paintSequence)).toBe(true);
    expect(Object.isFrozen(layers.paintSequence[0])).toBe(true);
  });

  it('preserves arbitrary non-body order around one body run', () => {
    const front = drawing('front');
    const body = drawing('body');

    expect(orderedPagePaintNodes(page([
      { layer: 'front', node: front },
      { layer: 'body', node: body },
    ]))).toEqual([front, body]);
  });

  it('rejects re-entry into body after the run has ended', () => {
    const first = drawing('body:first');
    const footer = drawing('footer');
    const second = drawing('body:second');

    expect(() => orderedPagePaintNodes(page([
      { layer: 'body', node: first },
      { layer: 'footer', node: footer },
      { layer: 'body', node: second },
    ]))).toThrow(/contiguous body paint run/i);
  });

  it('rejects a sequence node that only shares the retained node ID', () => {
    const retained = drawing('body');
    const stale = drawing('body');
    const layout = page([{ layer: 'body', node: retained }]);
    const invalid = {
      ...layout,
      layers: {
        ...layout.layers,
        paintSequence: [{
          layer: 'body' as const, node: stale, coordinateSpace: 'section-logical' as const,
        }],
      },
    };

    expect(() => orderedPagePaintNodes(invalid)).toThrow(/not the retained body node/i);
  });

  it('rejects duplicate and missing replacement identities', () => {
    const first = drawing('first');
    const second = drawing('second');
    const layers = createPageLayers([
      { layer: 'body', node: first },
      { layer: 'body', node: second },
    ]);

    expect(() => replacePageLayerNodes(layers, 'body', [first, first])).toThrow(/unique/i);
    expect(() => replacePageLayerNodes(layers, 'body', [first, drawing('other')]))
      .toThrow(/missing replacement/i);
  });
});

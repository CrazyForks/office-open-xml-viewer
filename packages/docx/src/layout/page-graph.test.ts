import { describe, expect, it } from 'vitest';
import type { DrawingLayout, LayoutPage, PageLayerId } from './types.js';
import { orderedPagePaintNodes } from './page-graph.js';

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
      paintOrder: entries.map((entry) => ({ layer: entry.layer, nodeId: entry.node.id })),
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
});

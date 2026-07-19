import { describe, expect, it, vi } from 'vitest';
import {
  attachStoryBlockLayoutAlgorithms,
  layoutStory,
} from './stories.js';
import type {
  BlockLayoutAlgorithms,
  LayoutServices,
  ParagraphLayout,
  ParagraphLayoutInput,
  StoryLayoutInput,
  TableLayout,
  TableLayoutInput,
} from './types.js';

const services = (): LayoutServices => Object.freeze({}) as LayoutServices;

function paragraph(
  input: ParagraphLayoutInput,
  flowDomainId: string,
  yPt: number,
  heightPt: number,
): ParagraphLayout {
  const bounds = { xPt: 12, yPt, widthPt: 80, heightPt };
  return {
    kind: 'paragraph',
    id: `${input.source.story}:${input.source.storyInstance}:${input.source.path.join('.')}`,
    source: input.source,
    flowDomainId,
    ordinaryFlow: true,
    flowBounds: bounds,
    inkBounds: { ...bounds, xPt: 10, widthPt: 84 },
    advancePt: heightPt,
    spacing: { beforePt: 0, afterPt: 0 },
    contextualSpacing: false,
    lines: [],
    borders: [],
    resources: [],
    drawings: [],
    textBoxes: [],
    events: [],
    exclusions: [],
  };
}

function table(
  input: TableLayoutInput,
  flowDomainId: string,
  yPt: number,
  heightPt: number,
): TableLayout {
  const bounds = { xPt: 12, yPt, widthPt: 80, heightPt };
  return {
    kind: 'table',
    id: input.id,
    source: input.source,
    flowDomainId,
    ordinaryFlow: true,
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: heightPt,
    columnWidthsPt: [80],
    rows: [],
    borders: [],
  };
}

function input(blocks: StoryLayoutInput['blocks']): StoryLayoutInput {
  return {
    source: { story: 'header', storyInstance: 'first', path: [] },
    container: {
      id: 'story:header:page:0',
      kind: 'header',
      bounds: { xPt: 12, yPt: 18, widthPt: 80, heightPt: 120 },
    },
    blocks,
  };
}

describe('layoutStory', () => {
  it('delegates mixed blocks to the shared flow dispatcher and retains union bounds', () => {
    const paragraphInput: ParagraphLayoutInput = {
      kind: 'paragraph',
      source: { story: 'header', storyInstance: 'first', path: [0] },
    };
    const tableInput: TableLayoutInput = {
      kind: 'table',
      id: 'header-table',
      source: { story: 'header' as const, storyInstance: 'first', path: [1] },
      flowDomainId: 'acquisition',
      ordinaryFlow: true,
      alignment: 'left' as const,
      indentPt: 0,
      bidiVisual: false,
      columnWidthsPt: [80],
      borders: {
        top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
      },
      rows: [],
    };
    const layoutParagraph = vi.fn<BlockLayoutAlgorithms['layoutParagraph']>(
      (block, placement) => {
        const layout = paragraph(block, placement.container.id, placement.cursor.yPt, 10);
        return {
          layout,
          nextCursor: { xPt: placement.cursor.xPt, yPt: placement.cursor.yPt + 10 },
        };
      },
    );
    const layoutTable = vi.fn<BlockLayoutAlgorithms['layoutTable']>(
      (block, placement) => {
        const layout = table(block, placement.container.id, placement.cursor.yPt, 15);
        return {
          layout,
          nextCursor: { xPt: placement.cursor.xPt, yPt: placement.cursor.yPt + 15 },
        };
      },
    );
    const ownedServices = services();
    attachStoryBlockLayoutAlgorithms(ownedServices, { layoutParagraph, layoutTable });

    const story = layoutStory(input([paragraphInput, tableInput]), ownedServices);

    expect(layoutParagraph).toHaveBeenCalledOnce();
    expect(layoutTable).toHaveBeenCalledOnce();
    expect(story.story).toBe('header');
    expect(story.blocks.map((block) => block.kind)).toEqual(['paragraph', 'table']);
    expect(story.flowBounds).toEqual({ xPt: 12, yPt: 18, widthPt: 80, heightPt: 25 });
    expect(story.inkBounds).toEqual({ xPt: 10, yPt: 18, widthPt: 84, heightPt: 25 });
    expect(story.clipBounds).toEqual({ xPt: 12, yPt: 18, widthPt: 80, heightPt: 120 });
    expect(story.advancePt).toBe(25);
  });

  it('retains an empty story at the container origin without inventing ink', () => {
    const ownedServices = services();
    const unused: BlockLayoutAlgorithms = {
      layoutParagraph() { throw new Error('not used'); },
      layoutTable() { throw new Error('not used'); },
    };
    attachStoryBlockLayoutAlgorithms(ownedServices, unused);

    const story = layoutStory(input([]), ownedServices);

    expect(story.flowBounds).toEqual({ xPt: 12, yPt: 18, widthPt: 0, heightPt: 0 });
    expect(story.inkBounds).toEqual({ xPt: 12, yPt: 18, widthPt: 0, heightPt: 0 });
    expect(story.advancePt).toBe(0);
  });

  it('rejects cross-document algorithm omission and duplicate attachment', () => {
    const first = services();
    const second = services();
    const unused: BlockLayoutAlgorithms = {
      layoutParagraph() { throw new Error('not used'); },
      layoutTable() { throw new Error('not used'); },
    };
    attachStoryBlockLayoutAlgorithms(first, unused);

    expect(() => attachStoryBlockLayoutAlgorithms(first, unused)).toThrow(/already attached/);
    expect(() => layoutStory(input([]), second)).toThrow(/not attached/);
  });

  it('rejects a block owned by another story instance before acquisition', () => {
    const ownedServices = services();
    const unused: BlockLayoutAlgorithms = {
      layoutParagraph() { throw new Error('must not acquire'); },
      layoutTable() { throw new Error('must not acquire'); },
    };
    attachStoryBlockLayoutAlgorithms(ownedServices, unused);

    expect(() => layoutStory(input([{
      kind: 'paragraph',
      source: { story: 'footer', storyInstance: 'default', path: [0] },
    }]), ownedServices)).toThrow(/not owned/);
  });
});

import type {
  LayoutPage,
  PageLayerId,
  PageLayers,
  PaintNode,
} from './types.js';

export const PAGE_LAYER_IDS = [
  'background',
  'behindText',
  'header',
  'body',
  'notes',
  'front',
  'footer',
] as const satisfies readonly PageLayerId[];

type MissingPageLayer = Exclude<PageLayerId, typeof PAGE_LAYER_IDS[number]>;
const pageLayersAreExhaustive: MissingPageLayer extends never ? true : never = true;
void pageLayersAreExhaustive;

export type PageLayerNode = Readonly<{
  layer: PageLayerId;
  node: PaintNode;
  coordinateSpace?: 'section-logical' | 'upright-physical';
}>;

export class PageGraphError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PageGraphError';
  }
}

export function createPageLayers(entries: readonly PageLayerNode[]): PageLayers {
  const nodes = new Map(PAGE_LAYER_IDS.map((layer) => [layer, [] as PaintNode[]]));
  for (const entry of entries) nodes.get(entry.layer)!.push(entry.node);
  return Object.freeze({
    paintSequence: Object.freeze(entries.map(({ layer, node, coordinateSpace }) => Object.freeze({
      layer,
      node,
      coordinateSpace: coordinateSpace ?? 'section-logical',
    }))),
    background: Object.freeze(nodes.get('background')!),
    behindText: Object.freeze(nodes.get('behindText')!),
    header: Object.freeze(nodes.get('header')!),
    body: Object.freeze(nodes.get('body')!),
    notes: Object.freeze(nodes.get('notes')!),
    front: Object.freeze(nodes.get('front')!),
    footer: Object.freeze(nodes.get('footer')!),
  });
}

export function replacePageLayerNodes(
  layers: PageLayers,
  layer: PageLayerId,
  replacements: readonly PaintNode[],
): PageLayers {
  const replacementById = new Map(replacements.map((node) => [node.id, node] as const));
  if (replacementById.size !== replacements.length || replacements.length !== layers[layer].length) {
    throw new PageGraphError(`Replacement ${layer} layer must preserve unique paint node identities`);
  }
  const entries = layers.paintSequence.map((entry): PageLayerNode => {
    if (entry.layer !== layer) return entry;
    const node = replacementById.get(entry.node.id);
    if (!node) throw new PageGraphError(`Missing replacement paint node ${entry.node.id}`);
    return { ...entry, node };
  });
  return createPageLayers(entries);
}

export function pageLayerNodes(page: LayoutPage): readonly PageLayerNode[] {
  return PAGE_LAYER_IDS.flatMap((layer) => (
    page.layers[layer].map((node) => ({ layer, node }))
  ));
}

export function orderedPagePaintNodes(page: LayoutPage): readonly PaintNode[] {
  let seenBody = false;
  let leftBody = false;
  for (const entry of page.layers.paintSequence) {
    if (entry.layer === 'body') {
      if (leftBody) {
        throw new PageGraphError(
          `Paint sequence must contain one contiguous body paint run; re-entered at ${entry.node.id}`,
        );
      }
      seenBody = true;
    } else if (seenBody) {
      leftBody = true;
    }
  }
  const nodes = new Map<string, PageLayerNode>();
  for (const entry of pageLayerNodes(page)) {
    if (nodes.has(entry.node.id)) throw new PageGraphError(`Duplicate paint node ${entry.node.id}`);
    nodes.set(entry.node.id, entry);
  }

  const painted = new Set<string>();
  const ordered: PaintNode[] = [];
  for (const entry of page.layers.paintSequence) {
    const target = nodes.get(entry.node.id);
    if (!target) throw new PageGraphError(`Missing paint node ${entry.node.id}`);
    if (target.layer !== entry.layer) {
      throw new PageGraphError(`Paint node ${entry.node.id} belongs to ${target.layer}, not ${entry.layer}`);
    }
    if (target.node !== entry.node) {
      throw new PageGraphError(`Paint node ${entry.node.id} is not the retained ${entry.layer} node`);
    }
    if (painted.has(entry.node.id)) throw new PageGraphError(`Duplicate paint reference ${entry.node.id}`);
    painted.add(entry.node.id);
    ordered.push(entry.node);
  }
  if (painted.size !== nodes.size) {
    const missing = [...nodes.keys()].find((id) => !painted.has(id));
    throw new PageGraphError(`Missing paint-order reference for ${missing ?? '<unknown>'}`);
  }
  return ordered;
}

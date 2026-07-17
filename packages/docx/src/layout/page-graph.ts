import type {
  LayoutPage,
  PageLayerId,
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

export function pageLayerNodes(page: LayoutPage): readonly PageLayerNode[] {
  return PAGE_LAYER_IDS.flatMap((layer) => (
    page.layers[layer].map((node) => ({ layer, node }))
  ));
}

export function orderedPagePaintNodes(page: LayoutPage): readonly PaintNode[] {
  let seenBody = false;
  let leftBody = false;
  for (const entry of page.layers.paintOrder) {
    if (entry.layer === 'body') {
      if (leftBody) {
        throw new PageGraphError(
          `Paint order must contain one contiguous body paint run; re-entered at ${entry.nodeId}`,
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
  for (const entry of page.layers.paintOrder) {
    const target = nodes.get(entry.nodeId);
    if (!target) throw new PageGraphError(`Missing paint node ${entry.nodeId}`);
    if (target.layer !== entry.layer) {
      throw new PageGraphError(`Paint node ${entry.nodeId} belongs to ${target.layer}, not ${entry.layer}`);
    }
    if (painted.has(entry.nodeId)) throw new PageGraphError(`Duplicate paint reference ${entry.nodeId}`);
    painted.add(entry.nodeId);
    ordered.push(target.node);
  }
  if (painted.size !== nodes.size) {
    const missing = [...nodes.keys()].find((id) => !painted.has(id));
    throw new PageGraphError(`Missing paint-order reference for ${missing ?? '<unknown>'}`);
  }
  return ordered;
}

export function orderedPagePaintEntries(
  page: LayoutPage,
): readonly Readonly<{ node: PaintNode; coordinateSpace: 'section-logical' | 'upright-physical' }>[] {
  orderedPagePaintNodes(page);
  const nodes = new Map(pageLayerNodes(page).map((entry) => [entry.node.id, entry] as const));
  return page.layers.paintOrder.map((entry) => {
    const target = nodes.get(entry.nodeId);
    if (!target || target.layer !== entry.layer) {
      throw new PageGraphError(`Missing paint node ${entry.nodeId} in ${entry.layer}`);
    }
    return Object.freeze({
      node: target.node,
      coordinateSpace: entry.coordinateSpace ?? 'section-logical',
    });
  });
}

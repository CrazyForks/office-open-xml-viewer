import type {
  LayoutPage,
  PageLayerId,
  PageLayers,
  PaintNode,
} from './types.js';
import {
  buildPageLayers,
  PAGE_LAYER_IDS,
  type PageLayerNode,
} from './page-layers.js';

export {
  buildPageLayers,
  PAGE_LAYER_IDS,
  type PageLayerNode,
} from './page-layers.js';

export class PageGraphError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PageGraphError';
  }
}

export const createPageLayers = buildPageLayers;

export function replacePageLayerNodes(
  layers: PageLayers,
  layer: PageLayerId,
  replacements: readonly PaintNode[],
): PageLayers {
  const replacementById = new Map(replacements.map((node) => [node.id, node] as const));
  if (replacementById.size !== replacements.length || replacements.length !== layers[layer].length) {
    throw new PageGraphError(`Replacement ${layer} layer must preserve unique paint node identities`);
  }
  const entries = layers.roots.map((entry): PageLayerNode => {
    if (entry.layer !== layer) return entry;
    const node = replacementById.get(entry.node.id);
    if (!node) throw new PageGraphError(`Missing replacement paint node ${entry.node.id}`);
    return { ...entry, node };
  });
  return createPageLayers(entries);
}

export function pageLayerNodes(page: LayoutPage): readonly PageLayerNode[] {
  return page.layers.roots;
}

export function orderedPagePaintNodes(page: LayoutPage): readonly PaintNode[] {
  let seenBody = false;
  let leftBody = false;
  for (const entry of page.layers.roots) {
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
  for (const entry of page.layers.roots) {
    if (nodes.has(entry.node.id)) throw new PageGraphError(`Duplicate paint node ${entry.node.id}`);
    nodes.set(entry.node.id, entry);
  }
  const semanticNodes = new Map<string, PageLayerNode>();
  for (const layer of PAGE_LAYER_IDS) {
    for (const node of page.layers[layer]) {
      if (semanticNodes.has(node.id)) {
        throw new PageGraphError(`Duplicate semantic page node ${node.id}`);
      }
      semanticNodes.set(node.id, { layer, node });
    }
  }
  if (semanticNodes.size !== nodes.size) {
    throw new PageGraphError('Semantic page layers do not match retained roots');
  }
  for (const [id, root] of nodes) {
    const semantic = semanticNodes.get(id);
    if (!semantic || semantic.layer !== root.layer || semantic.node !== root.node) {
      throw new PageGraphError(`Paint root ${id} is not the retained ${root.layer} node`);
    }
  }

  const representedRoots = new Set<string>();
  const drawingIds = new Set<string>();
  for (const entry of page.layers.paintOrder) {
    const root = nodes.get(entry.rootNodeId);
    if (!root) throw new PageGraphError(`Missing paint root ${entry.rootNodeId}`);
    if (root.layer !== entry.sourceLayer) {
      throw new PageGraphError(
        `Paint root ${entry.rootNodeId} belongs to ${root.layer}, not ${entry.sourceLayer}`,
      );
    }
    representedRoots.add(entry.rootNodeId);
    if (entry.kind === 'node') {
      if (entry.node !== root.node || entry.node.id !== entry.rootNodeId) {
        throw new PageGraphError(
          `Paint root ${entry.rootNodeId} is not the retained ${entry.sourceLayer} node`,
        );
      }
      continue;
    }
    if (!entry.node.anchorLayer) {
      throw new PageGraphError(`Drawing paint entry ${entry.node.id} is not anchored`);
    }
    if (drawingIds.has(entry.node.id)) {
      throw new PageGraphError(`Duplicate drawing paint reference ${entry.node.id}`);
    }
    drawingIds.add(entry.node.id);
  }
  if (representedRoots.size !== nodes.size) {
    const missing = [...nodes.keys()].find((id) => !representedRoots.has(id));
    throw new PageGraphError(`Missing paint-order reference for ${missing ?? '<unknown>'}`);
  }
  return page.layers.roots.map(({ node }) => node);
}

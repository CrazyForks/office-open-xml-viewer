import type {
  DocumentLayout,
  LayoutPage,
  LayoutRect,
  PaintNode,
  PaintResourceKind,
} from '../layout/types.js';
import { orderedPagePaintEntries } from '../layout/page-graph.js';
import {
  enqueueDeferredFrontPaint,
  withDeferredFrontPaintSession,
  type DeferredFrontPaintState,
} from './deferred-front-session.js';
import { paintDrawingLayout } from './canvas-drawing.js';
import { paintParagraphLayout } from './canvas-text.js';
import { paintTableLayout } from './canvas-table.js';
import { appendDeferredPaintFrame, canvasPaintFrame } from './deferred-paint-frame.js';
import type { PaintResourceSession } from './resource-session.js';
import type {
  CanvasPaintContext,
  CanvasPaintResourceHandlers,
  CanvasPaintResourcePainter,
  PaintCanvas2D,
  PaintPageOptions,
} from './types.js';

const missingResourcePainter: CanvasPaintResourcePainter = Object.freeze({
  paint(resourceKey: string, kind: PaintResourceKind): never {
    throw new Error(
      `Missing retained resource painter for ${resourceKey}: expected ${kind}`,
    );
  },
});

export function createCanvasPaintResourcePainter(
  session: PaintResourceSession,
  handlers: CanvasPaintResourceHandlers,
): CanvasPaintResourcePainter {
  return Object.freeze({
    paint(
      resourceKey: string,
      kind: PaintResourceKind,
      bounds: LayoutRect,
      ctx: PaintCanvas2D,
    ): void {
      switch (kind) {
        case 'image':
          handlers.image(session.resolve(resourceKey, kind), bounds, ctx);
          return;
        case 'chart':
          handlers.chart(session.resolve(resourceKey, kind), bounds, ctx);
          return;
        case 'math':
          handlers.math(session.resolve(resourceKey, kind), bounds, ctx);
          return;
        case 'picture-bullet':
          handlers['picture-bullet'](session.resolve(resourceKey, kind), bounds, ctx);
          return;
        default: {
          const exhaustive: never = kind;
          throw new Error(`Unknown retained resource kind: ${String(exhaustive)}`);
        }
      }
    },
  });
}

function paintNode(node: PaintNode, context: CanvasPaintContext): void {
  switch (node.kind) {
    case 'drawing':
      paintDrawingLayout(node, context);
      return;
    case 'paragraph':
      paintParagraphLayout(node, context);
      return;
    case 'table':
      paintTableLayout(node, context, node.resolvedFloatingTables ?? []);
      return;
    case 'textbox':
    case 'note':
      throw new Error(`Unsupported page paint node kind: ${node.kind}`);
    default: {
      const exhaustive: never = node;
      throw new Error(`Unknown page paint node kind: ${String(exhaustive)}`);
    }
  }
}

function paintBodyNode(node: PaintNode, context: CanvasPaintContext): void {
  if (node.kind !== 'drawing') {
    paintNode(node, context);
    return;
  }
  const paint = () => paintDrawingLayout(node, context);
  const deferredPaint = context.deferredPaintWrapper?.(paint) ?? paint;
  if (context.bodyDrawingPass === 'discover-behind') {
    if (node.anchorLayer?.behindDoc) {
      context.deferBehindDrawing?.(node, [], deferredPaint);
    }
    return;
  }
  if (node.anchorLayer?.behindDoc) {
    if (!context.deferBehindDrawing?.(node, [], deferredPaint)) paint();
    return;
  }
  if (node.anchorLayer) {
    if (!context.deferFrontDrawing?.(node, [], deferredPaint)) paint();
    return;
  }
  paint();
}

function applyRegionTransform(
  ctx: PaintCanvas2D,
  matrix: LayoutPage['sectionRegions'][number]['coordinateSpace']['logicalToPhysical'],
): void {
  const transform = (ctx as PaintCanvas2D & {
    transform?: (a: number, b: number, c: number, d: number, e: number, f: number) => void;
  }).transform;
  if (transform) {
    transform.call(ctx, matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f);
    return;
  }
  ctx.translate(matrix.e, matrix.f);
  if (matrix.a === 0 && matrix.b === 1 && matrix.c === -1 && matrix.d === 0) {
    ctx.rotate(Math.PI / 2);
  } else if (matrix.a === 0 && matrix.b === -1 && matrix.c === 1 && matrix.d === 0) {
    ctx.rotate(-Math.PI / 2);
  } else if (matrix.b === 0 && matrix.c === 0) {
    ctx.scale(matrix.a, matrix.d);
  } else {
    throw new Error('Canvas context cannot apply the retained section transform');
  }
}

type PagePaintEntry = Readonly<{
  node: PaintNode;
  layer: LayoutPage['layers']['paintOrder'][number]['layer'];
  coordinateSpace: 'section-logical' | 'upright-physical';
}>;

function paintInEntryRegion(
  entry: PagePaintEntry,
  context: CanvasPaintContext,
  regionByDomain: ReadonlyMap<string, LayoutPage['sectionRegions'][number]>,
  paint: (entryContext: CanvasPaintContext) => void,
  enterFrame = true,
): void {
  const region = regionByDomain.get(entry.node.flowDomainId);
  const matrix = entry.coordinateSpace === 'upright-physical'
    ? undefined
    : region?.coordinateSpace.logicalToPhysical;
  const frame = canvasPaintFrame(context.ctx, () => {
    if (matrix && (
      matrix.a !== 1 || matrix.b !== 0 || matrix.c !== 0
      || matrix.d !== 1 || matrix.e !== 0 || matrix.f !== 0
    )) applyRegionTransform(context.ctx, matrix);
  });
  const entryContext: CanvasPaintContext = {
    ...context,
    ...(matrix ? {
      pointToCss: {
        a: matrix.a * context.scale,
        b: matrix.b * context.scale,
        c: matrix.c * context.scale,
        d: matrix.d * context.scale,
        e: matrix.e * context.scale,
        f: matrix.f * context.scale,
      },
    } : {}),
    deferredPaintWrapper: appendDeferredPaintFrame(context.deferredPaintWrapper, frame),
  };
  if (enterFrame) frame(() => paint(entryContext))();
  else paint(entryContext);
}

/** Paint a completed page into an already initialized point-space surface. */
export function paintLayoutPageContent(
  page: LayoutPage,
  context: CanvasPaintContext,
): void {
  const regionByDomain = new Map(page.sectionRegions.flatMap((region) => (
    region.flowDomainIds.map((domainId) => [domainId, region] as const)
  )));
  const ordered = orderedPagePaintEntries(page);
  const entries: readonly PagePaintEntry[] = ordered.map((entry, index) => ({
    ...entry,
    layer: page.layers.paintOrder[index]!.layer,
  }));
  const firstBodyEntry = entries.findIndex((entry) => entry.layer === 'body');
  if (firstBodyEntry === -1) {
    for (const entry of entries) {
      paintInEntryRegion(entry, context, regionByDomain, (entryContext) => {
        paintNode(entry.node, entryContext);
      });
    }
    return;
  }
  let lastBodyEntry = firstBodyEntry;
  while (entries[lastBodyEntry + 1]?.layer === 'body') lastBodyEntry += 1;
  for (const entry of entries.slice(0, firstBodyEntry)) {
    paintInEntryRegion(entry, context, regionByDomain, (entryContext) => {
      paintNode(entry.node, entryContext);
    });
  }
  const behind: Array<Readonly<{
    drawing: import('../layout/types.js').DrawingLayout;
    paint: () => void;
    encounterOrder: number;
  }>> = [];
  for (const entry of entries.slice(firstBodyEntry, lastBodyEntry + 1)) {
    // Discovery walks immutable retained geometry only. Keep the section frame
    // in deferred replay wrappers, but do not apply it to this ink-free walk;
    // normal paint is the destination-final logical-to-physical owner.
    paintInEntryRegion(entry, context, regionByDomain, (entryContext) => {
      paintBodyNode(entry.node, {
        ...entryContext,
        bodyDrawingPass: 'discover-behind',
        deferBehindDrawing: (drawing, _textBoxes, paint) => {
          behind.push({ drawing, paint, encounterOrder: behind.length });
          return true;
        },
      });
    }, false);
  }
  behind.sort((a, b) =>
    a.drawing.anchorLayer!.relativeHeight - b.drawing.anchorLayer!.relativeHeight
    || a.drawing.anchorLayer!.sourceOrder - b.drawing.anchorLayer!.sourceOrder
    || a.encounterOrder - b.encounterOrder);
  const paintedBehind = new Set(behind.map(({ drawing }) => drawing));
  // wp:anchor z-order is page-relative (ECMA-376 Part 1 §20.4.2.3), so retained
  // behindDoc owners must paint before the first body entry, not at paragraph entry.
  for (const deferred of behind) {
    deferred.paint();
  }

  const frontState: DeferredFrontPaintState = {};
  withDeferredFrontPaintSession(frontState, () => {
    for (const entry of entries.slice(firstBodyEntry, lastBodyEntry + 1)) {
      paintInEntryRegion(entry, context, regionByDomain, (entryContext) => {
        if (entry.layer !== 'body') {
          paintNode(entry.node, entryContext);
          return;
        }
        paintBodyNode(entry.node, {
          ...entryContext,
          bodyDrawingPass: 'normal',
          deferBehindDrawing: (drawing) => paintedBehind.has(drawing),
          deferFrontDrawing: (drawing, _textBoxes, paint) => enqueueDeferredFrontPaint(
            frontState,
            paint,
            drawing.anchorLayer ? {
              relativeHeight: drawing.anchorLayer.relativeHeight,
              sourceOrder: drawing.anchorLayer.sourceOrder,
            } : undefined,
          ),
        });
      });
    }
  });
  for (const entry of entries.slice(lastBodyEntry + 1)) {
    paintInEntryRegion(entry, context, regionByDomain, (entryContext) => {
      paintNode(entry.node, entryContext);
    });
  }
}

export async function paintLayoutPage(
  layout: DocumentLayout,
  pageIndex: number,
  target: HTMLCanvasElement | OffscreenCanvas,
  options: PaintPageOptions,
  resources: CanvasPaintResourcePainter = missingResourcePainter,
): Promise<void> {
  const page = layout.pages[pageIndex];
  if (!page) throw new RangeError(`Page ${pageIndex} is outside the layout`);
  const ctx = target.getContext('2d') as
    | CanvasRenderingContext2D
    | OffscreenCanvasRenderingContext2D
    | null;
  if (!ctx) throw new Error('Canvas 2D context is unavailable');

  const pixelScale = options.scale * options.dpr;
  target.width = Math.ceil(page.geometry.widthPt * pixelScale);
  target.height = Math.ceil(page.geometry.heightPt * pixelScale);
  ctx.save();
  try {
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, target.width, target.height);
    ctx.setTransform(pixelScale, 0, 0, pixelScale, 0, 0);
    paintLayoutPageContent(page, {
      ctx, scale: options.scale, dpr: options.dpr, resources,
    });
  } finally {
    ctx.restore();
  }
}

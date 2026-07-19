import type {
  ParagraphLayout,
  ResolvedFloatingTablePlacementLayout,
  TableLayout,
} from '../layout/types.js';
import { composeAffine, scaleAffine, translationAffine } from './affine.js';
import { paintStrokeSegment } from './canvas-border.js';
import { paintParagraphLayout } from './canvas-text.js';
import { canvasPaintFrame } from './deferred-paint-frame.js';
import type { CanvasPaintContext } from './types.js';

function paintPlacedChild(
  layout: ParagraphLayout | TableLayout,
  placement: Readonly<{ xPt: number; yPt: number }>,
  context: CanvasPaintContext,
): void {
  const dxPt = placement.xPt - layout.flowBounds.xPt;
  const dyPt = placement.yPt - layout.flowBounds.yPt;
  const parentLayoutTranslation = context.layoutTranslationPt ?? {
    xPt: 0,
    yPt: 0,
  };
  const pointToCss = composeAffine(
    context.pointToCss ?? scaleAffine(context.scale),
    translationAffine(dxPt, dyPt),
  );
  const frame = canvasPaintFrame(context.ctx, () => context.ctx.translate(dxPt, dyPt));
  const childContext = {
    ...context,
    pointToCss,
    layoutTranslationPt: {
      xPt: parentLayoutTranslation.xPt + dxPt,
      yPt: parentLayoutTranslation.yPt + dyPt,
    },
  };
  frame(() => {
    if (layout.kind === 'paragraph') paintParagraphLayout(layout, childContext);
    else paintTableLayout(layout, childContext);
  })();
}

function paintTableContents(
  node: TableLayout,
  context: CanvasPaintContext,
  floatingTables: readonly ResolvedFloatingTablePlacementLayout[],
): void {
  for (const row of node.rows) {
    for (const cell of row.cells) {
      const ownsContinuationPaint = 'visualMergeOwnership' in cell
        && cell.visualMergeOwnership === 'continuation';
      if (cell.verticalMerge === 'continue' && !ownsContinuationPaint) continue;
      if (cell.background) {
        context.ctx.fillStyle = cell.background.color;
        context.ctx.fillRect(
          cell.flowBounds.xPt,
          cell.flowBounds.yPt,
          cell.flowBounds.widthPt,
          cell.flowBounds.heightPt,
        );
      }
      const paintBlocks = (blockContext: CanvasPaintContext): void => {
        for (const block of cell.blocks) {
          paintPlacedChild(block.layout, {
            // Paragraphs are normalized to the cell content origin. A nested
            // table's own flowBounds additionally retain jc/tblInd placement
            // within that band, so translate its coordinate space by the outer
            // content origin without erasing that local offset.
            xPt: cell.contentBounds.xPt
              + (block.layout.kind === 'table' ? block.layout.flowBounds.xPt : 0),
            yPt: cell.flowBounds.yPt + block.offsetPt
              + (block.layout.kind === 'table' ? block.layout.flowBounds.yPt : 0),
          }, blockContext);
        }
      };
      if (!cell.clipBounds) {
        paintBlocks(context);
        continue;
      }
      const clipBounds = cell.clipBounds;
      const frame = canvasPaintFrame(context.ctx, () => {
        context.ctx.beginPath();
        context.ctx.rect(
          clipBounds.xPt,
          clipBounds.yPt,
          clipBounds.widthPt,
          clipBounds.heightPt,
        );
        context.ctx.clip();
      });
      frame(() => paintBlocks(context))();
    }
  }
  paintResolvedFloatingTablePlacements(floatingTables, context);
  for (const border of node.borders) paintStrokeSegment(border, context);
}

/** Paint stored table geometry. No inheritance, sizing, shaping, or conflict resolution occurs here. */
export function paintTableLayout(
  node: TableLayout,
  context: CanvasPaintContext,
  floatingTables?: readonly ResolvedFloatingTablePlacementLayout[],
): void {
  const placements = floatingTables ?? node.resolvedFloatingTables ?? [];
  if (!node.clipBounds) {
    paintTableContents(node, context, placements);
    return;
  }
  const clipBounds = node.clipBounds;
  const frame = canvasPaintFrame(context.ctx, () => {
    context.ctx.beginPath();
    context.ctx.rect(
      clipBounds.xPt,
      clipBounds.yPt,
      clipBounds.widthPt,
      clipBounds.heightPt,
    );
    context.ctx.clip();
  });
  frame(() => paintTableContents(node, context, placements))();
}

/** Paint only point-space placements already resolved by the layout adapter. */
export function paintResolvedFloatingTablePlacements(
  placements: readonly ResolvedFloatingTablePlacementLayout[],
  context: CanvasPaintContext,
): void {
  const parentTranslation = context.layoutTranslationPt ?? { xPt: 0, yPt: 0 };
  for (const placement of placements) {
    paintPlacedChild(placement.child, {
      xPt: placement.xPt - parentTranslation.xPt,
      yPt: placement.yPt - parentTranslation.yPt,
    }, context);
  }
}

/** Bridge an unscaled renderer canvas to the retained point-space table painter. */
export function paintPlacedTableLayout(
  node: TableLayout,
  placement: Readonly<{ xPt: number; yPt: number }>,
  context: CanvasPaintContext,
  floatingTables?: readonly ResolvedFloatingTablePlacementLayout[],
): void {
  const dxPt = placement.xPt - node.flowBounds.xPt;
  const dyPt = placement.yPt - node.flowBounds.yPt;
  const pointToCss = composeAffine(
    context.pointToCss ?? scaleAffine(context.scale),
    translationAffine(dxPt, dyPt),
  );
  const frame = canvasPaintFrame(context.ctx, () => {
    context.ctx.translate(dxPt * context.scale, dyPt * context.scale);
    context.ctx.scale(context.scale, context.scale);
  });
  const placedContext: CanvasPaintContext = {
    ...context,
    pointToCss,
    layoutTranslationPt: { xPt: dxPt, yPt: dyPt },
  };
  frame(() => paintTableLayout(node, placedContext, floatingTables))();
}

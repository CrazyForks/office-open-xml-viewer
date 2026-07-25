import type { LayoutRect, PaintResourceKind } from '../layout/types.js';
import type { CanvasPaintContext } from './types.js';

/** Paint one retained non-text resource using the orientation selected during
 * layout acquisition. A vertical section rotates its section-logical frame
 * +90 degrees; physical graphics counter-rotate locally so their authored
 * DrawingML transform is subsequently composed in an upright frame. */
export function paintRetainedResource(
  resourceKey: string,
  resourceKind: PaintResourceKind,
  bounds: LayoutRect,
  orientation: 'upright-physical' | undefined,
  context: CanvasPaintContext,
): void {
  if (orientation !== 'upright-physical') {
    context.resources.paint(resourceKey, resourceKind, bounds, context.ctx);
    return;
  }
  const { ctx } = context;
  ctx.save();
  ctx.translate(
    bounds.xPt + bounds.widthPt / 2,
    bounds.yPt + bounds.heightPt / 2,
  );
  ctx.rotate(-Math.PI / 2);
  context.resources.paint(resourceKey, resourceKind, {
    xPt: -bounds.heightPt / 2,
    yPt: -bounds.widthPt / 2,
    widthPt: bounds.heightPt,
    heightPt: bounds.widthPt,
  }, ctx);
  ctx.restore();
}

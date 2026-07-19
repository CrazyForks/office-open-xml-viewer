import type {
  DeferredCanvasPaintWrapper,
  PaintCanvas2D,
} from './types.js';

/** Create a replayable Canvas save/apply/restore frame. */
export function canvasPaintFrame(
  ctx: PaintCanvas2D,
  apply: () => void,
): DeferredCanvasPaintWrapper {
  return (paint) => () => {
    ctx.save();
    try {
      apply();
      paint();
    } finally {
      ctx.restore();
    }
  };
}

/** Append one child frame to the exact recursive path already retained by its parent. */
export function appendDeferredPaintFrame(
  parent: DeferredCanvasPaintWrapper | undefined,
  child: DeferredCanvasPaintWrapper,
): DeferredCanvasPaintWrapper {
  return (paint) => parent ? parent(child(paint)) : child(paint);
}

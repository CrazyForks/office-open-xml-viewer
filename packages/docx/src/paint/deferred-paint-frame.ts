import type { PaintCanvas2D } from './types.js';

/** Enter one synchronous Canvas save/apply/restore frame. */
export function canvasPaintFrame(
  ctx: PaintCanvas2D,
  apply: () => void,
): (paint: () => void) => () => void {
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

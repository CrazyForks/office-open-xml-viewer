import type { Matrix2DData } from '../layout/types.js';
import type { PaintCanvas2D } from './types.js';

/** Apply a retained point-space affine to the current Canvas frame. */
export function applyCanvasTransform(
  ctx: PaintCanvas2D,
  matrix: Matrix2DData,
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
    throw new Error('Canvas context cannot apply the retained point-space transform');
  }
}

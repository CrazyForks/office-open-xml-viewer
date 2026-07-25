import type { Matrix2DData } from '../layout/types.js';

export {
  composeAffine,
  inverseMapAffinePoint,
  inverseMapAffineVector,
  mapAffinePoint,
  quarterTurnAffine,
  scaleAffine,
  translationAffine,
} from '../layout/affine.js';

/** CSS transform for the orientation component of a point-to-CSS affine map. */
export function cssTransformFor(matrix: Matrix2DData): string | undefined {
  const inlineScale = Math.hypot(matrix.a, matrix.b);
  const blockScale = Math.hypot(matrix.c, matrix.d);
  const a = matrix.a / inlineScale;
  const b = matrix.b / inlineScale;
  const c = matrix.c / blockScale;
  const d = matrix.d / blockScale;
  if (a === 1 && b === 0 && c === 0 && d === 1) return undefined;
  if (a === 0 && b === 1 && c === -1 && d === 0) return 'rotate(90deg)';
  if (a === 0 && b === -1 && c === 1 && d === 0) return 'rotate(-90deg)';
  return `matrix(${a}, ${b}, ${c}, ${d}, 0, 0)`;
}

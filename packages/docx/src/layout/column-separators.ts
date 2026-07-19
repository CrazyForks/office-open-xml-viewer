import { transformPoint } from './coordinate-space.js';
import type { ColumnSeparatorLayout, PageSectionRegion, PointPt } from './types.js';

function freezePoint(point: PointPt): PointPt {
  return Object.freeze(point);
}

/** Geometry follows the retained section band because ink bounds cannot prove
 * where Word terminates a section-scoped column rule. */
export function columnSeparatorSegments(
  regions: readonly PageSectionRegion[],
): readonly ColumnSeparatorLayout[] {
  const segments: ColumnSeparatorLayout[] = [];
  for (const region of regions) {
    const { columns, columnSeparator } = region.section;
    if (!columnSeparator
      || columns.length < 2
      || region.blockEndPt <= region.blockStartPt) continue;
    const ownedColumns = new Set(region.columnIndexes);
    const populationOrder = region.columnFlowDirection === 'rtl'
      ? columns.map((_, index) => index).reverse()
      : columns.map((_, index) => index);
    for (let ordinal = 0; ordinal < populationOrder.length - 1; ordinal += 1) {
      const populatedColumnIndex = populationOrder[ordinal]!;
      if (!ownedColumns.has(populatedColumnIndex)) continue;
      const followingColumnIndex = populationOrder[ordinal + 1]!;
      const leftIndex = Math.min(populatedColumnIndex, followingColumnIndex);
      const rightIndex = Math.max(populatedColumnIndex, followingColumnIndex);
      const left = columns[leftIndex]!;
      const right = columns[rightIndex]!;
      const inlinePt = (left.xPt + left.wPt + right.xPt) / 2;
      segments.push(Object.freeze({
        start: freezePoint(transformPoint(region.coordinateSpace.logicalToPhysical, {
          xPt: inlinePt,
          yPt: region.blockStartPt,
        })),
        end: freezePoint(transformPoint(region.coordinateSpace.logicalToPhysical, {
          xPt: inlinePt,
          yPt: region.blockEndPt,
        })),
      }));
    }
  }
  return Object.freeze(segments);
}

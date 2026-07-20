import { transformPoint } from './coordinate-space.js';
import type { ColumnSeparatorLayout, PageSectionRegion, PointPt } from './types.js';
import { wordColumnSeparatorBlockBand } from './section-compatibility.js';

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
    const band = wordColumnSeparatorBlockBand(
      region.blockStartPt,
      region.blockEndPt,
    );
    if (!columnSeparator
      || columns.length < 2
      || band.blockEndPt <= band.blockStartPt) continue;
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
          yPt: band.blockStartPt,
        })),
        end: freezePoint(transformPoint(region.coordinateSpace.logicalToPhysical, {
          xPt: inlinePt,
          yPt: band.blockEndPt,
        })),
      }));
    }
  }
  return Object.freeze(segments);
}

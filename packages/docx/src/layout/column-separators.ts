import { transformPoint } from './coordinate-space.js';
import type { PageSectionRegion, PointPt } from './types.js';

export interface ColumnSeparatorSegment {
  readonly start: PointPt;
  readonly end: PointPt;
}

/** Geometry follows the retained section band because ink bounds cannot prove
 * where Word terminates a section-scoped column rule. */
export function columnSeparatorSegments(
  regions: readonly PageSectionRegion[],
): readonly ColumnSeparatorSegment[] {
  const segments: ColumnSeparatorSegment[] = [];
  for (const region of regions) {
    // The current region model owns every normalized section column; until
    // column-subset regions exist, its section remains the decoration authority.
    const { columns, columnSeparator } = region.section;
    if (!columnSeparator
      || columns.length < 2
      || region.blockEndPt <= region.blockStartPt) continue;
    for (let index = 0; index < columns.length - 1; index += 1) {
      const left = columns[index]!;
      const right = columns[index + 1]!;
      const inlinePt = (left.xPt + left.wPt + right.xPt) / 2;
      segments.push({
        start: transformPoint(region.coordinateSpace.logicalToPhysical, {
          xPt: inlinePt,
          yPt: region.blockStartPt,
        }),
        end: transformPoint(region.coordinateSpace.logicalToPhysical, {
          xPt: inlinePt,
          yPt: region.blockEndPt,
        }),
      });
    }
  }
  return segments;
}

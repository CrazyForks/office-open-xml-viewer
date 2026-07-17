import type { ColumnSeparatorSegment } from '../layout/column-separators.js';

export interface RasterizedColumnSeparator {
  readonly segment: ColumnSeparatorSegment;
  readonly widthPt: number;
}

function snapTangent(valuePt: number, deviceScale: number): number {
  return Math.round(valuePt * deviceScale) / deviceScale;
}

function snapNormal(valuePt: number, deviceScale: number, widthDevicePx: number): number {
  const valueDevicePx = valuePt * deviceScale;
  const snappedDevicePx = widthDevicePx % 2 === 0
    ? Math.round(valueDevicePx)
    : Math.round(valueDevicePx - 0.5) + 0.5;
  return snappedDevicePx / deviceScale;
}

/** Raster alignment belongs to paint: retained geometry stays in exact points,
 * while the stroke is aligned to its actual device-pixel width here. */
export function rasterizeColumnSeparator(
  segment: ColumnSeparatorSegment,
  scale: number,
  dpr: number,
): RasterizedColumnSeparator {
  const deviceScale = scale * dpr;
  const widthCssPx = Math.max(1, Math.round(0.5 * scale));
  const widthDevicePx = Math.max(1, Math.round(widthCssPx * dpr));
  const widthPt = widthDevicePx / deviceScale;
  if (segment.start.xPt === segment.end.xPt) {
    const xPt = snapNormal(segment.start.xPt, deviceScale, widthDevicePx);
    return {
      segment: {
        start: { xPt, yPt: snapTangent(segment.start.yPt, deviceScale) },
        end: { xPt, yPt: snapTangent(segment.end.yPt, deviceScale) },
      },
      widthPt,
    };
  }
  if (segment.start.yPt === segment.end.yPt) {
    const yPt = snapNormal(segment.start.yPt, deviceScale, widthDevicePx);
    return {
      segment: {
        start: { xPt: snapTangent(segment.start.xPt, deviceScale), yPt },
        end: { xPt: snapTangent(segment.end.xPt, deviceScale), yPt },
      },
      widthPt,
    };
  }
  return { segment, widthPt };
}

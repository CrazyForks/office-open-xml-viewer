type ReflectionCanvas = HTMLCanvasElement | OffscreenCanvas;
type ReflectionContext = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

export interface ReflectionBlurBounds {
  x: number;
  y: number;
  w: number;
  h: number;
}

interface ReflectionBlurBand {
  y: number;
  h: number;
  radius: number;
}

/**
 * PowerPoint keeps a floor reflection sharp where it meets the source and
 * progressively increases the authored blur toward the far edge. ECMA-376
 * §20.1.8.50 defines the authored blur radius but does not prescribe a Canvas
 * rasterisation algorithm. Treat that value as the far-edge radius and
 * approximate the observed Office result with narrow monotonic bands. The
 * source is painted once; only bitmap copies are repeated, keeping glyph
 * shaping out of the band loop.
 */
function reflectionBlurBands(
  bounds: ReflectionBlurBounds,
  maxBlur: number,
): ReflectionBlurBand[] {
  if (!(maxBlur > 0) || !(bounds.h > 0)) {
    return [{ y: bounds.y, h: Math.max(0, bounds.h), radius: 0 }];
  }

  // Around eight samples per blur pixel keeps adjacent filter radii close
  // enough to avoid visible steps. Bound the number of bitmap passes because a
  // presentation may contain many reflected runs.
  const count = Math.max(4, Math.min(24, Math.ceil(maxBlur * 8)));
  const bottom = bounds.y + bounds.h;
  const bands: ReflectionBlurBand[] = [];
  for (let index = 0; index < count; index++) {
    const near = index / count;
    const far = (index + 1) / count;
    const y = bottom - far * bounds.h;
    bands.push({
      y,
      h: bottom - near * bounds.h - y,
      // The first band touches the source and must remain sharp. The last band
      // reaches the authored blur radius at the far edge.
      radius: maxBlur * index / (count - 1),
    });
  }
  return bands;
}

/** Paint a sharp reflection source with blur that grows away from its bottom edge. */
export function paintDistanceAwareReflectionBlur(
  target: ReflectionContext,
  source: ReflectionCanvas,
  bounds: ReflectionBlurBounds,
  maxBlur: number,
  canvasWidth: number,
): void {
  for (const band of reflectionBlurBands(bounds, maxBlur)) {
    target.save();
    target.beginPath();
    target.rect(0, band.y, canvasWidth, band.h);
    target.clip();
    target.filter = band.radius > 0 ? `blur(${band.radius}px)` : 'none';
    target.drawImage(source as CanvasImageSource, 0, 0);
    target.restore();
  }
}

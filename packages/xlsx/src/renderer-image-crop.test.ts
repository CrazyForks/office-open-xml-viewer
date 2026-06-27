import { describe, it, expect, vi } from 'vitest';
import { drawImageCropped } from './renderer';

/**
 * `drawImageCropped` honors an ECMA-376 §20.1.8.55 `<a:srcRect>` crop by drawing
 * only the visible source sub-rectangle into the (unchanged) destination box via
 * the 9-arg `ctx.drawImage` — the bug fix for sample-27, whose picture was a
 * horizontally-cropped PNG that previously rendered whole and squished. The same
 * helper serves the top-level `twoCellAnchor` picture and the `grpSp` /
 * `oneCellAnchor` leaf pic, so every placement crops uniformly.
 *
 * The crop is raster-only: a metafile (WMF/EMF) is rasterized to the CROPPED
 * display box by the decoder, so its bitmap pixels no longer map to source
 * fractions and the crop is skipped (full draw).
 */

/** A decoded bitmap stand-in exposing native pixel `width`/`height`. */
const fakeImg = (w: number, h: number): CanvasImageSource =>
  ({ width: w, height: h }) as unknown as CanvasImageSource;

function spyCtx(): { ctx: CanvasRenderingContext2D; drawImage: ReturnType<typeof vi.fn> } {
  const drawImage = vi.fn();
  return { ctx: { drawImage } as unknown as CanvasRenderingContext2D, drawImage };
}

describe('drawImageCropped srcRect crop', () => {
  it('draws only the visible sub-rectangle for a horizontal raster crop (9-arg drawImage)', () => {
    const { ctx, drawImage } = spyCtx();
    const img = fakeImg(2860, 1368); // sample-27 PNG native pixel size
    // sample-27: left 0.3256, right 0.03829, no vertical crop.
    drawImageCropped(ctx, img, { l: 0.3256, t: 0, r: 0.03829, b: 0 }, 'image/png', 10, 20, 305, 229);

    expect(drawImage).toHaveBeenCalledTimes(1);
    const call = drawImage.mock.calls[0];
    expect(call).toHaveLength(9); // img + (sx,sy,sw,sh) + (dx,dy,dw,dh)
    const [, sx, sy, sw, sh, dx, dy, dw, dh] = call;
    expect(sx).toBeCloseTo(0.3256 * 2860, 3); // skip the left 32.56%
    expect(sy).toBe(0);
    expect(sw).toBeCloseTo((1 - 0.3256 - 0.03829) * 2860, 3); // keep the middle band
    expect(sh).toBe(1368); // full height (no vertical crop)
    // Destination box is unchanged — the slice stretches to fill the rect.
    expect([dx, dy, dw, dh]).toEqual([10, 20, 305, 229]);
  });

  it('crops the vertical axis correctly (sy/sh) for a top+bottom crop', () => {
    const { ctx, drawImage } = spyCtx();
    // top 0.10, bottom 0.25 → sy=0.10·H, sh=(1−0.10−0.25)·H; no horizontal crop.
    drawImageCropped(ctx, fakeImg(400, 200), { l: 0, t: 0.1, r: 0, b: 0.25 }, 'image/png', 0, 0, 80, 40);

    const [, sx, sy, sw, sh] = drawImage.mock.calls[0];
    expect(sx).toBe(0);
    expect(sw).toBe(400); // full width
    expect(sy).toBeCloseTo(0.1 * 200, 6); // 20
    expect(sh).toBeCloseTo((1 - 0.1 - 0.25) * 200, 6); // 130
  });

  it('clamps a crop that extends past the image and never produces a zero-size source', () => {
    const { ctx, drawImage } = spyCtx();
    // Pathological insets (sum > 1, negative overscan) must clamp to a ≥1px rect.
    drawImageCropped(ctx, fakeImg(100, 100), { l: 0.9, t: -0.2, r: 0.9, b: 1.5 }, 'image/png', 0, 0, 40, 40);

    const [, , , sw, sh] = drawImage.mock.calls[0];
    expect(sw).toBeGreaterThanOrEqual(1);
    expect(sh).toBeGreaterThanOrEqual(1);
  });

  it('draws the whole image (4-arg) when there is no crop', () => {
    const { ctx, drawImage } = spyCtx();
    drawImageCropped(ctx, fakeImg(100, 100), undefined, 'image/png', 0, 0, 50, 50);
    expect(drawImage.mock.calls[0]).toHaveLength(5); // img + (dx,dy,dw,dh)
    expect(drawImage).toHaveBeenCalledWith(expect.anything(), 0, 0, 50, 50);
  });

  it('treats an all-zero srcRect as no crop (full draw)', () => {
    const { ctx, drawImage } = spyCtx();
    drawImageCropped(ctx, fakeImg(100, 100), { l: 0, t: 0, r: 0, b: 0 }, 'image/png', 0, 0, 50, 50);
    expect(drawImage.mock.calls[0]).toHaveLength(5);
  });

  it('skips the crop for a metafile (WMF) — it is rasterized to the display box', () => {
    const { ctx, drawImage } = spyCtx();
    drawImageCropped(ctx, fakeImg(200, 200), { l: 0.3, t: 0, r: 0.1, b: 0 }, 'image/wmf', 0, 0, 50, 50);
    expect(drawImage.mock.calls[0]).toHaveLength(5); // full draw, not 9-arg
  });

  it('skips the crop for an EMF metafile too', () => {
    const { ctx, drawImage } = spyCtx();
    drawImageCropped(ctx, fakeImg(200, 200), { l: 0.2, t: 0.2, r: 0, b: 0 }, 'image/emf', 0, 0, 50, 50);
    expect(drawImage.mock.calls[0]).toHaveLength(5);
  });
});

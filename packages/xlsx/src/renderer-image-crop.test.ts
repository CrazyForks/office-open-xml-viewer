import { describe, it, expect, vi } from 'vitest';
import { drawAnchorImage } from './renderer';
import type { ImageAnchor } from './types';

/**
 * `drawAnchorImage` honors an ECMA-376 §20.1.8.55 `<a:srcRect>` crop by drawing
 * only the visible source sub-rectangle into the (unchanged) anchor box via the
 * 9-arg `ctx.drawImage` — the bug fix for sample-27, whose picture was a
 * horizontally-cropped PNG that previously rendered whole and squished.
 *
 * The crop is raster-only: a metafile (WMF/EMF) is rasterized to the CROPPED
 * display box by the decoder, so its bitmap pixels no longer map to source
 * fractions and the crop is skipped (full draw), matching the docx renderer.
 */

/** Minimal `ImageAnchor` with a 1×1-cell anchor; callers override the crop. */
function anchor(over: Partial<ImageAnchor>): ImageAnchor {
  return {
    fromCol: 0,
    fromColOff: 0,
    fromRow: 0,
    fromRowOff: 0,
    toCol: 1,
    toColOff: 0,
    toRow: 1,
    toRowOff: 0,
    nativeExtCx: 0,
    nativeExtCy: 0,
    imagePath: 'xl/media/image1.png',
    mimeType: 'image/png',
    ...over,
  };
}

/** A decoded bitmap stand-in exposing native pixel `width`/`height`. */
const fakeImg = (w: number, h: number): CanvasImageSource =>
  ({ width: w, height: h }) as unknown as CanvasImageSource;

function spyCtx(): { ctx: CanvasRenderingContext2D; drawImage: ReturnType<typeof vi.fn> } {
  const drawImage = vi.fn();
  return { ctx: { drawImage } as unknown as CanvasRenderingContext2D, drawImage };
}

describe('drawAnchorImage srcRect crop', () => {
  it('draws only the visible sub-rectangle for a raster crop (9-arg drawImage)', () => {
    const { ctx, drawImage } = spyCtx();
    const img = fakeImg(2860, 1368); // sample-27 PNG native pixel size
    // sample-27: left 0.3256, right 0.03829, no vertical crop.
    const a = anchor({ srcRect: { l: 0.3256, t: 0, r: 0.03829, b: 0 } });

    drawAnchorImage(ctx, img, a, 10, 20, 305, 229);

    expect(drawImage).toHaveBeenCalledTimes(1);
    const call = drawImage.mock.calls[0];
    expect(call).toHaveLength(9); // img + (sx,sy,sw,sh) + (dx,dy,dw,dh)
    const [, sx, sy, sw, sh, dx, dy, dw, dh] = call;
    expect(sx).toBeCloseTo(0.3256 * 2860, 3); // skip the left 32.56%
    expect(sy).toBe(0);
    expect(sw).toBeCloseTo((1 - 0.3256 - 0.03829) * 2860, 3); // keep the middle band
    expect(sh).toBe(1368); // full height (no vertical crop)
    // Destination box is unchanged — the slice stretches to fill the anchor rect.
    expect([dx, dy, dw, dh]).toEqual([10, 20, 305, 229]);
  });

  it('clamps a crop that extends past the image and never produces a zero-size source', () => {
    const { ctx, drawImage } = spyCtx();
    // Pathological insets (sum > 1, negative) must clamp to a ≥1px source rect.
    const a = anchor({ srcRect: { l: 0.9, t: -0.2, r: 0.9, b: 1.5 } });

    drawAnchorImage(ctx, fakeImg(100, 100), a, 0, 0, 40, 40);

    const [, , , sw, sh] = drawImage.mock.calls[0];
    expect(sw).toBeGreaterThanOrEqual(1);
    expect(sh).toBeGreaterThanOrEqual(1);
  });

  it('draws the whole image (4-arg) when there is no crop', () => {
    const { ctx, drawImage } = spyCtx();
    drawAnchorImage(ctx, fakeImg(100, 100), anchor({}), 0, 0, 50, 50);
    expect(drawImage.mock.calls[0]).toHaveLength(5); // img + (dx,dy,dw,dh)
    expect(drawImage).toHaveBeenCalledWith(expect.anything(), 0, 0, 50, 50);
  });

  it('treats an all-zero srcRect as no crop (full draw)', () => {
    const { ctx, drawImage } = spyCtx();
    const a = anchor({ srcRect: { l: 0, t: 0, r: 0, b: 0 } });
    drawAnchorImage(ctx, fakeImg(100, 100), a, 0, 0, 50, 50);
    expect(drawImage.mock.calls[0]).toHaveLength(5);
  });

  it('skips the crop for a metafile (WMF) — it is rasterized to the display box', () => {
    const { ctx, drawImage } = spyCtx();
    const a = anchor({ mimeType: 'image/wmf', srcRect: { l: 0.3, t: 0, r: 0.1, b: 0 } });
    drawAnchorImage(ctx, fakeImg(200, 200), a, 0, 0, 50, 50);
    expect(drawImage.mock.calls[0]).toHaveLength(5); // full draw, not 9-arg
  });

  it('skips the crop for an EMF metafile too', () => {
    const { ctx, drawImage } = spyCtx();
    const a = anchor({ mimeType: 'image/emf', srcRect: { l: 0.2, t: 0.2, r: 0, b: 0 } });
    drawAnchorImage(ctx, fakeImg(200, 200), a, 0, 0, 50, 50);
    expect(drawImage.mock.calls[0]).toHaveLength(5);
  });
});

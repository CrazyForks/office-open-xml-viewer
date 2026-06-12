import { describe, it, expect } from 'vitest';
import {
  bevelHeightProfile,
  edt1d,
  gaussianBlur,
  computeBevelNormals,
  shadePixel,
  shadeParamsFor,
  lightDirFromRig,
  applyBevelShading,
  applyExtrusion,
  type BevelShadeParams,
  type BevelCtx,
} from './bevel-shading';

/** Minimal ImageData-backed BevelCtx for the compositor functions. */
function fakeCtx(w: number, h: number, fill?: (x: number, y: number) => [number, number, number, number]): BevelCtx & { data: Uint8ClampedArray } {
  const data = new Uint8ClampedArray(w * h * 4);
  if (fill) {
    for (let y = 0; y < h; y++)
      for (let x = 0; x < w; x++) {
        const [r, g, b, a] = fill(x, y);
        const o = (y * w + x) * 4;
        data[o] = r; data[o + 1] = g; data[o + 2] = b; data[o + 3] = a;
      }
  }
  return {
    data,
    canvas: { width: w, height: h },
    getImageData() {
      return { data: data.slice(), width: w, height: h } as unknown as ImageData;
    },
    putImageData(img: ImageData) {
      data.set(img.data);
    },
  };
}

describe('bevelHeightProfile', () => {
  it('circle profile rises as a quarter-circle: h(0)=0, h(w)=1, convex', () => {
    const w = 10;
    const p = bevelHeightProfile('circle', w);
    expect(p(0)).toBeCloseTo(0, 5);
    expect(p(w)).toBeCloseTo(1, 5);
    // Quarter circle h = sqrt(1-(1-t)^2): at t=0.5 → sqrt(1-0.25)=0.866 (convex,
    // above the linear 0.5).
    expect(p(w * 0.5)).toBeCloseTo(Math.sqrt(1 - 0.25), 5);
    expect(p(w * 0.5)).toBeGreaterThan(0.5);
  });

  it('hardEdge is a linear chamfer: h(t) = t/w', () => {
    const w = 8;
    const p = bevelHeightProfile('hardEdge', w);
    expect(p(0)).toBeCloseTo(0, 5);
    expect(p(w)).toBeCloseTo(1, 5);
    expect(p(w * 0.5)).toBeCloseTo(0.5, 5);
  });

  it('relaxedInset (default) is a smooth S-curve in [0,1]', () => {
    const w = 10;
    const p = bevelHeightProfile('relaxedInset', w);
    expect(p(0)).toBeCloseTo(0, 5);
    expect(p(w)).toBeCloseTo(1, 5);
    // Monotonic increasing.
    expect(p(w * 0.25)).toBeLessThan(p(w * 0.75));
  });

  it('clamps distance beyond the band to the flat top (h=1)', () => {
    const w = 5;
    const p = bevelHeightProfile('circle', w);
    expect(p(w * 2)).toBeCloseTo(1, 5);
  });

  it('zero width yields a flat profile (always 1, no bevel)', () => {
    const p = bevelHeightProfile('circle', 0);
    expect(p(0)).toBeCloseTo(1, 5);
    expect(p(10)).toBeCloseTo(1, 5);
  });
});

describe('edt1d (squared Euclidean distance transform, Felzenszwalb)', () => {
  it('computes squared distance to the nearest zero on a 1D line', () => {
    // f: 0 at index 2, +inf elsewhere → squared distance to index 2.
    const INF = 1e20;
    const f = [INF, INF, 0, INF, INF];
    const out = edt1d(f);
    expect(Array.from(out)).toEqual([4, 1, 0, 1, 4]);
  });

  it('handles two seeds, taking the nearer', () => {
    const INF = 1e20;
    const f = [0, INF, INF, INF, 0];
    const out = edt1d(f);
    expect(Array.from(out)).toEqual([0, 1, 4, 1, 0]);
  });
});

describe('lightDirFromRig', () => {
  it('dir="t" points the light from the top (negative screen-y), toward viewer', () => {
    const L = lightDirFromRig('threePt', 't');
    expect(L.y).toBeLessThan(0); // light comes from above (screen y down)
    expect(L.z).toBeGreaterThan(0); // and from in front of the surface
    // unit length
    expect(Math.hypot(L.x, L.y, L.z)).toBeCloseTo(1, 5);
  });

  it('dir="l" comes from the left (negative x)', () => {
    const L = lightDirFromRig('threePt', 'l');
    expect(L.x).toBeLessThan(0);
  });

  it('dir="br" comes from bottom-right (positive x, positive y)', () => {
    const L = lightDirFromRig('threePt', 'br');
    expect(L.x).toBeGreaterThan(0);
    expect(L.y).toBeGreaterThan(0);
  });
});

describe('shadePixel (Lambert + weak specular)', () => {
  const params: BevelShadeParams = {
    light: lightDirFromRig('threePt', 't'),
    material: 'matte',
    ambient: 0.55,
    diffuse: 0.45,
    specular: 0.0,
    shininess: 8,
  };

  it('a flat top-facing normal (0,0,1) gets a mid factor near ambient+diffuse·(L·N)', () => {
    const f = shadePixel({ x: 0, y: 0, z: 1 }, params);
    // L·N = L.z. factor = ambient + diffuse*max(0,L.z).
    const expected = params.ambient + params.diffuse * params.light.z;
    expect(f).toBeCloseTo(expected, 5);
  });

  it('a normal tilted toward the light is brighter than one tilted away', () => {
    const toward = shadePixel(normalize(params.light.x, params.light.y, 1), params);
    const away = shadePixel(normalize(-params.light.x, -params.light.y, 1), params);
    expect(toward).toBeGreaterThan(away);
  });

  it('plastic material adds a specular highlight (brighter peak than matte)', () => {
    const matteP = { ...params, specular: 0.0 };
    const plasticP = { ...params, specular: 0.4, material: 'plastic' as const };
    // Normal exactly on the light direction → max specular.
    const n = normalize(params.light.x, params.light.y, params.light.z);
    expect(shadePixel(n, plasticP)).toBeGreaterThan(shadePixel(n, matteP));
  });

  it('clamps the factor to >= 0 (no negative)', () => {
    const f = shadePixel({ x: 0, y: 0, z: -1 }, params);
    expect(f).toBeGreaterThanOrEqual(0);
  });
});

describe('computeBevelNormals (distance-driven tilt, coverage-gradient azimuth)', () => {
  it('produces (0,0,1) on the flat interior and tilted normals on the band', () => {
    // 7x7 filled square, bevel width 2. Interior pixels (far from edge) flat;
    // band pixels tilt outward.
    const W = 7;
    const H = 7;
    const alpha = new Uint8ClampedArray(W * H).fill(255);
    const { normals, bandMask } = computeBevelNormals(alpha, W, H, 2, 'circle', 1);
    // Centre pixel (3,3) is interior → flat normal.
    const ci = (3 * W + 3) * 3;
    expect(normals[ci + 2]).toBeGreaterThan(0.9); // nz ≈ 1
    expect(bandMask[3 * W + 3]).toBe(0); // not in band
    // An edge-band pixel (e.g. (3,0), on the top edge, distance 1 < band 2) is
    // in the band and its normal tilts upward (ny < 0 → faces up toward the top).
    const ei = (0 * W + 3) * 3;
    expect(bandMask[0 * W + 3]).toBe(1);
    expect(normals[ei + 1]).toBeLessThan(0); // ny<0 faces up/out at the top edge
  });

  it('ignores fully transparent pixels (outside the silhouette)', () => {
    const W = 5;
    const H = 5;
    const alpha = new Uint8ClampedArray(W * H); // all zero
    const { bandMask } = computeBevelNormals(alpha, W, H, 2, 'circle', 1);
    expect(bandMask.every((v) => v === 0)).toBe(true);
  });
});

/**
 * Anti-aliased ellipse coverage mask (mimics a Canvas `clip()` rasterisation:
 * interior 255, exterior 0, a 1-px fractional rim). Used to exercise the bevel
 * lip on a curved silhouette with a wide band — the case that exposed the
 * EDT-gradient faceting regression (PR #410 → this fix).
 */
function ellipseAlpha(W: number, H: number, ss = 4): Uint8ClampedArray {
  const a = new Uint8ClampedArray(W * H);
  const cx = W / 2;
  const cy = H / 2;
  const rx = W / 2 - 1;
  const ry = H / 2 - 1;
  for (let y = 0; y < H; y++)
    for (let x = 0; x < W; x++) {
      let cnt = 0;
      for (let sy = 0; sy < ss; sy++)
        for (let sx = 0; sx < ss; sx++) {
          const px = x + (sx + 0.5) / ss - 0.5;
          const py = y + (sy + 0.5) / ss - 0.5;
          if (((px - cx) / rx) ** 2 + ((py - cy) / ry) ** 2 <= 1) cnt++;
        }
      a[y * W + x] = Math.round((255 * cnt) / (ss * ss));
    }
  return a;
}

/**
 * Anti-aliased annulus (ring) coverage: outer ellipse minus a concentric inner
 * ellipse. A wide bevel band reaching toward the ring's medial axis is the
 * worst case for the OLD EDT-gradient azimuth (the nearest-feature flips across
 * the medial ridge). Used by the scale-sweep below to prove the coverage-gradient
 * azimuth tracks a SINGLE rim smoothly even on a thin ring at every scale.
 */
function ringAlpha(W: number, H: number, thickFrac: number, ss = 4): Uint8ClampedArray {
  const a = new Uint8ClampedArray(W * H);
  const cx = W / 2;
  const cy = H / 2;
  const rx = W / 2 - 1;
  const ry = H / 2 - 1;
  const irx = rx * (1 - thickFrac);
  const iry = ry * (1 - thickFrac);
  for (let y = 0; y < H; y++)
    for (let x = 0; x < W; x++) {
      let cnt = 0;
      for (let sy = 0; sy < ss; sy++)
        for (let sx = 0; sx < ss; sx++) {
          const px = x + (sx + 0.5) / ss - 0.5;
          const py = y + (sy + 0.5) / ss - 0.5;
          const outer = ((px - cx) / rx) ** 2 + ((py - cy) / ry) ** 2 <= 1;
          const inner = ((px - cx) / irx) ** 2 + ((py - cy) / iry) ** 2 <= 1;
          if (outer && !inner) cnt++;
        }
      a[y * W + x] = Math.round((255 * cnt) / (ss * ss));
    }
  return a;
}

describe('gaussianBlur (3-box cascade, zero-padded)', () => {
  it('smooths a step and falls off across a clipped (canvas-edge) boundary', () => {
    // A solid block touching the left edge: coverage 1 for x<W/2, 0 otherwise.
    const W = 32;
    const H = 4;
    const src = new Float64Array(W * H);
    for (let y = 0; y < H; y++) for (let x = 0; x < W / 2; x++) src[y * W + x] = 1;
    const out = gaussianBlur(src, W, H, 4);
    // Monotonic falloff across the interior step (x from W/2-4 to W/2+4).
    const row = 1 * W;
    for (let x = W / 2 - 3; x < W / 2 + 3; x++) {
      expect(out[row + x]).toBeGreaterThan(out[row + x + 1]);
    }
    // Zero padding (NOT clamp-to-edge): the solid block flush with x=0 falls off
    // toward the left border, so the value AT the border is below the interior
    // plateau — this is what gives a clipped silhouette an outward ∇C lip.
    expect(out[row + 0]).toBeLessThan(out[row + 4]);
  });
});

describe('computeBevelNormals — scale-invariant azimuth (issue #410→#413→#415→here)', () => {
  // DECISIVE SCALE SWEEP. The bevel facet bug has now slipped through THREE times,
  // each time at a LARGER raster scale, because every prior fix was a post-filter
  // on the EDT gradient whose radius was a fraction of the (growing) band:
  //   1. #410 mesh cracks — grew with shape size.
  //   2. #413 height-field blur — fixed devScale ≤ 2, regressed at 4.
  //   3. #415 tangential normal blur (radius 0.25·band, gated ≥24px) — fixed
  //      devScale 4, REGRESSED at devScale 8 (the user's real DPR-8 render).
  // Each prior test pinned a SINGLE scale (e.g. #415 asserted band=128 / devScale 4
  // only), so the next larger scale walked straight through. This suite instead
  // PARAMETRISES devScale ∈ {1,2,4,8} and asserts ONE scale-independent threshold
  // for every scale, so a regression at any DPR fails here.
  //
  // The redesign sources the lip AZIMUTH from ∇(Gaussian-blurred coverage), which
  // is analytically C^∞ at every scale (no EDT Voronoi structure enters the
  // direction), so smoothness is invariant — and in fact IMPROVES with scale.
  // Measured ceilings (production code): ellipse & ring max 1-step azimuth jump
  // ≤ 0.0094 rad at devScale 1, shrinking to ≤ 0.0030 at devScale 8; highlight
  // luminance 2nd-difference ≤ 0.0011 → ≤ 0.0002. For comparison the RAW EDT
  // gradient (pre-patch) jumps 0.12–0.24 rad — a 7°–13° chord — at all scales.
  // A single ceiling of 0.02 rad / 0.003 lum sits well above the redesign at
  // every scale yet an order of magnitude below the raw facet.
  const DEV_SCALES = [1, 2, 4, 8];
  // sample-11 slide 6 body offscreen at devScale 1 ≈ 397×503 px, hardEdge band 32.
  const BASE = { W: 397, H: 503, band: 32 };
  const AZ_JUMP_MAX = 0.02; // rad, scale-independent
  const LUM_2ND_DIFF_MAX = 0.003; // scale-independent

  function ringAzimuthAndHighlight(alpha: Uint8ClampedArray, W: number, H: number, band: number) {
    const { normals, bandMask } = computeBevelNormals(alpha, W, H, band, 'hardEdge', band / 2);
    const params = shadeParamsFor('matte', lightDirFromRig('threePt', 't'));
    const faceFactor = shadePixel({ x: 0, y: 0, z: 1 }, params) || 1;
    const cx = W / 2;
    const cy = H / 2;
    const rx = W / 2 - 1;
    const ry = H / 2 - 1;
    const inset = band * 0.5; // band midline, where the lit/shadow terminator sits
    const az: number[] = [];
    const lum: number[] = [];
    // 0.1° steps so a chord spanning several degrees registers as a run of equal
    // azimuths broken by a sudden jump.
    for (let deg = 0; deg < 360; deg += 0.1) {
      const ang = (deg * Math.PI) / 180;
      const x = Math.round(cx + Math.cos(ang) * (rx - inset));
      const y = Math.round(cy + Math.sin(ang) * (ry - inset));
      if (x < 0 || y < 0 || x >= W || y >= H) continue;
      const i = y * W + x;
      if (!bandMask[i]) continue;
      const nx = normals[i * 3];
      const ny = normals[i * 3 + 1];
      const nz = normals[i * 3 + 2];
      if (Math.hypot(nx, ny) < 1e-3) continue;
      az.push(Math.atan2(ny, nx));
      lum.push(shadePixel({ x: nx, y: ny, z: nz }, params) / faceFactor);
    }
    let maxJump = 0;
    let net = 0;
    for (let i = 1; i < az.length; i++) {
      let d = az[i] - az[i - 1];
      while (d > Math.PI) d -= 2 * Math.PI;
      while (d < -Math.PI) d += 2 * Math.PI;
      maxJump = Math.max(maxJump, Math.abs(d));
      net += d;
    }
    let maxLum2 = 0;
    for (let i = 2; i < lum.length; i++) {
      maxLum2 = Math.max(maxLum2, Math.abs(lum[i] - 2 * lum[i - 1] + lum[i - 2]));
    }
    return { maxJump, maxLum2, net, n: az.length };
  }

  for (const ds of DEV_SCALES) {
    const W = BASE.W * ds;
    const H = BASE.H * ds;
    const band = BASE.band * ds;

    it(`ellipse lip is smooth at devScale ${ds} (band ${band}px)`, () => {
      const r = ringAzimuthAndHighlight(ellipseAlpha(W, H), W, H, band);
      expect(r.n).toBeGreaterThan(3000);
      // (a) no facet: largest single-step azimuth jump under the shared ceiling.
      expect(r.maxJump).toBeLessThan(AZ_JUMP_MAX);
      // (b) highlight has no chorded kink: tangential 2nd-difference bounded.
      expect(r.maxLum2).toBeLessThan(LUM_2ND_DIFF_MAX);
      // (c) the azimuth makes one full turn — it tracks the rim, not a few facets.
      expect(Math.abs(Math.abs(r.net) - 2 * Math.PI)).toBeLessThan(0.5);
    });

    it(`thin-ring lip is smooth at devScale ${ds} (band ${band}px)`, () => {
      // A ring whose band reaches toward the medial axis — the medial-ridge case
      // that breaks an EDT-nearest-feature azimuth. The coverage σ (0.25·band) is
      // small enough to stay local to the OUTER rim, so it stays smooth.
      const r = ringAzimuthAndHighlight(ringAlpha(W, H, 0.18), W, H, band);
      expect(r.n).toBeGreaterThan(3000);
      expect(r.maxJump).toBeLessThan(AZ_JUMP_MAX);
      expect(r.maxLum2).toBeLessThan(LUM_2ND_DIFF_MAX);
      expect(Math.abs(Math.abs(r.net) - 2 * Math.PI)).toBeLessThan(0.5);
    });
  }

  it('smoothness does NOT degrade as scale grows (the property prior patches lacked)', () => {
    // The defining scale-invariance assertion: the max azimuth jump at devScale 8
    // must be NO WORSE than at devScale 1. Every prior EDT-gradient patch had the
    // opposite behaviour (smooth at the tuned scale, faceted above it); this would
    // be RED for all of them and is GREEN for the coverage-gradient azimuth.
    const small = ringAzimuthAndHighlight(
      ellipseAlpha(BASE.W * 1, BASE.H * 1),
      BASE.W * 1,
      BASE.H * 1,
      BASE.band * 1,
    );
    const large = ringAzimuthAndHighlight(
      ellipseAlpha(BASE.W * 8, BASE.H * 8),
      BASE.W * 8,
      BASE.H * 8,
      BASE.band * 8,
    );
    expect(large.maxJump).toBeLessThanOrEqual(small.maxJump + 1e-6);
  });
});

describe('applyExtrusion (side-wall sweep, §20.1.5.12)', () => {
  it('fills the swept band behind the front face with the wall colour', () => {
    // 20x20 with an opaque 8x8 block at (6,6)-(13,13). Push it right+down by
    // (4,4): the side wall is the L-shaped band uncovered to the right/below.
    const W = 20;
    const H = 20;
    const ctx = fakeCtx(W, H, (x, y) =>
      x >= 6 && x < 14 && y >= 6 && y < 14 ? [10, 20, 200, 255] : [0, 0, 0, 0],
    );
    applyExtrusion(ctx, { offsetX: 4, offsetY: 4, rgb: [50, 50, 50] });
    // A pixel just below-right of the block (e.g. (15,15)) should now be wall.
    const wall = (15 * W + 15) * 4;
    expect(ctx.data[wall + 3]).toBe(255);
    expect(ctx.data[wall]).toBe(50);
    // The front face is untouched.
    const face = (8 * W + 8) * 4;
    expect(ctx.data[face]).toBe(10);
    expect(ctx.data[face + 2]).toBe(200);
  });

  it('is a no-op for a sub-pixel offset (face-on camera)', () => {
    const W = 10;
    const H = 10;
    const ctx = fakeCtx(W, H, (x, y) => (x >= 3 && x < 7 && y >= 3 && y < 7 ? [10, 10, 10, 255] : [0, 0, 0, 0]));
    const before = ctx.data.slice();
    applyExtrusion(ctx, { offsetX: 0.2, offsetY: 0.1, rgb: [99, 99, 99] });
    expect(Array.from(ctx.data)).toEqual(Array.from(before));
  });
});

describe('applyBevelShading (end-to-end on a filled square)', () => {
  it('brightens the top lip and darkens the bottom lip under a top light', () => {
    const W = 24;
    const H = 24;
    // Mid-grey opaque square filling the canvas.
    const ctx = fakeCtx(W, H, () => [120, 120, 120, 255]);
    applyBevelShading(ctx, {
      widthPx: 5,
      heightPx: 4,
      prst: 'circle',
      material: 'matte',
      light: lightDirFromRig('threePt', 't'),
    });
    const lumAt = (x: number, y: number) => {
      const o = (y * W + x) * 4;
      return ctx.data[o] * 0.299 + ctx.data[o + 1] * 0.587 + ctx.data[o + 2] * 0.114;
    };
    // Top lip (y=1) lit brighter than the flat centre; bottom lip (y=22) darker.
    expect(lumAt(12, 1)).toBeGreaterThan(lumAt(12, 12));
    expect(lumAt(12, 22)).toBeLessThan(lumAt(12, 12));
  });
});

function normalize(x: number, y: number, z: number) {
  const m = Math.hypot(x, y, z) || 1;
  return { x: x / m, y: y / m, z: z / m };
}

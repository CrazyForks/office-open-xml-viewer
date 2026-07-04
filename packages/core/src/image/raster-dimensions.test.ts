import { describe, it, expect } from 'vitest';
import {
  sniffRasterDimensions,
  rasterExceedsBudget,
  rasterHeaderExceedsBudget,
} from './raster-dimensions.js';
import { MAX_RASTER_DIMENSION, MAX_RASTER_PIXELS } from './pixel-budget.js';

// ── Raster header dimension sniff + budget (RB1 decode-bomb guard) ───────────
//
// These fixtures are SYNTHETIC: each carries a valid header declaring enormous
// pixel dimensions but only a handful of bytes of "pixel data" — the shape of a
// decompression bomb (a tiny compressed blip that decodes to a multi-GB RGBA
// surface). The guard must recognize the declared dimensions from the header
// alone and report the image as over budget, WITHOUT decoding it.

/** Big-endian u32 into a byte array at `o`. */
function putBeU32(b: Uint8Array, o: number, v: number): void {
  b[o] = (v >>> 24) & 0xff;
  b[o + 1] = (v >>> 16) & 0xff;
  b[o + 2] = (v >>> 8) & 0xff;
  b[o + 3] = v & 0xff;
}
/** Little-endian u16. */
function putLeU16(b: Uint8Array, o: number, v: number): void {
  b[o] = v & 0xff;
  b[o + 1] = (v >>> 8) & 0xff;
}
/** Little-endian u32. */
function putLeU32(b: Uint8Array, o: number, v: number): void {
  b[o] = v & 0xff;
  b[o + 1] = (v >>> 8) & 0xff;
  b[o + 2] = (v >>> 16) & 0xff;
  b[o + 3] = (v >>> 24) & 0xff;
}
/** Big-endian u16. */
function putBeU16(b: Uint8Array, o: number, v: number): void {
  b[o] = (v >>> 8) & 0xff;
  b[o + 1] = v & 0xff;
}

// ── Header builders (declare `w × h`, carry almost no payload) ───────────────

function pngHeader(w: number, h: number): Uint8Array {
  // 8-byte signature + IHDR chunk (length + "IHDR" + W + H + …). We only need
  // through offset 24; a couple of trailing bytes stand in for the rest.
  const b = new Uint8Array(26);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  putBeU32(b, 8, 13); // IHDR length
  b.set([0x49, 0x48, 0x44, 0x52], 12); // "IHDR"
  putBeU32(b, 16, w);
  putBeU32(b, 20, h);
  return b;
}

function gifHeader(w: number, h: number): Uint8Array {
  const b = new Uint8Array(13);
  b.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61], 0); // "GIF89a"
  putLeU16(b, 6, w);
  putLeU16(b, 8, h);
  return b;
}

function bmpHeader(w: number, h: number): Uint8Array {
  // "BM" + 14-byte file header, then a 40-byte BITMAPINFOHEADER.
  const b = new Uint8Array(54);
  b.set([0x42, 0x4d], 0); // "BM"
  putLeU32(b, 14, 40); // biSize = BITMAPINFOHEADER
  putLeU32(b, 18, w >>> 0); // signed i32, positive here
  putLeU32(b, 22, h >>> 0);
  return b;
}

function webpVp8xHeader(w: number, h: number): Uint8Array {
  // RIFF…WEBP + "VP8X" chunk with 24-bit canvas width-1 / height-1.
  const b = new Uint8Array(30);
  b.set([0x52, 0x49, 0x46, 0x46], 0); // "RIFF"
  putLeU32(b, 4, 0); // file size (ignored by the sniff)
  b.set([0x57, 0x45, 0x42, 0x50], 8); // "WEBP"
  b.set([0x56, 0x50, 0x38, 0x58], 12); // "VP8X"
  putLeU32(b, 16, 10); // VP8X payload size (ignored)
  // flags byte at 20; canvas width-1 (24-bit LE) at 24, height-1 at 27.
  const w1 = w - 1;
  const h1 = h - 1;
  b[24] = w1 & 0xff;
  b[25] = (w1 >>> 8) & 0xff;
  b[26] = (w1 >>> 16) & 0xff;
  b[27] = h1 & 0xff;
  b[28] = (h1 >>> 8) & 0xff;
  b[29] = (h1 >>> 16) & 0xff;
  return b;
}

/** A minimal JPEG: SOI, an APP0 segment, then an SOF0 declaring `w × h`. */
function jpegHeader(w: number, h: number, padApp0 = 0): Uint8Array {
  // SOI (2) + APP0 marker (2) + APP0 length (2) + APP0 payload (padApp0) +
  // SOF0 marker (2) + SOF0 length (2) + [prec(1) h(2) w(2) …].
  const app0Len = 2 + padApp0; // length field counts itself
  const sof0PayloadLen = 2 + 1 + 2 + 2; // len + prec + h + w
  const size = 2 + 2 + app0Len + 2 + sof0PayloadLen;
  const b = new Uint8Array(size);
  let i = 0;
  b[i++] = 0xff;
  b[i++] = 0xd8; // SOI
  b[i++] = 0xff;
  b[i++] = 0xe0; // APP0
  putBeU16(b, i, app0Len);
  i += 2 + padApp0; // skip APP0 payload
  b[i++] = 0xff;
  b[i++] = 0xc0; // SOF0
  putBeU16(b, i, sof0PayloadLen);
  i += 2;
  b[i++] = 8; // precision
  putBeU16(b, i, h);
  i += 2;
  putBeU16(b, i, w);
  i += 2;
  return b;
}

describe('sniffRasterDimensions — reads declared pixel dimensions from the header', () => {
  it('reads PNG IHDR dimensions', () => {
    expect(sniffRasterDimensions(pngHeader(1920, 1080))).toEqual({ width: 1920, height: 1080 });
    expect(sniffRasterDimensions(pngHeader(60000, 60000))).toEqual({ width: 60000, height: 60000 });
  });

  it('reads GIF logical-screen dimensions', () => {
    expect(sniffRasterDimensions(gifHeader(800, 600))).toEqual({ width: 800, height: 600 });
  });

  it('reads BMP BITMAPINFOHEADER dimensions (and |height| for top-down)', () => {
    expect(sniffRasterDimensions(bmpHeader(1024, 768))).toEqual({ width: 1024, height: 768 });
    // Negative height (top-down bitmap) is budgeted by magnitude.
    const b = bmpHeader(1024, 0);
    putLeU32(b, 22, (-768 >>> 0) as number);
    expect(sniffRasterDimensions(b)).toEqual({ width: 1024, height: 768 });
  });

  it('reads WebP VP8X canvas dimensions', () => {
    expect(sniffRasterDimensions(webpVp8xHeader(4000, 3000))).toEqual({ width: 4000, height: 3000 });
  });

  it('reads JPEG SOF dimensions, even past an APP0 segment', () => {
    expect(sniffRasterDimensions(jpegHeader(3264, 2448))).toEqual({ width: 3264, height: 2448 });
    // With a large APP0 (EXIF/ICC stand-in) the walker still reaches the SOF.
    expect(sniffRasterDimensions(jpegHeader(3264, 2448, 400))).toEqual({
      width: 3264,
      height: 2448,
    });
  });

  it('returns null for unrecognized / too-short / SVG headers (fail-open)', () => {
    expect(sniffRasterDimensions(new Uint8Array([]))).toBeNull();
    expect(sniffRasterDimensions(new Uint8Array([0x00, 0x01, 0x02, 0x03]))).toBeNull();
    // SVG starts with "<?xml" or "<svg" — not a recognized raster.
    expect(sniffRasterDimensions(new TextEncoder().encode('<svg width="9"></svg>'))).toBeNull();
    // PNG signature but a first chunk that is not IHDR → don't trust it.
    const notIhdr = pngHeader(10, 10);
    notIhdr[12] = 0x00; // corrupt "IHDR" → "\0HDR"
    expect(sniffRasterDimensions(notIhdr)).toBeNull();
  });
});

describe('rasterExceedsBudget — enforces the shared pixel budget', () => {
  it('accepts in-budget dimensions', () => {
    expect(rasterExceedsBudget({ width: 1920, height: 1080 })).toBe(false);
    // A ~35 MP scan (within the 64 MP budget and under the per-axis cap).
    expect(rasterExceedsBudget({ width: 7016, height: 4961 })).toBe(false);
  });

  it('rejects an over-megapixel image within the per-axis cap', () => {
    // Both axes ≤ 32767 but the product blows the 64 MP budget.
    const w = 30000;
    const h = 30000;
    expect(w).toBeLessThanOrEqual(MAX_RASTER_DIMENSION);
    expect(h).toBeLessThanOrEqual(MAX_RASTER_DIMENSION);
    expect(w * h).toBeGreaterThan(MAX_RASTER_PIXELS);
    expect(rasterExceedsBudget({ width: w, height: h })).toBe(true);
  });

  it('rejects an over-dimension image even when it would be within the MP budget', () => {
    // 40000 × 1: only 40 000 px (well under 64 MP) but past the 32767 axis cap,
    // so it could never be drawn to a canvas — reject it.
    expect(rasterExceedsBudget({ width: 40000, height: 1 })).toBe(true);
  });

  it('treats non-positive or non-finite dimensions as over budget (untrustworthy header)', () => {
    expect(rasterExceedsBudget({ width: 0, height: 100 })).toBe(true);
    expect(rasterExceedsBudget({ width: -5, height: 100 })).toBe(true);
    expect(rasterExceedsBudget({ width: NaN, height: 100 })).toBe(true);
  });

  it('accepts exactly at the caps (boundary)', () => {
    expect(rasterExceedsBudget({ width: MAX_RASTER_DIMENSION, height: 1 })).toBe(false);
    // A 1 × MAX_RASTER_PIXELS strip is exactly the MP budget (but > axis cap):
    // use a squarer at-budget case instead.
    expect(rasterExceedsBudget({ width: 8192, height: 8192 })).toBe(false); // 64 MP exactly
    expect(8192 * 8192).toBe(MAX_RASTER_PIXELS);
  });
});

describe('rasterHeaderExceedsBudget — end-to-end neutralization of decode bombs', () => {
  it('flags a 60000×60000 PNG bomb (tiny payload, huge declared size)', () => {
    const bomb = pngHeader(60000, 60000);
    // The fixture is a couple dozen bytes — a real decode would be ~14 GB RGBA.
    expect(bomb.length).toBeLessThan(64);
    expect(rasterHeaderExceedsBudget(bomb)).toBe(true);
  });

  it('flags GIF / BMP / WebP / JPEG bombs alike', () => {
    expect(rasterHeaderExceedsBudget(gifHeader(50000, 50000))).toBe(true);
    expect(rasterHeaderExceedsBudget(bmpHeader(40000, 40000))).toBe(true);
    expect(rasterHeaderExceedsBudget(webpVp8xHeader(16384, 16384))).toBe(true); // 268 MP
    expect(rasterHeaderExceedsBudget(jpegHeader(60000, 60000))).toBe(true);
  });

  it('does NOT flag a normal, in-budget image (no false positive)', () => {
    expect(rasterHeaderExceedsBudget(pngHeader(1920, 1080))).toBe(false);
    expect(rasterHeaderExceedsBudget(jpegHeader(4032, 3024))).toBe(false);
    expect(rasterHeaderExceedsBudget(gifHeader(500, 500))).toBe(false);
  });

  it('does NOT flag an unrecognized header (fail-open — the caller decodes normally)', () => {
    expect(rasterHeaderExceedsBudget(new Uint8Array([0xde, 0xad, 0xbe, 0xef]))).toBe(false);
    expect(rasterHeaderExceedsBudget(new TextEncoder().encode('<svg/>'))).toBe(false);
  });
});

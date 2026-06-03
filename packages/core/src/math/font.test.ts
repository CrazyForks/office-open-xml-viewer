import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseMathFont, type MathFont } from './font';

let font: MathFont;
beforeAll(() => {
  const url = new URL('../../assets/LatinModernMath.otf', import.meta.url);
  const buf = readFileSync(fileURLToPath(url));
  font = parseMathFont(
    buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
  );
});

describe('parseMathFont', () => {
  it('reads unitsPerEm', () => {
    expect(font.unitsPerEm).toBe(1000);
  });
  it('exposes positive ascent/descent', () => {
    expect(font.ascent).toBeGreaterThan(0);
    expect(font.descent).toBeGreaterThan(0);
  });
  it('maps ASCII and Greek to advances', () => {
    const x = font.glyphForChar('x'.codePointAt(0)!);
    expect(x).toBeGreaterThan(0);
    expect(font.advance(x)).toBeGreaterThan(0);
    const sum = font.glyphForChar('∑'.codePointAt(0)!);
    expect(sum).toBeGreaterThan(0);
  });
  it('finds the MATH table offset', () => {
    expect(font.tableOffset('MATH')).toBeGreaterThan(0);
    expect(font.tableOffset('nope')).toBe(-1);
  });
});

import { describe, it, expect } from 'vitest';
import { cssTailFor, fontStackFor } from './renderer.js';

describe('fontStackFor — default Latin chain (regression)', () => {
  it('returns the Calibri/Carlito default chain for an unnamed cell', () => {
    const stack = fontStackFor(null);
    expect(stack.startsWith('"Calibri", "Carlito"')).toBe(true);
    // Arabic + non-CJK script fallbacks retained.
    expect(stack).toContain('"Noto Naskh Arabic"');
    expect(stack).toContain('"Noto Sans Hebrew"');
    expect(stack).toContain('"Noto Sans Thai"');
    expect(stack).toContain('"Noto Sans Devanagari"');
    expect(stack.endsWith('sans-serif')).toBe(true);
  });

  it('leads with the named Latin face, then the default chain', () => {
    const stack = fontStackFor('Arial');
    expect(stack.startsWith('"Arial", "Calibri", "Carlito"')).toBe(true);
  });
});

describe('fontStackFor — CJK language-specific Noto ordering', () => {
  it('Korean sans (Malgun Gothic) → Noto Sans KR leads the tail', () => {
    const tail = cssTailFor('Malgun Gothic');
    expect(tail.startsWith('"Noto Sans KR"')).toBe(true);
    expect(tail.indexOf('Noto Sans KR')).toBeLessThan(tail.indexOf('Noto Sans JP'));
    expect(tail.endsWith('sans-serif')).toBe(true);
  });

  it('Simplified Chinese serif (SimSun) → Noto Serif SC leads', () => {
    const tail = cssTailFor('SimSun');
    expect(tail.startsWith('"Noto Serif SC"')).toBe(true);
    expect(tail.endsWith('serif')).toBe(true);
  });

  it('Simplified Chinese sans (Microsoft YaHei) → Noto Sans SC leads', () => {
    expect(cssTailFor('Microsoft YaHei').startsWith('"Noto Sans SC"')).toBe(true);
  });

  it('Traditional Chinese (PMingLiU serif, JhengHei sans)', () => {
    expect(cssTailFor('PMingLiU').startsWith('"Noto Serif TC"')).toBe(true);
    expect(cssTailFor('Microsoft JhengHei').startsWith('"Noto Sans TC"')).toBe(true);
  });

  it('Japanese faces lead with Noto Sans JP (xlsx previously had no CJK fallback)', () => {
    const tail = cssTailFor('Meiryo');
    expect(tail.startsWith('"Noto Sans JP"')).toBe(true);
    // Still keeps the Latin metric substitutes after the CJK face.
    expect(tail).toContain('"Calibri"');
  });

  it('non-CJK named face falls back to the default chain', () => {
    expect(cssTailFor('Times New Roman')).toBe(fontStackFor(null));
  });
});

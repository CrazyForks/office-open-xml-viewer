import { describe, it, expect } from 'vitest';
import { hexToRgba } from './paint.js';

// hexToRgba is the colour pipeline's exit point: the pptx parser resolves a
// run's a:highlight (§21.1.2.3.4) to either a 6-char opaque hex or, when an
// alpha transform is present, an 8-char RRGGBBAA hex, and the renderer feeds it
// straight to hexToRgba before fillRect. Both shapes must round-trip; this is
// the only conversion the highlight box relies on.
describe('hexToRgba', () => {
  it('converts 6-char hex to opaque rgba', () => {
    expect(hexToRgba('FFFF00')).toBe('rgba(255,255,0,1)');
  });

  it('tolerates a leading # on 6-char hex', () => {
    expect(hexToRgba('#00FF00')).toBe('rgba(0,255,0,1)');
  });

  it('reads alpha from the 8-char RRGGBBAA form', () => {
    // AA = 80 → 128/255 ≈ 0.502 (a translucent marker from <a:alpha>).
    expect(hexToRgba('00FF0080')).toBe(`rgba(0,255,0,${128 / 255})`);
  });

  it('8-char alpha overrides the explicit alpha argument', () => {
    // The trailing AA wins so callers can pass colours uniformly.
    expect(hexToRgba('FF000000', 1)).toBe('rgba(255,0,0,0)');
  });

  it('applies the alpha argument only to 6-char hex', () => {
    expect(hexToRgba('FFFF00', 0.5)).toBe('rgba(255,255,0,0.5)');
  });
});

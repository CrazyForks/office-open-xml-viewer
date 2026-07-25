import { describe, expect, it } from 'vitest';
import { fontAdvanceBiasEm } from './font-advance-metrics.js';

describe('fontAdvanceBiasEm', () => {
  it('returns the public Word/PDF-calibrated Georgia allowance by resolved face', () => {
    expect(fontAdvanceBiasEm('Georgia')).toBe(0.0105);
    expect(fontAdvanceBiasEm('  "georgia" ')).toBe(0.0105);
    expect(fontAdvanceBiasEm('Times New Roman')).toBe(0);
    expect(fontAdvanceBiasEm('serif')).toBe(0);
    expect(fontAdvanceBiasEm(undefined)).toBe(0);
  });
});

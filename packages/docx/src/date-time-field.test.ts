import { describe, it, expect } from 'vitest';
import { resolveFieldText, type LineLayoutEnvironment } from './line-layout.js';
import type { FieldRun } from './types';

// ECMA-376 §17.16.5.16 DATE / §17.16.5.72 TIME — resolveFieldText formats the
// injected `currentDateMs` through the field's `\@` picture (§17.16.4.1), falling
// back to the authored cached result when there is no picture or an unsupported
// token. Reference instant: Tue Jan 3 2006 17:28:34 (matches the spec example).
const REF_MS = new Date(2006, 0, 3, 17, 28, 34).getTime();

function state(currentDateMs?: number): LineLayoutEnvironment {
  return { currentDateMs, totalPages: 1, pageIndex: 0 };
}

function field(fieldType: string, instruction: string, fallbackText = ''): FieldRun {
  return { fieldType, instruction, fallbackText } as unknown as FieldRun;
}

describe('resolveFieldText — DATE/TIME (§17.16.5.16 / §17.16.5.72)', () => {
  it('formats a TIME field YYYY picture to the injected year (sample-28 footer)', () => {
    // sample-28: ` TIME  \@ "YYYY"  \* MERGEFORMAT ` cached "2019".
    const f = field('time', ' TIME  \\@ "YYYY"  \\* MERGEFORMAT ', '2019');
    expect(resolveFieldText(f, state(REF_MS))).toBe('2006');
    // The cached "2019" is NOT emitted — the field recomputes.
    expect(resolveFieldText(f, state(REF_MS))).not.toBe('2019');
  });

  it('formats a DATE field with a full picture', () => {
    const f = field('date', ' DATE \\@ "dddd, MMMM dd, yyyy" ', 'cached');
    expect(resolveFieldText(f, state(REF_MS))).toBe('Tuesday, January 03, 2006');
  });

  it('falls back to the cached result when the field has no \\@ picture', () => {
    const f = field('date', ' DATE \\* MERGEFORMAT ', 'Jan 2, 2019');
    expect(resolveFieldText(f, state(REF_MS))).toBe('Jan 2, 2019');
  });

  it('falls back to the cached result for an unsupported picture token', () => {
    // `bbbb` (Thai Buddhist era) is unimplemented — keep the cache.
    const f = field('time', ' TIME \\@ "bbbb" ', '2562');
    expect(resolveFieldText(f, state(REF_MS))).toBe('2562');
  });

  it('defaults to the real current time when currentDateMs is absent', () => {
    const f = field('date', ' DATE \\@ "yyyy" ', 'cached');
    const nowYear = String(new Date().getFullYear());
    expect(resolveFieldText(f, state(undefined))).toBe(nowYear);
  });

  it('leaves non-date/time fields (fallback text) untouched', () => {
    const f = field('other', ' AUTHOR ', 'Jane Doe');
    expect(resolveFieldText(f, state(REF_MS))).toBe('Jane Doe');
  });
});

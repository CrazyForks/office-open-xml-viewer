import { describe, it, expect } from 'vitest';
import { excelSerialToUtcDate } from './excel-date';

// Helper: read back the calendar date a serial maps to, as YYYY-MM-DD in UTC.
function iso(serial: number, date1904: boolean): string {
  const d = excelSerialToUtcDate(serial, date1904);
  return `${d.getUTCFullYear().toString().padStart(4, '0')}-${(d.getUTCMonth() + 1)
    .toString()
    .padStart(2, '0')}-${d.getUTCDate().toString().padStart(2, '0')}`;
}

describe('excelSerialToUtcDate — 1900 date system (ECMA-376 §18.17.4.1)', () => {
  it('maps serial 1 to 1900-01-01 (base date 1899-12-30 is serial 0)', () => {
    expect(iso(1, false)).toBe('1900-01-01');
  });

  it('applies the Lotus 1900-02-29 leap-year bug: serial < 60 is shifted +1 day', () => {
    // Serial 59 is the last serial before the phantom 1900-02-29. Without the
    // +1 compat adjustment a naive epoch would render it 1900-02-27; Excel
    // renders 1900-02-28.
    expect(iso(59, false)).toBe('1900-02-28');
  });

  it('renders serial 61 as 1900-03-01 (the day after the phantom leap day)', () => {
    // Serial 60 is the phantom 1900-02-29 (no adjustment); serial 61 is the
    // first serial ≥ 60 and must land on 1900-03-01.
    expect(iso(61, false)).toBe('1900-03-01');
  });

  it('serial 60 keeps the pre-existing core behaviour (no +1 shift, → 1900-02-28)', () => {
    // 1900-02-29 does not exist in the proleptic Gregorian calendar. The
    // existing core `formatExcelDate` treated serials ≥ 60 without the +1
    // shift, so serial 60 lands on 1900-02-28. We deliberately preserve that.
    expect(iso(60, false)).toBe('1900-02-28');
  });

  it('maps a modern serial correctly (45292 → 2024-01-01)', () => {
    expect(iso(45292, false)).toBe('2024-01-01');
  });

  it('preserves the fractional time-of-day component', () => {
    // Serial 1.5 = midday on 1900-01-01.
    const d = excelSerialToUtcDate(1.5, false);
    expect(d.getUTCFullYear()).toBe(1900);
    expect(d.getUTCMonth() + 1).toBe(1);
    expect(d.getUTCDate()).toBe(1);
    expect(d.getUTCHours()).toBe(12);
  });
});

describe('excelSerialToUtcDate — 1904 date system (ECMA-376 §18.17.4.1)', () => {
  it('maps serial 0 to the 1904 base date 1904-01-01', () => {
    expect(iso(0, true)).toBe('1904-01-01');
  });

  it('maps serial 1 to 1904-01-02 (no leap-year bug in the 1904 system)', () => {
    expect(iso(1, true)).toBe('1904-01-02');
  });

  it('is offset exactly 1462 days from the 1900 system for the same calendar date', () => {
    // 2024-01-01 is serial 45292 in the 1900 system and 43830 in the 1904
    // system. 45292 − 43830 = 1462 (4 years + 1 day = the classic Mac shift).
    expect(iso(43830, true)).toBe('2024-01-01');
    expect(45292 - 43830).toBe(1462);
  });

  it('preserves the fractional time-of-day component', () => {
    const d = excelSerialToUtcDate(0.5, true);
    expect(d.getUTCFullYear()).toBe(1904);
    expect(d.getUTCMonth() + 1).toBe(1);
    expect(d.getUTCDate()).toBe(1);
    expect(d.getUTCHours()).toBe(12);
  });
});

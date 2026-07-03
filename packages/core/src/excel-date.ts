// Shared Excel serial-date → JS `Date` conversion (ECMA-376 §18.17.4.1).
//
// A "serial date-time" is a signed number of days relative to a base date. Two
// base dates exist, selected by the `<workbookPr date1904>` attribute
// (§18.2.28) for cells and by `<c:date1904>` (§21.2.2.38) for charts:
//
//   • 1900 date system (default): base date is 1899-12-30 (serial 0), so
//     serial 1 → 1900-01-01. Excel additionally reproduces the Lotus 1-2-3
//     leap-year bug, treating the non-existent 1900-02-29 as serial 60. That
//     phantom day shifts every serial ≥ 60 forward by one relative to the
//     proleptic Gregorian calendar, so to recover the correct calendar date we
//     add one day back for serials < 60 (which sit before the phantom day and
//     are therefore off by one under a bug-free epoch). Serial 60 itself has no
//     real Gregorian date; we leave it unshifted (→ 1900-02-28), matching the
//     long-standing core behaviour rather than inventing a mapping.
//
//   • 1904 date system: base date is 1904-01-01 (serial 0), so serial 1 →
//     1904-01-02. This system has no leap-year bug. The 1904 base is exactly
//     1462 days after the 1900 base date, which is why a Mac-authored (1904)
//     workbook shows dates 1462 days off when read with a 1900 epoch.
//
// Times are carried by the fractional part of the serial (§18.17.4.2): 0.5 =
// noon. All arithmetic is done in UTC so the caller reads `getUTC*` accessors
// without local-timezone drift.

const MS_PER_DAY = 86_400_000;

// Base dates as UTC epoch milliseconds.
//   1900: 1899-12-30 (serial 0)
//   1904: 1904-01-01 (serial 0)
const BASE_1900_MS = Date.UTC(1899, 11, 30);
const BASE_1904_MS = Date.UTC(1904, 0, 1);

/**
 * Convert an Excel date-time serial to a UTC `Date`.
 *
 * @param serial   Signed serial date-time (integer part = day, fraction = time).
 * @param date1904 `true` for the 1904 date system (Mac-authored workbooks /
 *                 `<c:date1904 val="1">`); `false`/omitted for the default 1900
 *                 date system.
 */
export function excelSerialToUtcDate(serial: number, date1904 = false): Date {
  if (date1904) {
    return new Date(BASE_1904_MS + serial * MS_PER_DAY);
  }
  // 1900 system: apply the Lotus leap-year-bug compensation. Serials < 60 sit
  // before the phantom 1900-02-29 and are one day short under the bug-free
  // 1899-12-30 epoch, so add a day back. Serials ≥ 60 need no adjustment.
  const adjusted = serial < 60 ? serial + 1 : serial;
  return new Date(BASE_1900_MS + adjusted * MS_PER_DAY);
}

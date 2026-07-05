// Office-style decimal rounding for DISPLAY formatting.
//
// Excel, Word and PowerPoint round numbers on their DECIMAL representation with
// half-up (round half away from zero) semantics: `#,##0.00` on 2.675 shows
// "2.68". JavaScript's `Number.prototype.toFixed` rounds the underlying BINARY
// double instead, and the nearest double to 2.675 is 2.67499999999999982…, so
// `(2.675).toFixed(2) === "2.67"`. Likewise 1.005 → "1.00", 8.575 → "8.57".
//
// `roundDecimalHalfUp` reproduces the spreadsheet behaviour by operating on the
// shortest round-tripping decimal string `String(value)` (the decimal the user
// authored), then rounding that string half-up with integer carry — no scaled
// multiply, so it never reintroduces a binary artefact. On values that are NOT
// on a `.xx5` boundary it agrees with `toFixed`, so adopting it is a no-op for
// real corpus data and only fixes the boundary cases.

/** Round `value` to `digits` fractional places using Office's decimal
 *  round-half-up (half away from zero) and return the fixed-precision string
 *  (always exactly `digits` fractional digits; no thousands separators).
 *
 *  - `2.675, 2` → `"2.68"`, `1.005, 2` → `"1.01"`, `9.995, 2` → `"10.00"`.
 *  - Negatives round away from zero by magnitude: `-2.675, 2` → `"-2.68"`.
 *  - A value that rounds to zero never carries a `-` sign (`-0.0001, 2` → `"0.00"`).
 *  - Non-finite inputs fall back to `String(value)` (`NaN`/`Infinity`).
 *
 *  `digits` must be a non-negative integer. */
export function roundDecimalHalfUp(value: number, digits: number): string {
  if (!Number.isFinite(value)) return String(value);
  const d = Math.max(0, Math.trunc(digits));

  const negative = value < 0;
  // Shortest decimal that round-trips to `value` — the decimal the author meant
  // (`String(2.675) === "2.675"`), free of the binary artefact `toFixed` sees.
  const decimal = expandExponential(Math.abs(value).toString());
  const [intPart, fracPartRaw = ''] = decimal.split('.');

  // Digit array for the integer part plus the fraction padded to at least d+1
  // places, so index `d` is the first dropped digit (the rounding decider).
  const fracPart = fracPartRaw.padEnd(d + 1, '0');
  const keptFrac = fracPart.slice(0, d);
  const decider = fracPart.charCodeAt(d) - 48; // digit at position d (0..9)

  let digitsArr = (intPart + keptFrac).split('').map((c) => c.charCodeAt(0) - 48);
  if (decider >= 5) {
    // Propagate the carry from the least-significant kept digit leftwards.
    let i = digitsArr.length - 1;
    for (; i >= 0; i--) {
      if (digitsArr[i] === 9) {
        digitsArr[i] = 0;
      } else {
        digitsArr[i] += 1;
        break;
      }
    }
    if (i < 0) digitsArr.unshift(1); // carried past the most-significant digit
  }

  // Split back into integer + fraction. The fraction is always the last d
  // digits (the integer part grows, but never shrinks, under carry).
  const combined = digitsArr.map((n) => String(n)).join('');
  const fracLen = d;
  const intStr = (fracLen > 0 ? combined.slice(0, combined.length - fracLen) : combined) || '0';
  const fracStr = fracLen > 0 ? combined.slice(combined.length - fracLen) : '';

  const intClean = intStr.replace(/^0+(?=\d)/, ''); // drop leading zeros, keep one
  const body = fracStr.length > 0 ? `${intClean}.${fracStr}` : intClean;

  // A value that rounds to all-zero must not keep the sign (avoid "-0.00").
  const isZero = /^[0.]*$/.test(body) && !/[1-9]/.test(body);
  return negative && !isZero ? `-${body}` : body;
}

/** Expand a JS numeric string in exponential form (`"1.2e-7"`, `"5e+21"`) into
 *  plain decimal digits so the string rounder can index its fraction. Non-
 *  exponential input is returned unchanged. Operates on the magnitude only
 *  (callers strip the sign first). */
function expandExponential(s: string): string {
  const m = /^(\d+)(?:\.(\d+))?[eE]([+-]?\d+)$/.exec(s);
  if (!m) return s;
  const [, intDigits, fracDigits = '', expStr] = m;
  const exp = parseInt(expStr, 10);
  const mantissa = intDigits + fracDigits;
  // Position of the decimal point measured from the start of `mantissa`.
  const pointPos = intDigits.length + exp;
  if (pointPos <= 0) {
    return '0.' + '0'.repeat(-pointPos) + mantissa;
  }
  if (pointPos >= mantissa.length) {
    return mantissa + '0'.repeat(pointPos - mantissa.length);
  }
  return mantissa.slice(0, pointPos) + '.' + mantissa.slice(pointPos);
}

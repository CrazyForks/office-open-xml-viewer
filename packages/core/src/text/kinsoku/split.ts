import type { KinsokuRules } from './rules.js';

/**
 * Adjust a CJK line-break position so it does not violate kinsoku.
 *
 * Given a line being split into `head = chars[0..splitAt)` (stays on the
 * current line) and `tail = chars[splitAt..]` (overflows to the next),
 * return the largest legal `splitAt' <= splitAt` such that:
 *   1. `tail'[0]` is not line-START-forbidden (行頭禁則 追い出し — the
 *      offending char and any preceding forbidden chars are pulled down
 *      onto the next line), AND
 *   2. `head'[last]` is not line-END-forbidden (push a dangling opener
 *      to the next line).
 *
 * Retraction is bounded: we never retract below `minSplit` (default 1,
 * so at least one code point always stays on a non-empty line and we
 * keep forward progress). If no legal split exists within that bound
 * (pathological run of forbidden chars), the original `splitAt` is
 * returned unchanged — Word likewise lets an over-long forbidden run
 * overflow rather than loop forever.
 *
 * `chars` must be an array of single code points (e.g. `[...text]`).
 */
export function kinsokuAdjustedSplit(
  chars: string[],
  splitAt: number,
  rules: KinsokuRules,
  minSplit = 1,
): number {
  if (!rules.enabled) return splitAt;
  if (splitAt <= 0 || splitAt >= chars.length) return splitAt;

  const startForbidden = (i: number): boolean =>
    i < chars.length && rules.lineStartForbidden.has(chars[i].codePointAt(0)!);
  const endForbidden = (i: number): boolean =>
    i >= 0 && rules.lineEndForbidden.has(chars[i].codePointAt(0)!);

  let s = splitAt;
  // Retract while the tail begins with a start-forbidden char OR the head
  // ends with an end-forbidden char. Each retraction moves one code point
  // from the head onto the tail (追い出し). Bounded by minSplit.
  while (s > minSplit && (startForbidden(s) || endForbidden(s - 1))) {
    s--;
  }
  // If we hit the floor and it is still illegal, no legal break exists in
  // range — fall back to the unrestricted split (never empty, never hang).
  if (s <= minSplit && (startForbidden(s) || endForbidden(s - 1))) {
    return splitAt;
  }
  return s;
}

/**
 * Cross-run 行頭禁則 (追い出し) helper. `kinsokuAdjustedSplit` only retracts
 * WITHIN the run being wrapped; when a line-start-forbidden char is the FIRST
 * code point of its run, the preceding character lives in an earlier segment on
 * the current line, so the offending char would be orphaned at the next line's
 * start. Given the current line's last text segment's code points (`lastChars`)
 * and that the overflowing run begins with a 行頭禁則 char, decide how many
 * trailing graphemes of `lastChars` to pull down so they lead the next line
 * ahead of that run.
 *
 * Returns the count k (≥1) to retract, or 0 to leave the line as-is (fall back).
 * Mirrors `kinsokuAdjustedSplit`'s bounded loop and re-validation: the pulled-down
 * graphemes must (1) start with a non-whitespace char that is itself NOT
 * line-start-forbidden (else the violation just moves), and (2) leave the
 * current line ending on a char that is NOT line-end-forbidden. `minKeep` is the
 * fewest code points that must remain on the current line (0 when other segments
 * precede this one, 1 when it is the only segment — never empty the line).
 *
 * ECMA-376 §17.3.1.16 (kinsoku toggle + default forbidden sets); §17.15.1.58/.59
 * (`noLineBreaksAfter` / `noLineBreaksBefore` custom sets).
 */
export function crossRunKinsokuRetract(
  lastChars: string[],
  rules: KinsokuRules,
  minKeep: number,
): number {
  if (!rules.enabled) return 0;
  const maxRetract = lastChars.length - minKeep;
  for (let k = 1; k <= maxRetract; k++) {
    const lead = lastChars[lastChars.length - k]; // becomes the next line's start
    if (/\s/.test(lead)) continue; // never orphan a whitespace at the line start
    if (rules.lineStartForbidden.has(lead.codePointAt(0)!)) continue; // still illegal — pull more
    const end = lastChars[lastChars.length - k - 1]; // new current-line end (may be undefined)
    if (end && rules.lineEndForbidden.has(end.codePointAt(0)!)) continue; // would dangle an opener
    return k;
  }
  return 0;
}

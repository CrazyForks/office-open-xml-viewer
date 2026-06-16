import { kinsokuAdjustedSplit, type KinsokuRules } from '@silurus/ooxml-core';

/** A measured grapheme: the character and its advance width in CSS px. */
export interface MeasuredChar {
  ch: string;
  w: number;
}

/**
 * Greedy line-break planner for a run of CJK characters, applying kinsoku
 * (ECMA-376 §17.15.1.58–.60) so a line never begins with a 行頭禁則 char
 * (、。」…) nor ends with a 行末禁則 char (「（…).
 *
 * The current line already holds `startWidth` px. Returns how many leading
 * chars of `chars` stay on it after kinsoku retraction; the caller emits those,
 * breaks, and calls again with the remainder and `startWidth = 0`.
 *
 * Progress: on an empty line (startWidth === 0) at least one char is always
 * placed, even if it overflows `maxWidth`. On a non-empty line the result may
 * be 0 — push the whole run to a fresh line (Word's 追い出し) — telling the
 * caller to break first.
 *
 * Note: this retracts WITHIN the run. Cross-run 追い出し (the run's first char
 * is 行頭禁則 and the preceding char lives in an earlier segment) is handled by
 * core's `crossRunKinsokuRetract` and is not wired here yet.
 */
export function fitCjkLine(
  chars: readonly MeasuredChar[],
  startWidth: number,
  maxWidth: number,
  rules: KinsokuRules,
): number {
  if (chars.length === 0) return 0;
  const lineEmpty = startWidth === 0;

  // Greedy fit by width.
  let raw = 0;
  let w = startWidth;
  for (const c of chars) {
    if (w + c.w > maxWidth) {
      if (raw > 0) break; // already have content on this line
      if (!lineEmpty) break; // non-empty line, nothing fits → caller breaks
      w += c.w; // empty line: force the first char so wrapping advances
      raw++;
      break;
    }
    w += c.w;
    raw++;
  }

  if (raw === 0) return 0; // non-empty line, nothing fits
  if (raw >= chars.length) return chars.length; // whole run fits

  const cps = chars.map((c) => c.ch);
  const minSplit = lineEmpty ? 1 : 0;
  return kinsokuAdjustedSplit(cps, raw, rules, minSplit);
}

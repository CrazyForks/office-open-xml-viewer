import type { TableBorderInput } from './types.js';
import { wordAuthoredBorderParticipates } from './table-compatibility.js';

/** Resolve one authored border layer before shared-edge conflict resolution.
 * `none` is an omitted layer while `nil` blocks lower-precedence fallback on
 * this side. Keeping that distinction here prevents callers that merge
 * tblPrEx/table inputs from replacing an authored nil with a lower layer. */
export function firstAuthoredTableBorder(
  ...borders: readonly (TableBorderInput | null)[]
): TableBorderInput | null {
  for (const border of borders) {
    // Compatibility-owned cascade distinction: nil participates while none falls through.
    if (border && wordAuthoredBorderParticipates(border.authoredStyle)) return border;
  }
  return null;
}

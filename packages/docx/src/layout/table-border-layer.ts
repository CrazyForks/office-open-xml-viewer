import type { TableBorderInput } from './types.js';
import { wordAuthoredBorderParticipates } from './table-compatibility.js';

/** Resolve one authored border layer before shared-edge conflict resolution.
 * `none` is an omitted layer while `nil` is an authored suppression. Keeping
 * that distinction here prevents callers that merge tblPrEx/table inputs from
 * accidentally turning an OOXML suppression into a fallback. */
export function firstAuthoredTableBorder(
  ...borders: readonly (TableBorderInput | null)[]
): TableBorderInput | null {
  for (const border of borders) {
    // Compatibility-owned distinction: nil suppresses while none falls through.
    if (border && wordAuthoredBorderParticipates(border.authoredStyle)) return border;
  }
  return null;
}

import type { LineLayout, ParagraphLayout, TableLayout } from './types.js';

export function footnoteIdsInRetainedLines(lines: readonly LineLayout[]): readonly string[] {
  return Object.freeze([...new Set(lines.flatMap((line) => line.placements.flatMap((placement) => (
    placement.kind === 'text' && placement.noteReference?.kind === 'footnote'
      ? [placement.noteReference.id]
      : []
  ))))]);
}

function tableFootnoteIds(table: TableLayout): readonly string[] {
  return table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks.flatMap(
    (block) => footnoteIdsInRetainedSlice(block.layout),
  )));
}

/** §17.11.21 / §17.18.34 assign the note to the physical page that paints its
 * reference; after splitting, parser-level paragraph/table identity is too broad. */
export function footnoteIdsInRetainedSlice(
  slice: ParagraphLayout | TableLayout,
): readonly string[] {
  const ids = slice.kind === 'paragraph'
    ? footnoteIdsInRetainedLines(slice.lines)
    : tableFootnoteIds(slice);
  return Object.freeze([...new Set(ids)]);
}

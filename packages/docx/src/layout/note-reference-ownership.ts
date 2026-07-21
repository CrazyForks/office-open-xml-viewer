import type { BodyElement, CellElement, DocNote } from '../types.js';
import type { LineLayout, ParagraphLayout, TableLayout } from './types.js';

/** Build default sequential numbering from first-reference order. Unreferenced
 * note-part entries do not consume a displayed number (§17.18.22/.34). */
export function buildNoteNumberMap(
  notes: readonly DocNote[] | undefined,
  referenceIds: readonly string[],
): Map<string, number> {
  const numbers = new Map<string, number>();
  if (!notes) return numbers;
  const available = new Set(notes.map((note) => note.id));
  referenceIds.forEach((id) => {
    if (available.has(id) && !numbers.has(id)) numbers.set(id, numbers.size + 1);
  });
  return numbers;
}

/** Index note-part entries by their OOXML id for story acquisition. */
export function indexNotes(notes: readonly DocNote[] | undefined): Map<string, DocNote> {
  const indexed = new Map<string, DocNote>();
  if (!notes) return indexed;
  for (const note of notes) indexed.set(note.id, note);
  return indexed;
}

/** Collect first reference order from the main document story. */
export function noteReferenceIdsInDocumentOrder(
  elements: readonly (BodyElement | CellElement)[],
  kind: 'footnote' | 'endnote',
): readonly string[] {
  const ids: string[] = [];
  const seen = new Set<string>();
  for (const element of elements) {
    if (element.type === 'paragraph') {
      for (const run of element.runs) {
        if (run.type !== 'text'
          || run.noteRef?.kind !== kind
          || run.noteRef.id.length === 0
          || seen.has(run.noteRef.id)) continue;
        seen.add(run.noteRef.id);
        ids.push(run.noteRef.id);
      }
    } else if (element.type === 'table') {
      for (const row of element.rows) for (const cell of row.cells) {
        for (const id of noteReferenceIdsInDocumentOrder(cell.content, kind)) {
          if (seen.has(id)) continue;
          seen.add(id);
          ids.push(id);
        }
      }
    }
  }
  return Object.freeze(ids);
}

function noteIdsInRetainedLines(
  lines: readonly LineLayout[],
  kind: 'footnote' | 'endnote',
): readonly string[] {
  return Object.freeze([...new Set(lines.flatMap((line) => line.placements.flatMap((placement) => (
    placement.kind === 'text' && placement.noteReference?.kind === kind
      ? [placement.noteReference.id]
      : []
  ))))]);
}

function tableNoteIds(
  table: TableLayout,
  kind: 'footnote' | 'endnote',
): readonly string[] {
  return table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks.flatMap(
    (block) => noteIdsInRetainedSlice(block.layout, kind),
  )));
}

function noteIdsInRetainedSlice(
  slice: ParagraphLayout | TableLayout,
  kind: 'footnote' | 'endnote',
): readonly string[] {
  const ids = slice.kind === 'paragraph'
    ? noteIdsInRetainedLines(slice.lines, kind)
    : tableNoteIds(slice, kind);
  return Object.freeze([...new Set(ids)]);
}

export function footnoteIdsInRetainedLines(lines: readonly LineLayout[]): readonly string[] {
  return noteIdsInRetainedLines(lines, 'footnote');
}

/** §17.11.21 / §17.18.34 assign the note to the physical page that paints its
 * reference; after splitting, parser-level paragraph/table identity is too broad. */
export function footnoteIdsInRetainedSlice(
  slice: ParagraphLayout | TableLayout,
): readonly string[] {
  return noteIdsInRetainedSlice(slice, 'footnote');
}

/** §17.18.22: unreferenced endnotes are not displayed. */
export function endnoteIdsInRetainedSlice(
  slice: ParagraphLayout | TableLayout,
): readonly string[] {
  return noteIdsInRetainedSlice(slice, 'endnote');
}

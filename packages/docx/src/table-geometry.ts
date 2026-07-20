// Table row-height resolution skeleton (ECMA-376 §17.4.80 `<w:trHeight>` /
// §17.18.37 ST_HeightRule + §17.4.85 `<w:vMerge>` span extension).
//
// The trHeight rule (exact / atLeast / auto), the auto/no-floor minimum, the
// gridSpan column slicing, and the vMerge restart-span post-pass are pure
// structural math that does NOT depend on whether the caller works in pt or
// scaled device units. The caller supplies cell-content measurement, keeping
// this compatibility kernel independent of paragraph acquisition details.
//
// Only DocTable/DocTableRow/DocTableCell types are imported (erased at runtime),
// so there is no import cycle with renderer.ts.

import type { DocTable, DocTableRow, DocTableCell } from './types.js';
import { resolveBorderConflict, resolveCellEdges } from './cell-border-conflict.js';
import { wordAuthoredAutoRowHeightUsesFloor } from './layout/table-compatibility.js';

/** Minimum table-row height (pt) when no `w:trHeight` floor applies — i.e. an
 *  `auto` row, or `atLeast`/`exact` with no `@val`. ECMA-376 leaves the auto
 *  minimum implementation-defined; this is the floor an empty row collapses to
 *  before content (cell margins + measured content) expands it. */
export const MIN_ROW_HEIGHT_PT = 10;

/** ECMA-376 §17.4.15 — clamp the number of shared table-grid columns skipped
 * before a row's first cell. Older parsed models omit the property, so zero is
 * the compatibility default required by the specification. */
export function rowGridBefore(row: DocTableRow, columnCount: number): number {
  const raw = row.gridBefore ?? 0;
  if (!Number.isFinite(raw)) return 0;
  const value = Math.max(0, Math.trunc(raw));
  const size = Math.max(0, Math.trunc(columnCount));
  // §17.4.15: a value larger than tblGrid is ignored; it is not clamped to
  // the trailing edge (which would collapse every real cell to zero width).
  return value > size ? 0 : value;
}

/** Last row index covered by the vMerge span that starts at (`startRi`,
 *  `startCi`). A span continues through following rows whose cell anchored in the
 *  same grid column carries `vMerge=continue` (ECMA-376 §17.4.85). Pure
 *  table-structure walk — no geometry. */
export function findMergeEndRow(table: DocTable, startRi: number, startCi: number): number {
  let endRi = startRi;
  for (let rj = startRi + 1; rj < table.rows.length; rj++) {
    const row = table.rows[rj];
    let ci = rowGridBefore(row, table.colWidths.length);
    let matched = false;
    for (const cell of row.cells) {
      if (ci === startCi) {
        if (cell.vMerge === false) matched = true;
        break;
      }
      if (ci > startCi) break;
      ci += cell.colSpan;
    }
    if (!matched) break;
    endRi = rj;
  }
  return endRi;
}

/** Measure the content height (already in the target units — scaled device
 *  units at `scale`, or pt at `scale=1`) of a single cell laid out at the given
 *  total cell width (`cellWidth`, the summed widths of the grid columns it
 *  spans). MUST include the cell's top/bottom margins. */
export type MeasureCellContentHeight = (cell: DocTableCell, cellWidth: number) => number;

interface TableGridOwner {
  cell: DocTableCell;
  ri: number;
  lastRi: number;
  ci: number;
  span: number;
}

function cellAtGridColumn(
  row: DocTableRow,
  targetCi: number,
  columnCount: number,
): DocTableCell | null {
  let ci = rowGridBefore(row, columnCount);
  for (const cell of row.cells) {
    if (targetCi >= ci && targetCi < ci + cell.colSpan) return cell;
    ci += cell.colSpan;
  }
  return null;
}

function paintWidth(candidate: ReturnType<typeof resolveBorderConflict>): number {
  const spec = candidate?.spec;
  if (!spec || spec.style === 'none' || spec.style === 'nil') return 0;
  return spec.width;
}

/** Width of each resolved horizontal table-grid boundary, from the outer top
 * through every shared row boundary to the outer bottom. Adjacent-cell
 * conflicts use the same §17.4.66 cascade as paint; a boundary inside one
 * vertically merged owner contributes no rule. */
export function resolvedHorizontalBoundaryWidths(table: DocTable): number[] {
  const rowCount = table.rows.length;
  const columnCount = table.colWidths.length;
  const widths = new Array<number>(rowCount + 1).fill(0);
  if (rowCount === 0 || columnCount === 0) return widths;

  const owners: Array<Array<TableGridOwner | null>> = Array.from(
    { length: rowCount },
    () => new Array<TableGridOwner | null>(columnCount).fill(null),
  );
  for (let ri = 0; ri < rowCount; ri++) {
    const row = table.rows[ri];
    let ci = rowGridBefore(row, columnCount);
    for (const cell of row.cells) {
      const span = Math.min(cell.colSpan, columnCount - ci);
      if (cell.vMerge !== false && span > 0) {
        const lastRi = cell.vMerge === true ? findMergeEndRow(table, ri, ci) : ri;
        const owner: TableGridOwner = { cell, ri, lastRi, ci, span };
        for (let rj = ri; rj <= lastRi; rj++) {
          for (let cj = ci; cj < ci + span; cj++) owners[rj][cj] = owner;
        }
      }
      ci += span;
    }
  }

  const edgesFor = (owner: TableGridOwner) => ({
    topRow: owner.ri === 0,
    bottomRow: owner.lastRi === rowCount - 1,
    leftCol: owner.ci === 0,
    rightCol: owner.ci + owner.span === columnCount,
  });
  const bottomCandidate = (owner: TableGridOwner) => {
    const terminal = owner.lastRi > owner.ri
      ? (cellAtGridColumn(table.rows[owner.lastRi], owner.ci, columnCount) ?? owner.cell)
      : owner.cell;
    return resolveCellEdges(terminal.borders, table.borders, edgesFor(owner), false).bottom;
  };

  for (let ci = 0; ci < columnCount; ci++) {
    const topOwner = owners[0][ci];
    if (topOwner) {
      const top = resolveCellEdges(
        topOwner.cell.borders,
        table.borders,
        edgesFor(topOwner),
        false,
      ).top;
      widths[0] = Math.max(widths[0], paintWidth(resolveBorderConflict(top, null)));
    }

    for (let boundary = 1; boundary < rowCount; boundary++) {
      // Runtime row pieces separated inside one emitted slice are one
      // continuous source row. Paint deliberately leaves that internal cut
      // open, so it has no footprint in either adjacent row box.
      if ((table.rows[boundary - 1] as DocTableRow & { pageCutBottom?: boolean })
        .pageCutBottom === true) continue;
      const above = owners[boundary - 1][ci];
      const below = owners[boundary][ci];
      if (above && above === below) continue;
      const aboveBottom = above ? bottomCandidate(above) : null;
      const belowTop = below
        ? resolveCellEdges(below.cell.borders, table.borders, edgesFor(below), false).top
        : null;
      widths[boundary] = Math.max(
        widths[boundary],
        paintWidth(resolveBorderConflict(aboveBottom, belowTop)),
      );
    }

    const bottomOwner = owners[rowCount - 1][ci];
    if (bottomOwner) {
      widths[rowCount] = Math.max(
        widths[rowCount],
        paintWidth(resolveBorderConflict(bottomCandidate(bottomOwner), null)),
      );
    }
  }
  return widths;
}

/**
 * Resolve per-row heights for a table whose grid columns have widths
 * `colWidths` (in the same target units as the measurer returns: px when
 * `scale` is the device scale, pt when `scale === 1`). Applies, once:
 *
 *   - ECMA-376 §17.4.80 / §17.18.37 (ST_HeightRule):
 *       exact   — height is exactly `w:trHeight/@val` (× `scale`); overflow is
 *                 clipped by the caller.
 *       atLeast — `@val` (× `scale`) is a lower bound; content can grow the row.
 *       auto    — by the literal text of §17.4.80, `@val` is IGNORED ("no
 *                 predetermined minimum or maximum size", advisory layout cache
 *                 only). `word-authored-auto-row-height-floor` owns the
 *                 compatibility deviation that treats an authored auto value as
 *                 a lower bound. With `@val` absent, auto falls back to
 *                 `MIN_ROW_HEIGHT_PT`.
 *   - gridSpan: a cell's width is the sum of the `cell.colSpan` columns it
 *     anchors (clamped to the remaining columns).
 *   - ECMA-376 §17.4.85 (vMerge): a `vMerge=restart` cell's content occupies the
 *     whole merged span. It is EXCLUDED from its first row's height (so the first
 *     row is not inflated) and instead a post-pass extends the span's LAST row
 *     when the restart cell's content exceeds the summed span height.
 *     `vMerge=continue` cells render no content.
 *
 * `measureCellContentHeight` supplies the unit-specific content measurement (see
 * its type doc). The restart cell is measured through the SAME callback in the
 * post-pass; the callback must be a pure read of layout state (it is, in both
 * callers) so re-measuring yields the value the first pass would have computed.
 */
/**
 * Height of ONE row (ECMA-376 §17.4.80 / §17.18.37 ST_HeightRule + gridSpan),
 * EXCLUDING the §17.4.85 vMerge span extension (the caller / the table-level
 * resolver applies that in a post-pass). `exact` returns exactly `@val × scale`;
 * `atLeast` floors at `@val × scale`;
 * `word-authored-auto-row-height-floor` gives an `auto` row with `@val` the
 * same floor despite the §17.4.80 literal "ignored". `auto` with no `@val` floors
 * at `MIN_ROW_HEIGHT_PT × scale`. A `vMerge=restart` cell is excluded (its
 * content is distributed across the span, not absorbed by its first row) and a
 * `vMerge=continue` cell renders no content. This is the single source of the
 * trHeight rule, shared by {@link resolveTableRowHeights} and direct
 * compatibility tests of a single row.
 */
export function resolveSingleRowHeight(
  row: DocTableRow,
  colWidths: number[],
  scale: number,
  measureCellContentHeight: MeasureCellContentHeight,
): number {
  if (row.rowHeight != null && row.rowHeightRule === 'exact') return row.rowHeight * scale;
  // `word-authored-auto-row-height-floor` owns the legacy-model deviation from
  // §17.4.80's literal auto behavior. With `@val` absent, auto still collapses
  // to the implementation-defined minimum.
  let rowH =
    row.rowHeight != null && (
      row.rowHeightRule === 'atLeast'
      || wordAuthoredAutoRowHeightUsesFloor(row.rowHeightRule, row.rowHeight)
    )
      ? row.rowHeight * scale
      : MIN_ROW_HEIGHT_PT * scale;

  let ci = rowGridBefore(row, colWidths.length);
  for (const cell of row.cells) {
    const span = Math.min(cell.colSpan, colWidths.length - ci);
    // vMerge=restart cells are sized by the span post-pass; vMerge=continue
    // cells render no content. Neither raises THIS row's height directly.
    if (cell.vMerge !== true && cell.vMerge !== false) {
      const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);
      const ch = measureCellContentHeight(cell, cellW);
      if (ch > rowH) rowH = ch;
    }
    ci += span;
  }
  return rowH;
}

/** Resolve row content boxes, before page-local horizontal rule footprints are
 * added. Keeping this geometry separate is essential for pagination: once a
 * table is sliced, the first/last rows resolve against that slice's outer
 * borders rather than the original table's interior boundaries. */
export function resolveTableRowContentHeights(
  table: DocTable,
  colWidths: number[],
  scale: number,
  measureCellContentHeight: MeasureCellContentHeight,
): number[] {
  const rowHeights = table.rows.map((row) =>
    resolveSingleRowHeight(row, colWidths, scale, measureCellContentHeight),
  );

  // §17.4.85 span extension: for each vMerge=restart cell, grow the span's last
  // row if the restart cell's full content is taller than the summed span rows.
  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    let ci = rowGridBefore(row, colWidths.length);
    for (const cell of row.cells) {
      const span = Math.min(cell.colSpan, colWidths.length - ci);
      if (cell.vMerge === true) {
        const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);
        const contentH = measureCellContentHeight(cell, cellW);
        const endRi = findMergeEndRow(table, ri, ci);
        let spanH = 0;
        for (let rj = ri; rj <= endRi; rj++) spanH += rowHeights[rj];
        if (spanH < contentH) {
          rowHeights[endRi] += contentH - spanH;
        }
      }
      ci += span;
    }
  }

  return rowHeights;
}

/** Add the horizontal rule footprint painted by this concrete table or page
 * slice. Auto/atLeast row boxes run between rule centres; exact trHeight already
 * defines the complete row box and is not expanded (§17.4.80). */
export function applyTableRowBoundaryFootprints(
  table: DocTable,
  contentHeights: readonly number[],
  scale: number,
): number[] {
  const horizontalBoundaries = resolvedHorizontalBoundaryWidths(table);
  return contentHeights.map((contentHeight, ri) => {
    if (table.rows[ri]?.rowHeightRule === 'exact') return contentHeight;
    return contentHeight + (
      (horizontalBoundaries[ri] ?? 0) + (horizontalBoundaries[ri + 1] ?? 0)
    ) * scale / 2;
  });
}

export function resolveTableRowHeights(
  table: DocTable,
  colWidths: number[],
  scale: number,
  measureCellContentHeight: MeasureCellContentHeight,
): number[] {
  return applyTableRowBoundaryFootprints(
    table,
    resolveTableRowContentHeights(table, colWidths, scale, measureCellContentHeight),
    scale,
  );
}

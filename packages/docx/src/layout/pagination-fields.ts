import { convergeLayout, type LayoutIteration } from './convergence.js';
import { stableFingerprint } from './fingerprint.js';
import type { FlowFragment } from '../layout-fragments.js';

// A resource-safety bound, not visual tuning. Sixteen exceeds the decimal digit
// transitions of any practical page count; convergeLayout's seen-set catches
// cycles earlier, while exhausting this limit hard-fails instead of returning
// stale field geometry.
const PAGINATION_FIELD_CONVERGENCE_LIMIT = 16;

/** Resolve PAGE/NUMPAGES-dependent pagination to a stable physical geometry. */
export function convergePaginationFields<T extends LayoutIteration>(
  acquire: (totalPagesHint: number) => T,
  limit = PAGINATION_FIELD_CONVERGENCE_LIMIT,
): T {
  const seed = acquire(1);
  return convergeLayout(seed, (current) => acquire(current.pageCount), limit);
}

/** Project retained flow into convergence-relevant plain data. Field values are
 * retained because equal-width PAGE values can still belong to different pages. */
export function paginationFieldFlowGeometry(fragment: FlowFragment): unknown {
  if (fragment.kind === 'paragraph') {
    return {
      kind: fragment.kind,
      flowBounds: fragment.flowBounds,
      inkBounds: fragment.inkBounds,
      clipBounds: fragment.clipBounds,
      advancePt: fragment.advancePt,
      spacing: fragment.spacing,
      lines: fragment.lines.map((line) => ({
        range: line.range,
        bounds: line.bounds,
        baselinePt: line.baselinePt,
        advancePt: line.advancePt,
        placements: line.placements.map((placement) => ({
          kind: placement.kind,
          range: placement.range,
          bounds: placement.bounds,
          ...('advancePt' in placement ? { advancePt: placement.advancePt } : {}),
          ...(placement.kind === 'text' && placement.dependency
            ? {
                field: {
                  dependency: placement.dependency,
                  text: placement.text,
                  sourceRunIndex: placement.sourceRunIndex,
                },
              }
            : {}),
        })),
      })),
      drawings: fragment.drawings.map((drawing) => ({
        flowBounds: drawing.flowBounds,
        inkBounds: drawing.inkBounds,
        transform: drawing.transform,
        clip: drawing.clip,
      })),
      textBoxes: fragment.textBoxes.map((textBox) => ({
        flowBounds: textBox.flowBounds,
        inkBounds: textBox.inkBounds,
        advancePt: textBox.advancePt,
      })),
      exclusions: fragment.exclusions.map((exclusion) => ({
        wrap: exclusion.wrap,
        bounds: exclusion.bounds,
        polygon: exclusion.polygon,
      })),
    };
  }
  return {
    kind: fragment.kind,
    columnWidthsPt: fragment.columnWidthsPt,
    continuesFromPreviousPage: fragment.continuesFromPreviousPage,
    continuesOnNextPage: fragment.continuesOnNextPage,
    rows: fragment.rows.map((row) => ({
      sourceRowIndex: row.sourceRowIndex,
      heightPt: row.heightPt,
      repeatedHeader: row.repeatedHeader,
      cells: row.cells.map((cell) => ({
        verticalMerge: cell.verticalMerge,
        boxHeightPt: cell.boxHeightPt,
        blocks: cell.blocks.map(paginationFieldFlowGeometry),
      })),
    })),
  };
}

function definedRuntimeGeometry(value: unknown): unknown {
  if (Array.isArray(value)) return value.map((entry) => definedRuntimeGeometry(entry));
  if (value !== null && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .filter(([, entry]) => entry !== undefined)
        .map(([key, entry]) => [key, definedRuntimeGeometry(entry)]),
    );
  }
  return value;
}

/**
 * Fingerprint pagination geometry after omitting optional runtime placement
 * facts whose absence may be represented by a missing key or `undefined`.
 * This normalization is local to field convergence; other layout contracts keep
 * rejecting undefined data through the ordinary fingerprint boundary.
 */
export function paginationFieldGeometryFingerprint(value: unknown): string {
  return stableFingerprint(
    'pagination-field-geometry',
    definedRuntimeGeometry(value),
  );
}

import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { HARD_MAX_XLSX_RENDERER_COORDINATE_INDEX_ENTRIES } from '@silurus/ooxml-core/worker';

/** Hard implementation ceiling for each coordinate-keyed renderer index. */
export const MAX_RENDERER_COORDINATE_INDEX_ENTRIES =
  HARD_MAX_XLSX_RENDERER_COORDINATE_INDEX_ENTRIES;

export interface CoordinateRange {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export interface CoordinateIndexIdentity {
  resource: string;
  operation: string;
}

function resourceLimitError(
  identity: CoordinateIndexIdentity,
  observed: number,
): OoxmlResourceLimitError {
  const limit = MAX_RENDERER_COORDINATE_INDEX_ENTRIES;
  return new OoxmlResourceLimitError(
    `XLSX renderer ${identity.resource} exceeded its hard limit of ${limit} entries`,
    {
      stage: 'rendering',
      violation: {
        format: 'xlsx',
        operation: identity.operation,
        resource: identity.resource,
        metric: 'entry-count',
        limit,
        observed,
        configurable: false,
        // Renderer-only failures do not have access to the package session's
        // inflation counters. Keep those independent measurements explicit
        // rather than fabricating package usage from worksheet model data.
        usage: {
          archiveEntryCount: 0,
          declaredInflatedBytes: 0,
          distinctInflatedBytes: 0,
          operationInflatedBytes: 0,
        },
      },
    },
  );
}

/**
 * Verify that one rectangular expansion cannot cross the hard ceiling before
 * entering its nested coordinate loops. `excludedEntries` covers topology
 * entries such as a merge anchor that are deliberately not inserted.
 */
export function assertCoordinateRangeArea(
  range: CoordinateRange,
  identity: CoordinateIndexIdentity,
  excludedEntries = 0,
): number {
  const { top, bottom, left, right } = range;
  if (bottom < top || right < left) return 0;

  const coordinates = [top, bottom, left, right];
  if (
    !coordinates.every(Number.isSafeInteger)
    || !Number.isSafeInteger(excludedEntries)
    || excludedEntries < 0
  ) {
    throw resourceLimitError(identity, MAX_RENDERER_COORDINATE_INDEX_ENTRIES + 1);
  }

  const height = bottom - top + 1;
  const width = right - left + 1;
  const maxArea = MAX_RENDERER_COORDINATE_INDEX_ENTRIES + excludedEntries;
  if (
    !Number.isSafeInteger(height)
    || !Number.isSafeInteger(width)
    || height > maxArea
    || width > Math.floor(maxArea / height)
  ) {
    throw resourceLimitError(identity, MAX_RENDERER_COORDINATE_INDEX_ENTRIES + 1);
  }

  const entries = height * width - excludedEntries;
  if (entries > MAX_RENDERER_COORDINATE_INDEX_ENTRIES) {
    throw resourceLimitError(identity, entries);
  }
  return entries;
}

/** Insert a coordinate while charging only an actual new Map entry. */
export function setCoordinateIndexValue<T>(
  map: Map<string, T>,
  key: string,
  value: T,
  identity: CoordinateIndexIdentity,
): void {
  if (!map.has(key) && map.size >= MAX_RENDERER_COORDINATE_INDEX_ENTRIES) {
    throw resourceLimitError(identity, map.size + 1);
  }
  map.set(key, value);
}

/** Build the renderer's canonical row:column cell lookup. */
export function buildCellCoordinateIndex<T extends { row: number; col: number }>(
  rows: readonly { cells: readonly T[] }[],
  identity: CoordinateIndexIdentity,
): Map<string, T> {
  const map = new Map<string, T>();
  for (const row of rows) {
    for (const cell of row.cells) {
      setCoordinateIndexValue(map, `${cell.row}:${cell.col}`, cell, identity);
    }
  }
  return map;
}

/** Insert a coordinate while charging only an actual new Set entry. */
export function addCoordinateIndexEntry(
  set: Set<string>,
  key: string,
  identity: CoordinateIndexIdentity,
): void {
  if (!set.has(key) && set.size >= MAX_RENDERER_COORDINATE_INDEX_ENTRIES) {
    throw resourceLimitError(identity, set.size + 1);
  }
  set.add(key);
}

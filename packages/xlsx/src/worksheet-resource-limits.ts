import {
  HARD_MAX_XLSX_WORKBOOK_CACHED_CELLS,
  HARD_MAX_XLSX_WORKBOOK_CACHED_ROWS,
  HARD_MAX_XLSX_WORKSHEET_CELLS,
  HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES,
  HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
  HARD_MAX_XLSX_WORKSHEET_ROWS,
} from '@silurus/ooxml-core/worker';
import {
  OoxmlResourceLimitError,
  type OoxmlResourceUsageSnapshot,
} from '@silurus/ooxml-core';
import type { Row, Worksheet } from './types.js';

export const XLSX_MAX_MATERIALIZED_ROWS = HARD_MAX_XLSX_WORKSHEET_ROWS;
export const XLSX_MAX_MATERIALIZED_CELLS = HARD_MAX_XLSX_WORKSHEET_CELLS;
export const XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES =
  HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES;
export const XLSX_MAX_MATERIALIZED_JSON_BYTES = HARD_MAX_XLSX_WORKSHEET_JSON_BYTES;
export const XLSX_MAX_CACHED_ROWS = HARD_MAX_XLSX_WORKBOOK_CACHED_ROWS;
export const XLSX_MAX_CACHED_CELLS = HARD_MAX_XLSX_WORKBOOK_CACHED_CELLS;

export interface WorksheetModelUsage {
  rows: number;
  cells: number;
  ownedUtf8Bytes: number;
}

export interface WorksheetCacheUsage {
  rows: number;
  cells: number;
}

const ZERO_RESOURCE_USAGE: OoxmlResourceUsageSnapshot = Object.freeze({
  archiveEntryCount: 0,
  declaredInflatedBytes: 0,
  distinctInflatedBytes: 0,
  operationInflatedBytes: 0,
});

function cappedAdd(left: number, right: number, limit: number): number {
  if (!Number.isSafeInteger(left) || !Number.isSafeInteger(right) || left < 0 || right < 0) {
    throw new Error('worksheet resource measurement must use non-negative safe integers');
  }
  return left > limit - right ? limit + 1 : left + right;
}

export function utf8Bytes(value: string, limit = Number.MAX_SAFE_INTEGER): number {
  let bytes = 0;
  for (let index = 0; index < value.length; index += 1) {
    const unit = value.charCodeAt(index);
    let width: number;
    if (unit <= 0x7f) width = 1;
    else if (unit <= 0x7ff) width = 2;
    else if (unit >= 0xd800 && unit <= 0xdbff && index + 1 < value.length) {
      const low = value.charCodeAt(index + 1);
      if (low >= 0xdc00 && low <= 0xdfff) {
        width = 4;
        index += 1;
      } else width = 3;
    } else width = 3;
    bytes = cappedAdd(bytes, width, limit);
    if (bytes > limit) return bytes;
  }
  return bytes;
}

function jsonStringBytes(value: string, limit: number): number {
  let bytes = 2;
  for (let index = 0; index < value.length; index += 1) {
    const unit = value.charCodeAt(index);
    let width: number;
    if (
      unit === 0x22 || unit === 0x5c || unit === 0x08 || unit === 0x09 ||
      unit === 0x0a || unit === 0x0c || unit === 0x0d
    ) width = 2;
    else if (unit <= 0x1f) width = 6;
    else if (unit <= 0x7f) width = 1;
    else if (unit <= 0x7ff) width = 2;
    else if (unit >= 0xd800 && unit <= 0xdbff && index + 1 < value.length) {
      const low = value.charCodeAt(index + 1);
      if (low >= 0xdc00 && low <= 0xdfff) {
        width = 4;
        index += 1;
      } else width = 6;
    } else if (unit >= 0xd800 && unit <= 0xdfff) width = 6;
    else width = 3;
    bytes = cappedAdd(bytes, width, limit);
    if (bytes > limit) return bytes;
  }
  return bytes;
}

interface ValueMeasurement {
  jsonBytes: number;
  ownedUtf8Bytes: number;
}

function measureValue(value: unknown, limit: number, inArray = false): ValueMeasurement {
  if (value === null) return { jsonBytes: 4, ownedUtf8Bytes: 0 };
  if (typeof value === 'string') {
    return {
      jsonBytes: jsonStringBytes(value, limit),
      ownedUtf8Bytes: utf8Bytes(value, limit),
    };
  }
  if (typeof value === 'boolean') return { jsonBytes: value ? 4 : 5, ownedUtf8Bytes: 0 };
  if (typeof value === 'number') {
    const serialized = Number.isFinite(value) ? String(Object.is(value, -0) ? 0 : value) : 'null';
    return { jsonBytes: serialized.length, ownedUtf8Bytes: 0 };
  }
  if (Array.isArray(value)) {
    let jsonBytes = 2;
    let ownedUtf8Bytes = 0;
    for (let index = 0; index < value.length; index += 1) {
      if (index !== 0) jsonBytes = cappedAdd(jsonBytes, 1, limit);
      const item = measureValue(value[index], limit, true);
      jsonBytes = cappedAdd(jsonBytes, item.jsonBytes, limit);
      ownedUtf8Bytes = cappedAdd(ownedUtf8Bytes, item.ownedUtf8Bytes, limit);
    }
    return { jsonBytes, ownedUtf8Bytes };
  }
  if (typeof value === 'object') {
    let jsonBytes = 2;
    let ownedUtf8Bytes = 0;
    let emitted = 0;
    for (const [key, entry] of Object.entries(value)) {
      if (entry === undefined || typeof entry === 'function' || typeof entry === 'symbol') continue;
      if (emitted++ !== 0) jsonBytes = cappedAdd(jsonBytes, 1, limit);
      jsonBytes = cappedAdd(jsonBytes, jsonStringBytes(key, limit), limit);
      jsonBytes = cappedAdd(jsonBytes, 1, limit);
      const child = measureValue(entry, limit);
      jsonBytes = cappedAdd(jsonBytes, child.jsonBytes, limit);
      ownedUtf8Bytes = cappedAdd(ownedUtf8Bytes, child.ownedUtf8Bytes, limit);
    }
    return { jsonBytes, ownedUtf8Bytes };
  }
  return { jsonBytes: inArray ? 4 : 0, ownedUtf8Bytes: 0 };
}

export function measureRows(rows: readonly Row[]): WorksheetModelUsage {
  const cells = rows.reduce(
    (total, row) => cappedAdd(total, row.cells.length, XLSX_MAX_MATERIALIZED_CELLS),
    0,
  );
  return {
    rows: rows.length,
    cells,
    // This is deliberately cell-scoped. It includes every retained string in
    // Cell.value (rich/phonetic formatting and the discriminator included) and
    // formula text. Shared values must be resolved before calling this helper,
    // so repeated shared-string references are charged once per materialized
    // cell. Ancillary worksheet strings are covered by the exact JSON ceiling.
    ownedUtf8Bytes: rows.reduce((rowTotal, row) => row.cells.reduce((cellTotal, cell) => {
      const valueBytes = measureValue(
        cell.value,
        XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES,
      ).ownedUtf8Bytes;
      const formulaBytes = cell.formula === undefined
        ? 0
        : utf8Bytes(cell.formula, XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES);
      return cappedAdd(
        cellTotal,
        cappedAdd(valueBytes, formulaBytes, XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES),
        XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES,
      );
    }, rowTotal), 0),
  };
}

export function measureWorksheet(worksheet: Worksheet): WorksheetModelUsage & { jsonBytes: number } {
  const model = measureRows(worksheet.rows);
  const measured = measureValue(worksheet, Math.max(
    XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES,
    XLSX_MAX_MATERIALIZED_JSON_BYTES,
  ));
  return { ...model, jsonBytes: measured.jsonBytes };
}

export function addWorksheetUsage(
  current: WorksheetModelUsage,
  addition: WorksheetModelUsage,
): WorksheetModelUsage {
  return {
    rows: cappedAdd(current.rows, addition.rows, XLSX_MAX_MATERIALIZED_ROWS),
    cells: cappedAdd(current.cells, addition.cells, XLSX_MAX_MATERIALIZED_CELLS),
    ownedUtf8Bytes: cappedAdd(
      current.ownedUtf8Bytes,
      addition.ownedUtf8Bytes,
      XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES,
    ),
  };
}

export function addWorksheetCacheUsage(
  current: WorksheetCacheUsage,
  addition: Pick<WorksheetModelUsage, 'rows' | 'cells'>,
  subtraction: Partial<WorksheetCacheUsage> = {},
): WorksheetCacheUsage {
  const baseRows = current.rows - (subtraction.rows ?? 0);
  const baseCells = current.cells - (subtraction.cells ?? 0);
  if (baseRows < 0 || baseCells < 0) {
    throw new Error('worksheet cache accounting underflow');
  }
  return {
    rows: cappedAdd(baseRows, addition.rows, XLSX_MAX_CACHED_ROWS),
    cells: cappedAdd(baseCells, addition.cells, XLSX_MAX_CACHED_CELLS),
  };
}

export function worksheetLimitError(
  operation: string,
  part: string | undefined,
  resource: 'worksheet-model' | 'worksheet-cell-content' | 'worksheet-json' | 'worksheet-cache',
  metric: 'rows' | 'cells' | 'owned-utf8-bytes' | 'bytes',
  limit: number,
  observed: number,
  usage?: OoxmlResourceUsageSnapshot,
): OoxmlResourceLimitError {
  const stage = resource === 'worksheet-json' ? 'serialization' : 'parsing';
  return new OoxmlResourceLimitError(
    `OOXML resource limit exceeded${part ? ` for ${part}` : ''}: ${metric} ${observed} > ${limit}`,
    {
      stage,
      violation: {
        format: 'xlsx',
        operation,
        resource,
        metric,
        ...(part === undefined ? {} : { part }),
        limit,
        observed: Math.min(observed, limit + 1),
        configurable: false,
        usage: usage ?? ZERO_RESOURCE_USAGE,
      },
    },
  );
}

export function assertWorksheetModelUsage(
  measured: WorksheetModelUsage,
  operation: string,
  part: string | undefined,
  usage?: OoxmlResourceUsageSnapshot,
): void {
  const checks = [
    ['rows', measured.rows, XLSX_MAX_MATERIALIZED_ROWS],
    ['cells', measured.cells, XLSX_MAX_MATERIALIZED_CELLS],
    ['owned-utf8-bytes', measured.ownedUtf8Bytes, XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES],
  ] as const;
  for (const [metric, observed, limit] of checks) {
    if (observed > limit) {
      throw worksheetLimitError(
        operation,
        part,
        metric === 'owned-utf8-bytes' ? 'worksheet-cell-content' : 'worksheet-model',
        metric,
        limit,
        observed,
        usage,
      );
    }
  }
}

export function assertWorksheetJsonBytes(
  observed: number,
  operation: string,
  part: string | undefined,
  usage?: OoxmlResourceUsageSnapshot,
): void {
  if (observed > XLSX_MAX_MATERIALIZED_JSON_BYTES) {
    throw worksheetLimitError(
      operation,
      part,
      'worksheet-json',
      'bytes',
      XLSX_MAX_MATERIALIZED_JSON_BYTES,
      observed,
      usage,
    );
  }
}

export function assertWorksheetCacheUsage(
  usage: WorksheetCacheUsage,
  operation: string,
  part: string | undefined,
  resourceUsage?: OoxmlResourceUsageSnapshot,
): void {
  if (usage.rows > XLSX_MAX_CACHED_ROWS) {
    throw worksheetLimitError(
      operation, part, 'worksheet-cache', 'rows', XLSX_MAX_CACHED_ROWS, usage.rows, resourceUsage,
    );
  }
  if (usage.cells > XLSX_MAX_CACHED_CELLS) {
    throw worksheetLimitError(
      operation, part, 'worksheet-cache', 'cells', XLSX_MAX_CACHED_CELLS, usage.cells, resourceUsage,
    );
  }
}

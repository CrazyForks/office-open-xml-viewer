import type { ParsedWorkbook, Row, Worksheet } from '@silurus/ooxml-xlsx';
import type { OoxmlResourceLimits, OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  BoundedPullSession,
  decodeOoxmlResourceUsage,
  normalizeLoadResourceOptions,
  OoxmlResourceDebugSession,
  parseResourceLimitError,
  resourcePolicyForWasm,
  type PullSessionCommand,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
// The package typecheck alias maps '@silurus/ooxml-xlsx' to types.ts (types
// only), so the resolver value is imported from source directly — mirroring the
// relative WASM import below. (index.ts re-exports it for external consumers.)
import {
  resolveSharedStringRows,
  resolveSharedStrings,
} from '../../xlsx/src/shared-strings.ts';
import {
  WorksheetPullWorker,
  XLSX_WORKSHEET_PULL_BYTES,
  type WorksheetWireChunk,
} from '../../xlsx/src/worksheet-pull-worker.ts';
// @ts-ignore — wasm-pack generated JS without a d.ts entry for the bare module path
import * as xlsxWasm from '../../xlsx/src/wasm/xlsx_parser.js';
import { InProcessPullTransport } from './in-process-pull-transport.ts';
import { loadWasmModule, resolveWasm } from './wasm-loader.ts';

let initialized = false;

interface XlsxArchiveHandle {
  free(): void;
  parse(): Uint8Array;
  resource_usage(): Uint8Array;
  open_sheet_cursor(sheetIndex: number, name: string): void;
  pull_sheet_cursor(rowCredit: number): Uint8Array;
  sheet_cursor_pull_finished(): boolean;
  sheet_cursor_resource_usage(): Uint8Array;
  acknowledge_sheet_cursor_terminal(): void;
  cancel_sheet_cursor(): void;
  close_sheet_cursor(): void;
}

interface XlsxArchiveConstructor {
  new (
    data: Uint8Array,
    maxArchiveEntryBytes?: bigint | null,
    maxTotalInflatedBytes?: bigint | null,
  ): XlsxArchiveHandle;
}

/** Options for the bounded Node worksheet iterator. */
export interface XlsxWorksheetRowIteratorOptions {
  /** Package-level inflated ZIP admission limits. */
  resourceLimits?: OoxmlResourceLimits;
  /** @deprecated Use `resourceLimits.maxArchiveEntryBytes`. */
  maxZipEntryBytes?: number;
  /** Emit a data-safe resource-usage card when iteration finishes. */
  debug?: boolean;
  /** Abort the current pull and cancel the native worksheet cursor. */
  signal?: AbortSignal;
}

export type XlsxWorksheetRowChunk =
  | {
      readonly kind: 'rows';
      readonly rows: Row[];
      readonly sequence: number;
      readonly wireBytes: number;
      readonly usage?: OoxmlResourceUsageSnapshot;
    }
  | {
      /** The terminal worksheet contains layout/tail data and always has zero rows. */
      readonly kind: 'finished';
      readonly worksheet: Worksheet;
      readonly sequence: number;
      readonly wireBytes: number;
      readonly usage?: OoxmlResourceUsageSnapshot;
    };

function ensureInit(): void {
  if (initialized) return;
  const wasmPath = resolveWasm(import.meta.url, '../../xlsx/src/wasm/xlsx_parser_bg.wasm');
  loadWasmModule(xlsxWasm as unknown as { initSync: (m: WebAssembly.Module) => unknown }, wasmPath);
  initialized = true;
}

/** Parse the workbook index (sheet list + styles + shared strings) from a
 *  `.xlsx` archive. Individual sheet cell data is parsed lazily via
 *  {@link parseSheet}. */
export function parseXlsx(buffer: ArrayBuffer | Uint8Array | Buffer): ParsedWorkbook {
  ensureInit();
  const bytes = toUint8(buffer);
  // `parse_xlsx` returns UTF-8 JSON bytes (Result<Vec<u8>, JsValue>); decode +
  // parse once. Matches the browser main-thread receiver.
  const json = (xlsxWasm as unknown as { parse_xlsx: (b: Uint8Array) => Uint8Array }).parse_xlsx(
    bytes,
  );
  return JSON.parse(new TextDecoder().decode(json)) as ParsedWorkbook;
}

/** Parse a single sheet's cell model without resolving shared-string cells
 *  (they stay `{type:'shared',si}`). Internal — callers resolve against the
 *  workbook's sharedStrings table. */
function parseSheetRaw(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  sheetIndex: number,
  sheetName: string,
): Worksheet {
  ensureInit();
  const bytes = toUint8(buffer);
  // `parse_sheet` returns UTF-8 JSON bytes (Result<Vec<u8>, JsValue>); decode +
  // parse once.
  const json = (xlsxWasm as unknown as {
    parse_sheet: (b: Uint8Array, idx: number, name: string) => Uint8Array;
  }).parse_sheet(bytes, sheetIndex, sheetName);
  return JSON.parse(new TextDecoder().decode(json)) as Worksheet;
}

/** Parse a single sheet's cell data and layout. The browser path does this
 *  on demand from a Web Worker; in Node we just call the WASM export
 *  synchronously. Shared-string cells are resolved to concrete text against the
 *  workbook table (matching the browser `XlsxWorkbook` path), so callers reading
 *  `cell.value` always get `{type:'text',...}` rather than a `{type:'shared'}`
 *  reference. */
export function parseSheet(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  sheetIndex: number,
  sheetName: string,
): Worksheet {
  const ws = parseSheetRaw(buffer, sheetIndex, sheetName);
  return resolveSharedStrings(ws, parseXlsx(buffer).sharedStrings);
}

/** Eagerly parse every sheet referenced by the workbook. Useful for batch
 *  jobs (diffing two workbooks, dumping to markdown) where you want the
 *  whole model in one go. */
export function parseXlsxAllSheets(
  buffer: ArrayBuffer | Uint8Array | Buffer,
): { workbook: ParsedWorkbook['workbook']; worksheets: Record<string, Worksheet> } {
  const parsed = parseXlsx(buffer);
  const worksheets: Record<string, Worksheet> = {};
  for (let i = 0; i < parsed.workbook.sheets.length; i++) {
    const meta = parsed.workbook.sheets[i];
    // Resolve against the already-parsed table (avoids re-parsing the workbook
    // index once per sheet, which routing through `parseSheet` would do).
    worksheets[meta.name] = resolveSharedStrings(
      parseSheetRaw(buffer, i, meta.name),
      parsed.sharedStrings,
    );
  }
  return { workbook: parsed.workbook, worksheets };
}

/**
 * Stream one worksheet as bounded complete-row batches. The final yielded value
 * is a row-free worksheet containing the sheet's layout and ancillary model.
 * Every row batch is provisional until a successful terminal worksheet is
 * received. If the terminal worksheet has `parseError`, consumers must discard
 * every previously yielded row batch and use that terminal placeholder.
 * Advancing past that terminal value commits the native operation; returning or
 * throwing from the consumer first cancels it. Existing synchronous parsing
 * helpers intentionally remain materializing compatibility APIs.
 */
export async function* iterateXlsxWorksheetRows(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  sheetIndex: number,
  options: XlsxWorksheetRowIteratorOptions = {},
): AsyncGenerator<XlsxWorksheetRowChunk, void, void> {
  const resourceOptions = normalizeLoadResourceOptions(options);
  const debug = new OoxmlResourceDebugSession({
    enabled: resourceOptions.debug,
    format: 'xlsx',
    mode: 'node',
    policy: resourceOptions.policy,
  });
  debug.setSourceBytes(toUint8(buffer).byteLength);
  try {
    if (!Number.isSafeInteger(sheetIndex) || sheetIndex < 0) {
      throw new RangeError('sheetIndex must be a non-negative safe integer');
    }
    ensureInit();
    const [maxEntry, maxTotal] = resourcePolicyForWasm(resourceOptions.policy);
    const Archive = (xlsxWasm as unknown as { XlsxArchive: XlsxArchiveConstructor }).XlsxArchive;
    const archive = createArchive(Archive, toUint8(buffer), maxEntry, maxTotal);
  let terminalAcknowledged = false;
  let operationError: unknown;
  let session: BoundedPullSession<ArrayBuffer, number> | undefined;
  let pull: WorksheetPullWorker | undefined;
  let rowBatches = 0;
  let emittedRows = 0;
  try {
    const workbook = JSON.parse(new TextDecoder().decode(archive.parse())) as ParsedWorkbook;
    debug.observeUsage(decodeUsage(archive.resource_usage()));
    debug.checkpoint('workbook index ready');
    const sheet = workbook.workbook.sheets[sheetIndex];
    if (!sheet) throw new RangeError(`Sheet index ${sheetIndex} out of range`);
    const identity = { sessionId: 1, operationId: 1, generation: 1 } as const;
    pull = new WorksheetPullWorker(() => archive);
    pull.reserveOpen(identity);
    await pull.open(sheetIndex, sheet.name, identity);
    const transport = new InProcessPullTransport<PullSessionResponse<ArrayBuffer, number>>(
      (command, respond) => pull?.dispatch(
        command as PullSessionCommand<number>,
        respond,
      ),
      () => { void pull?.reset(); },
    );
    session = new BoundedPullSession(transport, {
      ...identity,
      maxByteCredit: XLSX_WORKSHEET_PULL_BYTES,
    });

    for (;;) {
      const chunk = await session.pull(XLSX_WORKSHEET_PULL_BYTES, { signal: options.signal });
      debug.observeUsage(chunk.usage);
      const decoded = JSON.parse(
        new TextDecoder().decode(new Uint8Array(chunk.payload)),
      ) as WorksheetWireChunk;
      if (chunk.done !== (decoded.kind === 'finished')) {
        throw new Error('worksheet cursor terminal marker mismatch');
      }
      if (decoded.kind === 'rows') {
        // The browser compatibility path resolves these after assembly. Node's
        // row iterator resolves each bounded batch before handing it over.
        resolveSharedStringRows(decoded.rows, workbook.sharedStrings);
        rowBatches += 1;
        emittedRows += decoded.rows.length;
        yield {
          kind: 'rows',
          rows: decoded.rows,
          sequence: chunk.sequence,
          wireBytes: chunk.byteLength,
          usage: chunk.usage,
        };
        await chunk.ack({ signal: options.signal });
        continue;
      }
      decoded.worksheet.rows = [];
      yield {
        kind: 'finished',
        worksheet: decoded.worksheet,
        sequence: chunk.sequence,
        wireBytes: chunk.byteLength,
        usage: chunk.usage,
      };
      await chunk.ack({ signal: options.signal });
      terminalAcknowledged = true;
      return;
    }
  } catch (error) {
    operationError = parseResourceLimitError(error) ?? error;
    debug.fail(operationError);
    await session?.cancel('request-error').catch(() => undefined);
    throw operationError;
  } finally {
    if (session && !terminalAcknowledged) {
      await session.cancel('closed').catch(() => undefined);
    }
    await pull?.reset().catch(() => undefined);
    try {
      archive.free();
    } catch (cleanupError) {
      if (operationError === undefined) throw cleanupError;
    }
    if (operationError === undefined) {
      debug.checkpoint(
        terminalAcknowledged ? 'worksheet stream complete' : 'worksheet stream closed',
      );
      debug.succeed({
        'row-batches': rowBatches,
        rows: emittedRows,
        completed: terminalAcknowledged ? 1 : 0,
      });
    }
  }
  } catch (error) {
    const normalized = parseResourceLimitError(error) ?? error;
    debug.fail(normalized);
    throw normalized;
  }
}

function createArchive(
  Archive: XlsxArchiveConstructor,
  bytes: Uint8Array,
  maxEntry: bigint,
  maxTotal: bigint,
): XlsxArchiveHandle {
  try {
    return new Archive(bytes, maxEntry, maxTotal);
  } catch (error) {
    throw parseResourceLimitError(error) ?? error;
  }
}

function decodeUsage(bytes: Uint8Array): OoxmlResourceUsageSnapshot | undefined {
  try {
    return decodeOoxmlResourceUsage(bytes);
  } catch (error) {
    if (String(error).includes('worksheet cursor usage is unavailable')) return undefined;
    throw error;
  }
}

function toUint8(buffer: ArrayBuffer | Uint8Array | Buffer): Uint8Array {
  return buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer as ArrayBuffer);
}

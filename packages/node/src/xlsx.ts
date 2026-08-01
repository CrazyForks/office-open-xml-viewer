import type { ParsedWorkbook, Row, Worksheet } from '@silurus/ooxml-xlsx';
import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  BoundedPullSession,
  decodeOoxmlResourceUsage,
  normalizeLoadResourceOptions,
  OoxmlResourceMetricsSession,
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
import type { OoxmlNodeSessionOptions } from './session-options.ts';
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

/** Options for the bounded Node workbook session. */
export type OpenXlsxWorkbookOptions = OoxmlNodeSessionOptions;

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

/**
 * An explicitly-owned Node workbook session. The workbook index and shared
 * strings are parsed once; worksheet bodies are pulled as bounded row batches
 * from the retained archive. Only one worksheet stream may be active at a time
 * because the native archive owns one worksheet cursor.
 */
export interface XlsxWorkbookSession {
  readonly sheetCount: number;
  readonly sheetNames: ReadonlyArray<string>;
  readonly resourceUsage: OoxmlResourceUsageSnapshot | undefined;
  worksheetRows(sheetIndex: number): AsyncGenerator<XlsxWorksheetRowChunk, void, void>;
  close(): Promise<void>;
}

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
 * Open a bounded XLSX workbook session. Unlike the materializing compatibility
 * helpers, this retains one archive, parses the workbook index once, and lets
 * callers consume worksheet rows sequentially without reopening the package.
 */
export async function openXlsxWorkbook(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  options: OpenXlsxWorkbookOptions = {},
): Promise<XlsxWorkbookSession> {
  const resourceOptions = normalizeLoadResourceOptions(options);
  const metrics = new OoxmlResourceMetricsSession({
    enabled: resourceOptions.debug || resourceOptions.onResourceMetrics !== undefined,
    format: 'xlsx',
    mode: 'node',
    scope: 'session',
    policy: resourceOptions.policy,
    onMetrics: resourceOptions.onResourceMetrics,
    emitToConsole: resourceOptions.debug,
  });
  const bytes = toUint8(buffer);
  metrics.setSourceBytes(bytes.byteLength);
  let archive: XlsxArchiveHandle | undefined;
  try {
    throwIfAborted(options.signal);
    ensureInit();
    const [maxEntry, maxTotal] = resourcePolicyForWasm(resourceOptions.policy);
    const Archive = (xlsxWasm as unknown as { XlsxArchive: XlsxArchiveConstructor }).XlsxArchive;
    archive = createArchive(Archive, bytes, maxEntry, maxTotal);
    const workbook = JSON.parse(new TextDecoder().decode(archive.parse())) as ParsedWorkbook;
    const usage = decodeUsage(archive.resource_usage());
    metrics.observeUsage(usage);
    metrics.checkpoint('workbook index ready');
    return new XlsxWorkbookSessionImpl(archive, workbook, metrics, usage, options.signal);
  } catch (error) {
    try {
      archive?.free();
    } catch {
      // Preserve the open/index failure.
    }
    const normalized = parseResourceLimitError(error) ?? error;
    metrics.fail(normalized);
    throw normalized;
  }
}

type ActiveWorksheetOperation = {
  readonly identity: Readonly<{ sessionId: number; operationId: number; generation: number }>;
  client?: BoundedPullSession<ArrayBuffer, number>;
  terminalAcknowledged: boolean;
  cleanupPromise?: Promise<void>;
};

class XlsxWorkbookSessionImpl implements XlsxWorkbookSession {
  readonly sheetCount: number;
  readonly sheetNames: ReadonlyArray<string>;

  private readonly pull: WorksheetPullWorker;
  private readonly transport: InProcessPullTransport<PullSessionResponse<ArrayBuffer, number>>;
  private workbook: ParsedWorkbook | null;
  private nextOperationId = 1;
  private active: ActiveWorksheetOperation | undefined;
  private closed = false;
  private closePromise: Promise<void> | undefined;
  private lastUsage: OoxmlResourceUsageSnapshot | undefined;
  private completedWorksheets = 0;
  private rowBatches = 0;
  private emittedRows = 0;

  constructor(
    private readonly archive: XlsxArchiveHandle,
    workbook: ParsedWorkbook,
    private readonly metrics: OoxmlResourceMetricsSession,
    usage: OoxmlResourceUsageSnapshot | undefined,
    private readonly signal?: AbortSignal,
  ) {
    this.workbook = workbook;
    this.sheetNames = Object.freeze(workbook.workbook.sheets.map((sheet) => sheet.name));
    this.sheetCount = this.sheetNames.length;
    this.lastUsage = usage;
    this.pull = new WorksheetPullWorker(() => this.archive);
    this.transport = new InProcessPullTransport(
      (command, respond) => this.pull.dispatchSafely(
        command as PullSessionCommand<number>,
        respond,
      ),
      () => undefined,
    );
  }

  get resourceUsage(): OoxmlResourceUsageSnapshot | undefined {
    if (this.closed) return this.lastUsage;
    try {
      this.lastUsage = decodeUsage(this.archive.resource_usage());
    } catch {
      // Keep the last valid diagnostic after a trapped or closing archive.
    }
    return this.lastUsage;
  }

  /**
   * Stream one worksheet as bounded complete-row batches. The final yielded
   * value is a row-free worksheet containing layout and ancillary model data.
   * Row batches remain provisional until the consumer advances past that
   * terminal value, which acknowledges the native operation.
   */
  async *worksheetRows(
    sheetIndex: number,
  ): AsyncGenerator<XlsxWorksheetRowChunk, void, void> {
    if (this.closed) throw new Error('XLSX workbook session is closed');
    if (this.active) throw new Error('another XLSX worksheet row stream is already active');
    if (!Number.isSafeInteger(sheetIndex) || sheetIndex < 0) {
      throw new RangeError('sheetIndex must be a non-negative safe integer');
    }
    const sheetName = this.requireSheetName(sheetIndex);
    throwIfAborted(this.signal);

    const operationId = this.nextOperationId++;
    const operation: ActiveWorksheetOperation = {
      identity: { sessionId: operationId, operationId, generation: 1 },
      terminalAcknowledged: false,
    };
    this.active = operation;
    let operationError: unknown;
    try {
      this.pull.reserveOpen(operation.identity);
      await this.pull.open(sheetIndex, sheetName, operation.identity);
      const client = new BoundedPullSession(this.transport, {
        ...operation.identity,
        maxByteCredit: XLSX_WORKSHEET_PULL_BYTES,
      });
      operation.client = client;

      for (;;) {
        if (this.closed) throw new Error('XLSX workbook session is closed');
        throwIfAborted(this.signal);
        const chunk = await client.pull(XLSX_WORKSHEET_PULL_BYTES, { signal: this.signal });
        this.lastUsage = chunk.usage ?? this.lastUsage;
        this.metrics.observeUsage(chunk.usage);
        const decoded = JSON.parse(
          new TextDecoder().decode(new Uint8Array(chunk.payload)),
        ) as WorksheetWireChunk;
        if (chunk.done !== (decoded.kind === 'finished')) {
          throw new Error('worksheet cursor terminal marker mismatch');
        }
        if (decoded.kind === 'rows') {
          resolveSharedStringRows(decoded.rows, this.requireWorkbook().sharedStrings);
          this.rowBatches += 1;
          this.emittedRows += decoded.rows.length;
          yield {
            kind: 'rows',
            rows: decoded.rows,
            sequence: chunk.sequence,
            wireBytes: chunk.byteLength,
            usage: chunk.usage,
          };
          await chunk.ack({ signal: this.signal });
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
        await chunk.ack({ signal: this.signal });
        operation.terminalAcknowledged = true;
        this.completedWorksheets += 1;
        this.metrics.checkpoint('worksheet stream complete', chunk.usage);
        return;
      }
    } catch (error) {
      operationError = parseResourceLimitError(error) ?? error;
      this.metrics.fail(operationError);
      await this.close().catch(() => undefined);
      throw operationError;
    } finally {
      try {
        await this.cleanupOperation(operation, operationError === undefined ? 'closed' : 'request-error');
      } catch (cleanupError) {
        if (operationError === undefined) {
          const normalized = parseResourceLimitError(cleanupError) ?? cleanupError;
          this.metrics.fail(normalized);
          await this.close().catch(() => undefined);
          throw normalized;
        }
      }
    }
  }

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise;
    this.closed = true;
    this.closePromise = this.release();
    return this.closePromise;
  }

  private cleanupOperation(
    operation: ActiveWorksheetOperation,
    reason: 'closed' | 'request-error',
  ): Promise<void> {
    if (operation.cleanupPromise) return operation.cleanupPromise;
    operation.cleanupPromise = (async () => {
      let cleanupError: unknown;
      if (operation.client && !operation.terminalAcknowledged) {
        try {
          await operation.client.cancel(reason);
        } catch (error) {
          cleanupError = parseResourceLimitError(error) ?? error;
        }
      }
      try {
        await this.pull.reset();
      } catch (error) {
        cleanupError ??= parseResourceLimitError(error) ?? error;
      }
      if (this.active === operation) this.active = undefined;
      if (cleanupError !== undefined) throw cleanupError;
    })();
    return operation.cleanupPromise;
  }

  private async release(): Promise<void> {
    let cleanupError: unknown;
    if (this.active) {
      try {
        await this.cleanupOperation(this.active, 'closed');
      } catch (error) {
        cleanupError = parseResourceLimitError(error) ?? error;
      }
    }
    this.transport.terminate();
    this.workbook = null;
    try {
      this.archive.free();
    } catch (error) {
      cleanupError ??= parseResourceLimitError(error) ?? error;
    }
    if (cleanupError !== undefined) {
      this.metrics.fail(cleanupError);
      throw cleanupError;
    }
    this.metrics.checkpoint('workbook session closed', this.lastUsage);
    this.metrics.succeed({
      worksheets: this.completedWorksheets,
      'row-batches': this.rowBatches,
      rows: this.emittedRows,
    });
  }

  private requireWorkbook(): ParsedWorkbook {
    if (!this.workbook) throw new Error('XLSX workbook session is closed');
    return this.workbook;
  }

  private requireSheetName(sheetIndex: number): string {
    const sheet = this.requireWorkbook().workbook.sheets[sheetIndex];
    if (!sheet) throw new RangeError(`Sheet index ${sheetIndex} out of range`);
    return sheet.name;
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

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  const error = new Error('XLSX workbook session was aborted');
  error.name = 'AbortError';
  throw error;
}

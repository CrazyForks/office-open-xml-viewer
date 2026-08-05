import type { ParsedWorkbook, Row, Worksheet } from '@silurus/ooxml-xlsx';
import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  BoundedPullSession,
  decodeOoxmlResourceUsage,
  parseResourceLimitError,
  type OoxmlResourceMetricsSession,
  type PullSessionCommand,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
// The Node facade consumes the format-owned coordinator through its explicit
// internal entry point rather than reconstructing XLSX orchestration here.
import {
  acquireXlsxNodeSession,
  resolveSharedStringRows,
  WorksheetPullWorker,
  XLSX_WORKSHEET_PULL_BYTES,
  type WorksheetWireChunk,
  type XlsxNodeArchive,
} from '@silurus/ooxml-xlsx/internal/session';
import { InProcessPullTransport } from '@silurus/ooxml-core/internal/in-process-pull-transport';
import type { OoxmlNodeSessionOptions } from './session-options.ts';
import { compileWasmModule, resolveWasm } from './wasm-loader.ts';
import { usingOwnedSession } from '@silurus/ooxml-core/internal/owned-session';

const xlsxWasmModule = compileWasmModule(resolveWasm(
    import.meta.url,
    'xlsx_parser_bg.wasm',
    '@silurus/ooxml-xlsx/wasm-binary',
  ));

export type DeepReadonly<T> =
  T extends (...args: never[]) => unknown ? T
    : T extends readonly (infer TValue)[] ? readonly DeepReadonly<TValue>[]
      : T extends object ? { readonly [TKey in keyof T]: DeepReadonly<T[TKey]> }
        : T;

export type ReadonlyParsedWorkbook = DeepReadonly<ParsedWorkbook>;

export interface MaterializedXlsxWorkbook {
  readonly workbookIndex: ParsedWorkbook;
  /** Caller-owned worksheets in workbook sheet-index order. */
  readonly worksheets: readonly Worksheet[];
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
  readonly workbookIndex: ReadonlyParsedWorkbook;
  readonly sheetCount: number;
  readonly sheetNames: ReadonlyArray<string>;
  readonly resourceUsage: OoxmlResourceUsageSnapshot | undefined;
  worksheetRows(sheetIndex: number): AsyncGenerator<XlsxWorksheetRowChunk, void, void>;
  close(): Promise<void>;
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
  const bytes = toUint8(buffer);
  const acquired = await acquireXlsxNodeSession(bytes, xlsxWasmModule, options);
  return new XlsxWorkbookSessionImpl(
    acquired.closeArchive,
    acquired.archive,
    acquired.workbookIndex,
    acquired.metrics,
    acquired.usage,
    options.signal,
  );
}

type ActiveWorksheetOperation = {
  readonly identity: Readonly<{ sessionId: number; operationId: number; generation: number }>;
  client?: BoundedPullSession<ArrayBuffer, number>;
  terminalAcknowledged: boolean;
  cleanupPromise?: Promise<void>;
};

class XlsxWorkbookSessionImpl implements XlsxWorkbookSession {
  readonly workbookIndex: ReadonlyParsedWorkbook;
  readonly sheetCount: number;
  readonly sheetNames: ReadonlyArray<string>;

  private readonly pull: WorksheetPullWorker;
  private readonly transport: InProcessPullTransport<PullSessionResponse<ArrayBuffer, number>>;
  private nextOperationId = 1;
  private active: ActiveWorksheetOperation | undefined;
  private closed = false;
  private closePromise: Promise<void> | undefined;
  private lastUsage: OoxmlResourceUsageSnapshot | undefined;
  private completedWorksheets = 0;
  private rowBatches = 0;
  private emittedRows = 0;

  constructor(
    private readonly closeArchive: () => void,
    private readonly archive: XlsxNodeArchive,
    workbook: ParsedWorkbook,
    private readonly metrics: OoxmlResourceMetricsSession,
    usage: OoxmlResourceUsageSnapshot | undefined,
    private readonly signal?: AbortSignal,
  ) {
    this.workbookIndex = freezeRecursively(workbook);
    this.sheetNames = Object.freeze(this.workbookIndex.workbook.sheets.map((sheet) => sheet.name));
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
          resolveSharedStringRows(
            decoded.rows,
            this.workbookIndex.sharedStrings as ParsedWorkbook['sharedStrings'],
          );
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
    try {
      this.closeArchive();
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

  private requireSheetName(sheetIndex: number): string {
    const sheet = this.workbookIndex.workbook.sheets[sheetIndex];
    if (!sheet) throw new RangeError(`Sheet index ${sheetIndex} out of range`);
    return sheet.name;
  }
}

/** Materialize only the workbook metadata/index through the canonical owned
 * archive. No worksheet cursor is opened. */
export async function materializeXlsxWorkbookIndex(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  options: OpenXlsxWorkbookOptions = {},
): Promise<ParsedWorkbook> {
  return usingOwnedSession(
    () => openXlsxWorkbook(buffer, options),
    async (session) => structuredClone(session.workbookIndex) as ParsedWorkbook,
  );
}

/** Materialize one caller-owned worksheet by sheet index. */
export async function materializeXlsxWorksheet(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  sheetIndex: number,
  options: OpenXlsxWorkbookOptions = {},
): Promise<Worksheet> {
  return usingOwnedSession(
    () => openXlsxWorkbook(buffer, options),
    (session) => materializeWorksheetFromSession(session, sheetIndex),
  );
}

/** Materialize the complete workbook without reopening the package per sheet. */
export async function materializeXlsxWorkbook(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  options: OpenXlsxWorkbookOptions = {},
): Promise<MaterializedXlsxWorkbook> {
  return usingOwnedSession(
    () => openXlsxWorkbook(buffer, options),
    async (session) => {
    const worksheets: Worksheet[] = [];
    for (let sheetIndex = 0; sheetIndex < session.sheetCount; sheetIndex += 1) {
      worksheets.push(await materializeWorksheetFromSession(session, sheetIndex));
    }
    return {
      workbookIndex: structuredClone(session.workbookIndex) as ParsedWorkbook,
      worksheets: Object.freeze(worksheets),
    };
    },
  );
}

async function materializeWorksheetFromSession(
  session: XlsxWorkbookSession,
  sheetIndex: number,
): Promise<Worksheet> {
  const rows: Row[] = [];
  let terminal: Worksheet | undefined;
  for await (const chunk of session.worksheetRows(sheetIndex)) {
    if (chunk.kind === 'rows') rows.push(...chunk.rows);
    else terminal = chunk.worksheet;
  }
  if (!terminal) throw new Error(`XLSX worksheet ${sheetIndex} did not produce a terminal model`);
  terminal.rows = terminal.parseError ? [] : rows;
  return structuredClone(terminal);
}

function freezeRecursively<T>(value: T, seen = new WeakSet<object>()): DeepReadonly<T> {
  if (value === null || typeof value !== 'object') return value as DeepReadonly<T>;
  const object = value as object;
  if (seen.has(object)) return value as DeepReadonly<T>;
  seen.add(object);
  for (const child of Object.values(object)) freezeRecursively(child, seen);
  return Object.freeze(value) as DeepReadonly<T>;
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

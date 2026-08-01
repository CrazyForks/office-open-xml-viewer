import type { DocxDocumentModel } from '@silurus/ooxml-docx';
import type {
  OoxmlResourceLimits,
  OoxmlResourceUsageSnapshot,
} from '@silurus/ooxml-core';
import {
  dropBitmapCacheByPath,
  dropSvgImageCache,
  OoxmlResourceLimitError,
} from '@silurus/ooxml-core';
import {
  decodeOoxmlResourceUsage,
  normalizeLoadResourceOptions,
  OoxmlResourceDebugSession,
  parseResourceLimitError,
  resourcePolicyForWasm,
  type PullSessionCommand,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import { normalizeDocxDocumentModel } from '../../docx/src/parser-model.ts';
import { materializeDocumentPullSession } from '../../docx/src/document-pull-client.ts';
import {
  DocumentPullWorker,
  type DocxDocumentCursorArchive,
} from '../../docx/src/document-pull-worker.ts';
import { layoutSourceModelAdapter } from '../../docx/src/layout-source-model-adapter.ts';
import { createLayoutServices } from '../../docx/src/layout-runtime.ts';
import { retainRenderWorkerDocumentLayout } from '../../docx/src/render-worker-layout.ts';
import {
  dropColorReplacedCache,
  renderLayoutSourceToCanvas,
  type DocxTextRunInfo,
} from '../../docx/src/renderer.ts';
// @ts-ignore — wasm-pack generated JS without a d.ts entry for the bare module path
import * as docxWasm from '../../docx/src/wasm/docx_parser.js';
import { InProcessPullTransport } from './in-process-pull-transport.ts';
import {
  installImageBitmapShim,
  installOffscreenCanvasShim,
  type NodeCanvasFactory,
  type NodeCanvasLike,
} from './render.ts';
import { loadWasmModule, resolveWasm } from './wasm-loader.ts';

let initialized = false;
let nodeCanvasRuntimeTail: Promise<void> = Promise.resolve();

interface DocxArchiveHandle extends DocxDocumentCursorArchive {
  free(): void;
  extract_image(path: string): Uint8Array;
  document_cursor_resource_usage(): Uint8Array;
}

interface DocxArchiveConstructor {
  new (
    data: Uint8Array,
    maxArchiveEntryBytes?: bigint | null,
    maxTotalInflatedBytes?: bigint | null,
  ): DocxArchiveHandle;
}

/** Options for the bounded Node DOCX page session. */
export interface OpenDocxDocumentOptions {
  /** Canvas implementation used for text measurement and page allocation. */
  factory: NodeCanvasFactory;
  /** Package-level inflated ZIP admission limits. */
  resourceLimits?: OoxmlResourceLimits;
  /** @deprecated Use `resourceLimits.maxArchiveEntryBytes`. */
  maxZipEntryBytes?: number;
  /** Emit a data-safe resource-usage card to the Node console. */
  debug?: boolean;
  /** Stable DATE/TIME field instant captured before pagination. */
  currentDate?: Date | number;
  /** Abort between parser units or page renders. Synchronous WASM work cannot be preempted. */
  signal?: AbortSignal;
}

export interface DocxPageRenderOptions {
  width?: number;
  dpr?: number;
  defaultTextColor?: string;
  showTrackChanges?: boolean;
  onTextRun?: (run: DocxTextRunInfo) => void;
}

export interface DocxRenderedPage {
  readonly pageIndex: number;
  readonly widthPt: number;
  readonly heightPt: number;
  /** Caller-owned after yield; retaining it is outside the session memory bound. */
  readonly canvas: NodeCanvasLike;
}

/**
 * Ready, paginated Node document. DOCX pagination is necessarily sequential and
 * retained because preceding flow determines later pages and total page count.
 * Parsing and transfer are nevertheless bounded: no whole document XML/JSON
 * value crosses the cursor, and only one page canvas is created per pull.
 */
export interface DocxDocumentSession extends AsyncIterable<DocxRenderedPage> {
  readonly pageCount: number;
  readonly resourceUsage: OoxmlResourceUsageSnapshot | undefined;
  pageSize(pageIndex: number): Readonly<{ widthPt: number; heightPt: number }>;
  renderPage(pageIndex: number, options?: DocxPageRenderOptions): Promise<NodeCanvasLike>;
  pages(options?: DocxPageRenderOptions): AsyncGenerator<DocxRenderedPage, void, void>;
  close(): Promise<void>;
}

function ensureInit(): void {
  if (initialized) return;
  const wasmPath = resolveWasm(import.meta.url, '../../docx/src/wasm/docx_parser_bg.wasm');
  loadWasmModule(docxWasm as unknown as { initSync: (m: WebAssembly.Module) => unknown }, wasmPath);
  initialized = true;
}

/** Parse a `.docx` archive in Node and return the same `DocxDocumentModel` the
 *  browser path produces. */
export function parseDocx(buffer: ArrayBuffer | Uint8Array | Buffer): DocxDocumentModel {
  ensureInit();
  const bytes =
    buffer instanceof Uint8Array
      ? buffer
      : new Uint8Array(buffer as ArrayBuffer);
  // `parse_docx` returns UTF-8 JSON bytes (Result<Vec<u8>, JsValue>); decode +
  // parse once. This is the intentionally materializing compatibility path;
  // `openDocxDocument` uses the bounded cursor instead.
  const json = (docxWasm as unknown as { parse_docx: (b: Uint8Array) => Uint8Array }).parse_docx(
    bytes,
  );
  return normalizeDocxDocumentModel(
    JSON.parse(new TextDecoder().decode(json)) as DocxDocumentModel,
  );
}

/**
 * Open a Node DOCX session that parses through the same acknowledged body cursor
 * as the browser Viewer, seals the canonical layout source, completes pagination,
 * and then renders one caller-owned canvas at a time. Existing `parseDocx`
 * remains the synchronous materializing compatibility API.
 */
export async function openDocxDocument(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  options: OpenDocxDocumentOptions,
): Promise<DocxDocumentSession> {
  const resourceOptions = normalizeLoadResourceOptions(options ?? {});
  const debug = new OoxmlResourceDebugSession({
    enabled: resourceOptions.debug,
    format: 'docx',
    mode: 'node',
    policy: resourceOptions.policy,
  });
  debug.setSourceBytes(toUint8(buffer).byteLength);
  const Archive = (docxWasm as unknown as { DocxArchive: DocxArchiveConstructor }).DocxArchive;
  let archive: DocxArchiveHandle | undefined;
  let pull: DocumentPullWorker | undefined;
  let transport: InProcessPullTransport<PullSessionResponse<ArrayBuffer, number>> | undefined;
  try {
    if (!options?.factory) throw new TypeError('openDocxDocument requires a canvas factory');
    throwIfAborted(options.signal);
    ensureInit();
    const [maxEntry, maxTotal] = resourcePolicyForWasm(resourceOptions.policy);
    archive = new Archive(toUint8(buffer), maxEntry, maxTotal);
    debug.checkpoint('container ready');
    pull = new DocumentPullWorker(() => archive);
    const identity = { sessionId: 1, operationId: 1, generation: 1 } as const;
    pull.open(identity);
    transport = new InProcessPullTransport(
      (command, respond) => pull?.dispatch(
        command as PullSessionCommand<number>,
        respond,
      ),
      () => undefined,
    );
    let usage: OoxmlResourceUsageSnapshot | undefined;
    const model = await materializeDocumentPullSession(transport, identity, {
      signal: options.signal,
      onUsage: (checkpoint) => {
        usage = checkpoint;
        debug.observeUsage(checkpoint);
      },
    });
    usage ??= decodeDocumentUsage(archive.document_cursor_resource_usage());
    debug.observeUsage(usage);
    debug.checkpoint('model streamed');
    await pull.reset();
    transport.terminate();

    throwIfAborted(options.signal);
    const adapted = layoutSourceModelAdapter(model);
    const measurementCanvas = options.factory.createCanvas(1, 1);
    const services = createLayoutServices(adapted.source, {
      measureContext: measurementCanvas.getContext('2d'),
    });
    const defaultCurrentDateMs = normalizeCurrentDate(options.currentDate);
    const retained = retainRenderWorkerDocumentLayout(
      adapted.source,
      services,
      defaultCurrentDateMs,
    );
    const layout = retained.layoutVariants.defaultLayout;
    const session = new DocxDocumentSessionImpl(
      archive,
      adapted.source,
      services,
      layout,
      options.factory,
      defaultCurrentDateMs,
      usage,
      options.signal,
    );
    debug.checkpoint('pagination ready');
    debug.succeed({ pages: session.pageCount });
    return session;
  } catch (error) {
    await pull?.reset().catch(() => undefined);
    transport?.terminate();
    try {
      archive?.free();
    } catch {
      // Preserve the open/parse/layout failure.
    }
    const normalized = parseResourceLimitError(error) ?? error;
    debug.fail(normalized);
    throw normalized;
  }
}

type SessionState = Readonly<{
  source: ReturnType<typeof layoutSourceModelAdapter>['source'];
  services: ReturnType<typeof createLayoutServices>;
}>;

type DefaultDocumentLayout =
  ReturnType<typeof retainRenderWorkerDocumentLayout>['layoutVariants']['defaultLayout'];

class DocxDocumentSessionImpl implements DocxDocumentSession {
  readonly pageCount: number;
  readonly resourceUsage: OoxmlResourceUsageSnapshot | undefined;
  private readonly sizes: ReadonlyArray<Readonly<{ widthPt: number; heightPt: number }>>;
  private state: SessionState | null;
  private renderTail: Promise<void> = Promise.resolve();
  private pagesStarted = false;
  private closed = false;
  private closePromise: Promise<void> | undefined;
  private resourceFailure: OoxmlResourceLimitError | null = null;
  private readonly fetchImage = async (path: string, mimeType: string): Promise<Blob> => {
    const bytes = this.archive.extract_image(path);
    return new Blob([new Uint8Array(bytes).slice() as BlobPart], { type: mimeType });
  };

  constructor(
    private readonly archive: DocxArchiveHandle,
    source: SessionState['source'],
    services: SessionState['services'],
    layout: DefaultDocumentLayout,
    private readonly factory: NodeCanvasFactory,
    private readonly defaultCurrentDateMs: number,
    usage: OoxmlResourceUsageSnapshot | undefined,
    private readonly signal?: AbortSignal,
  ) {
    this.state = { source, services };
    this.pageCount = layout.pages.length;
    this.resourceUsage = usage;
    this.sizes = Object.freeze(layout.pages.map((page) => Object.freeze({
      widthPt: page.geometry.widthPt,
      heightPt: page.geometry.heightPt,
    })));
  }

  pageSize(pageIndex: number): Readonly<{ widthPt: number; heightPt: number }> {
    const size = this.sizes[pageIndex];
    if (!size) throw new RangeError(`DOCX page index ${pageIndex} out of range`);
    return size;
  }

  [Symbol.asyncIterator](): AsyncGenerator<DocxRenderedPage, void, void> {
    return this.pages();
  }

  renderPage(pageIndex: number, options: DocxPageRenderOptions = {}): Promise<NodeCanvasLike> {
    if (this.closed) return Promise.reject(new Error('DOCX document session is closed'));
    if (this.resourceFailure) return Promise.reject(this.resourceFailure);
    this.pageSize(pageIndex);
    return this.enqueueRender(async () => {
      throwIfAborted(this.signal);
      const state = this.requireState();
      const canvas = this.factory.createCanvas(1, 1);
      await withNodeCanvasRuntime(this.factory, () => renderLayoutSourceToCanvas(
        state.source,
        canvas as unknown as HTMLCanvasElement,
        pageIndex,
        {
          ...options,
          currentDate: this.defaultCurrentDateMs,
          defaultCurrentDateMs: this.defaultCurrentDateMs,
          layoutServices: state.services,
          fetchImage: this.fetchImage,
        },
      ));
      throwIfAborted(this.signal);
      return canvas;
    }).catch((error: unknown) => {
      const normalized = parseResourceLimitError(error) ?? error;
      if (normalized instanceof OoxmlResourceLimitError) {
        this.resourceFailure ??= normalized;
      }
      throw normalized;
    });
  }

  async *pages(options: DocxPageRenderOptions = {}): AsyncGenerator<DocxRenderedPage, void, void> {
    if (this.closed) throw new Error('DOCX document session is closed');
    if (this.pagesStarted) throw new Error('DOCX page stream is one-pass and was already consumed');
    this.pagesStarted = true;
    let operationError: unknown;
    try {
      for (let pageIndex = 0; pageIndex < this.pageCount; pageIndex += 1) {
        const canvas = await this.renderPage(pageIndex, options);
        const size = this.pageSize(pageIndex);
        yield { pageIndex, ...size, canvas };
      }
    } catch (error) {
      operationError = parseResourceLimitError(error) ?? error;
      throw operationError;
    } finally {
      try {
        await this.close();
      } catch (cleanupError) {
        if (operationError === undefined) throw cleanupError;
      }
    }
  }

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise;
    this.closed = true;
    this.closePromise = this.release();
    return this.closePromise;
  }

  private enqueueRender<T>(operation: () => Promise<T>): Promise<T> {
    const result = this.renderTail.then(operation, operation);
    this.renderTail = result.then(() => undefined, () => undefined);
    return result;
  }

  private async release(): Promise<void> {
    await this.renderTail;
    dropBitmapCacheByPath(this.fetchImage);
    dropColorReplacedCache(this.fetchImage);
    dropSvgImageCache(this.fetchImage);
    this.state = null;
    try {
      this.archive.free();
    } catch (error) {
      throw parseResourceLimitError(error) ?? error;
    }
  }

  private requireState(): SessionState {
    if (!this.state) throw new Error('DOCX document session is closed');
    return this.state;
  }
}

function withNodeCanvasRuntime<T>(
  factory: NodeCanvasFactory,
  operation: () => Promise<T>,
): Promise<T> {
  const run = async (): Promise<T> => {
    const restoreImageBitmap = typeof globalThis.createImageBitmap === 'function'
      ? () => undefined
      : installImageBitmapShim(factory);
    const restoreOffscreen = installOffscreenCanvasShim(factory);
    try {
      return await operation();
    } finally {
      restoreOffscreen();
      restoreImageBitmap();
    }
  };
  const result = nodeCanvasRuntimeTail.then(run, run);
  nodeCanvasRuntimeTail = result.then(() => undefined, () => undefined);
  return result;
}

function normalizeCurrentDate(value: Date | number | undefined): number {
  const current = value instanceof Date ? value.getTime() : (value ?? Date.now());
  if (!Number.isFinite(current)) {
    throw new RangeError('currentDate must resolve to finite epoch milliseconds');
  }
  return current;
}

function decodeDocumentUsage(bytes: Uint8Array): OoxmlResourceUsageSnapshot | undefined {
  try {
    return decodeOoxmlResourceUsage(bytes);
  } catch (error) {
    if (String(error).includes('document cursor usage is unavailable')) return undefined;
    throw error;
  }
}

function toUint8(buffer: ArrayBuffer | Uint8Array | Buffer): Uint8Array {
  return buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer as ArrayBuffer);
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  const error = new Error('DOCX document session was aborted');
  error.name = 'AbortError';
  throw error;
}

import type {
  OoxmlResourceLimits,
  OoxmlResourceMetrics,
  OoxmlResourceUsageSnapshot,
} from '@silurus/ooxml-core';
import type { Presentation, Slide } from '@silurus/ooxml-pptx';
import {
  normalizeLoadResourceOptions,
  OoxmlResourceMetricsSession,
  parseResourceLimitError,
  resourcePolicyForWasm,
  type PullSessionCommand,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore — wasm-pack generated JS without a d.ts entry for the bare module path
import * as pptxWasm from '../../pptx/src/wasm/pptx_parser.js';
import { normalizePresentationBootstrap } from '../../pptx/src/presentation-preflight.ts';
import { PptxSlidePullClient } from '../../pptx/src/slide-pull-client.ts';
import {
  readPptxSlideCursorUsage,
  type PptxSlideCursorArchive,
} from '../../pptx/src/slide-cursor-operation.ts';
import { SlidePullWorker } from '../../pptx/src/slide-pull-worker.ts';
import type { PresentationBootstrap } from '../../pptx/src/worker-protocol.ts';
import { InProcessPullTransport } from './in-process-pull-transport.ts';
import { loadWasmModule, resolveWasm } from './wasm-loader.ts';

let initialized = false;

interface PptxArchiveHandle extends PptxSlideCursorArchive {
  free(): void;
  assert_healthy(): void;
  presentation_bootstrap(): Uint8Array;
}

interface PptxArchiveConstructor {
  new (
    data: Uint8Array,
    maxArchiveEntryBytes?: bigint | null,
    maxTotalInflatedBytes?: bigint | null,
  ): PptxArchiveHandle;
}

/** Options for the bounded Node presentation session. */
export interface OpenPptxPresentationOptions {
  /** Package-level inflated ZIP admission limits. */
  resourceLimits?: OoxmlResourceLimits;
  /** @deprecated Use `resourceLimits.maxArchiveEntryBytes`. */
  maxZipEntryBytes?: number;
  /** Emit a content-free resource-usage card when the session closes. */
  debug?: boolean;
  /** Receive the machine-readable session report without enabling console output. */
  onResourceMetrics?: (metrics: OoxmlResourceMetrics) => void;
  /** Abort the current slide pull and close the session. */
  signal?: AbortSignal;
}

/**
 * One-pass Node session over canonical complete PPTX slide units. The library
 * retains at most the package bootstrap and one yielded slide at a time. Slides
 * deliberately remain ordinary objects: copies retained by the caller are
 * caller-owned and are outside this session's memory bound.
 */
export interface PptxPresentationSession extends AsyncIterable<Slide> {
  readonly slideCount: number;
  readonly slideWidth: number;
  readonly slideHeight: number;
  readonly resourceUsage: OoxmlResourceUsageSnapshot | undefined;
  slides(): AsyncGenerator<Slide, void, void>;
  close(): Promise<void>;
}

function ensureInit(): void {
  if (initialized) return;
  // The wasm-pack `--target web` JS module exports `initSync` and the
  // parser functions. Locate the sibling `.wasm` file in the pptx package
  // and feed its bytes into `initSync` so the module is fully linked
  // before the first `parse_pptx` call.
  const wasmPath = resolveWasm(import.meta.url, '../../pptx/src/wasm/pptx_parser_bg.wasm');
  loadWasmModule(pptxWasm as unknown as { initSync: (m: WebAssembly.Module) => unknown }, wasmPath);
  initialized = true;
}

/** Parse a `.pptx` archive in Node and return the same `Presentation` model
 *  the browser path produces. Synchronous WASM init is performed on first call
 *  and cached for subsequent invocations. */
export function parsePptx(buffer: ArrayBuffer | Uint8Array | Buffer): Presentation {
  ensureInit();
  const bytes =
    buffer instanceof Uint8Array
      ? buffer
      : new Uint8Array(buffer as ArrayBuffer);
  // `parse_pptx` returns UTF-8 JSON bytes (Result<Vec<u8>, JsValue>); decode +
  // parse once. Matches the browser main-thread receiver.
  const json = (pptxWasm as unknown as { parse_pptx: (b: Uint8Array) => Uint8Array }).parse_pptx(
    bytes,
  );
  return JSON.parse(new TextDecoder().decode(json)) as Presentation;
}

/**
 * Open a one-pass, pull-based PPTX session for Node batch rendering. Existing
 * materializing helpers remain unchanged; this additive path avoids retaining a
 * complete `Presentation` while callers render or otherwise consume each slide.
 * Exhausting or breaking the iterator closes the retained WASM archive.
 */
export async function openPptxPresentation(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  options: OpenPptxPresentationOptions = {},
): Promise<PptxPresentationSession> {
  const resourceOptions = normalizeLoadResourceOptions(options);
  const metrics = new OoxmlResourceMetricsSession({
    enabled: resourceOptions.debug || resourceOptions.onResourceMetrics !== undefined,
    format: 'pptx',
    mode: 'node',
    scope: 'session',
    policy: resourceOptions.policy,
    onMetrics: resourceOptions.onResourceMetrics,
    emitToConsole: resourceOptions.debug,
  });
  metrics.setSourceBytes(toUint8(buffer).byteLength);
  const Archive: PptxArchiveConstructor = pptxWasm.PptxArchive;
  let archive: PptxArchiveHandle | undefined;
  try {
    throwIfAborted(options.signal);
    ensureInit();
    const [maxEntry, maxTotal] = resourcePolicyForWasm(resourceOptions.policy);
    archive = new Archive(toUint8(buffer), maxEntry, maxTotal);
    const bootstrap = normalizePresentationBootstrap(JSON.parse(
      new TextDecoder().decode(archive.presentation_bootstrap()),
    ) as PresentationBootstrap);
    metrics.checkpoint('presentation bootstrap ready');
    return new PptxPresentationSessionImpl(archive, bootstrap, metrics, options.signal);
  } catch (error) {
    try {
      archive?.free();
    } catch {
      // Preserve the open/bootstrap failure.
    }
    const normalized = parseResourceLimitError(error) ?? error;
    metrics.fail(normalized);
    throw normalized;
  }
}

class PptxPresentationSessionImpl implements PptxPresentationSession {
  readonly slideCount: number;
  readonly slideWidth: number;
  readonly slideHeight: number;

  private readonly slidePull: SlidePullWorker;
  private readonly slideClient: PptxSlidePullClient;
  private readonly transport: InProcessPullTransport<PullSessionResponse<ArrayBuffer, number>>;
  private started = false;
  private closed = false;
  private closePromise: Promise<void> | undefined;
  private usage: OoxmlResourceUsageSnapshot | undefined;
  private consumedSlides = 0;

  constructor(
    private readonly archive: PptxArchiveHandle,
    bootstrap: PresentationBootstrap,
    private readonly metrics: OoxmlResourceMetricsSession,
    private readonly signal?: AbortSignal,
  ) {
    this.slideCount = bootstrap.slideCount;
    this.slideWidth = bootstrap.slideWidth;
    this.slideHeight = bootstrap.slideHeight;
    this.slidePull = new SlidePullWorker(() => this.archive);
    this.transport = new InProcessPullTransport(
      (command, respond) => this.slidePull.dispatchSafely(
        command as PullSessionCommand<number>,
        respond,
      ),
      () => undefined,
    );
    this.slideClient = new PptxSlidePullClient({
      slideCount: this.slideCount,
      transport: this.transport,
      open: async (slideIndex, identity) => {
        this.slidePull.reserveOpen(identity);
        await this.slidePull.open(slideIndex, identity);
      },
      onUsage: (usage) => {
        this.usage = usage;
        this.metrics.observeUsage(usage);
      },
    });
  }

  get resourceUsage(): OoxmlResourceUsageSnapshot | undefined {
    return this.usage;
  }

  [Symbol.asyncIterator](): AsyncGenerator<Slide, void, void> {
    return this.slides();
  }

  async *slides(): AsyncGenerator<Slide, void, void> {
    if (this.closed) throw new Error('PPTX presentation session is closed');
    if (this.started) throw new Error('PPTX presentation session is one-pass and was already consumed');
    this.started = true;
    let operationError: unknown;
    try {
      for (let index = 0; index < this.slideCount; index += 1) {
        throwIfAborted(this.signal);
        const slide = await this.slideClient.load(index);
        if (!slide) throw new Error(`PPTX slide ${index} was not decoded`);
        this.usage ??= await this.slidePull.run(() => readPptxSlideCursorUsage(
          (operation) => operation(this.archive),
        ));
        this.metrics.observeUsage(this.usage);
        this.consumedSlides = index + 1;
        yield slide;
      }
    } catch (error) {
      operationError = parseResourceLimitError(error) ?? error;
      this.metrics.fail(operationError);
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
    this.slideClient.cancelAll();
    this.closePromise = this.release();
    return this.closePromise;
  }

  private async release(): Promise<void> {
    let operationError: unknown;
    try {
      await this.slidePull.reset();
    } catch (error) {
      operationError = parseResourceLimitError(error) ?? error;
    }
    this.transport.terminate();
    try {
      this.archive.free();
    } catch (cleanupError) {
      operationError ??= parseResourceLimitError(cleanupError) ?? cleanupError;
    }
    if (operationError !== undefined) {
      this.metrics.fail(operationError);
      throw operationError;
    }
    this.metrics.checkpoint('presentation session closed');
    this.metrics.succeed({ slides: this.consumedSlides });
  }
}

/** Extract raw bytes for a single media entry (e.g. `ppt/media/image1.png`)
 *  from the source archive. Mirrors `extract_media` on the browser worker. */
export function extractMedia(buffer: ArrayBuffer | Uint8Array | Buffer, path: string): Uint8Array {
  ensureInit();
  const bytes =
    buffer instanceof Uint8Array
      ? buffer
      : new Uint8Array(buffer as ArrayBuffer);
  return (pptxWasm as unknown as { extract_media: (b: Uint8Array, p: string) => Uint8Array }).extract_media(bytes, path);
}

/** Extract raw bytes for a single embedded image entry (e.g.
 *  `ppt/media/image1.png`) from the source archive. Mirrors `extract_image`
 *  on the browser worker (twin of {@link extractMedia}); pictures and blip
 *  fills now carry only zip paths, so the render path reads image bytes lazily
 *  through this. `maxZipEntryBytes` mirrors the worker's per-entry guard and is
 *  optional; omission selects the shared standard policy default. */
export function extractImage(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  path: string,
  maxZipEntryBytes?: number,
): Uint8Array {
  ensureInit();
  const bytes =
    buffer instanceof Uint8Array
      ? buffer
      : new Uint8Array(buffer as ArrayBuffer);
  return (
    pptxWasm as unknown as {
      extract_image: (b: Uint8Array, p: string, max?: bigint) => Uint8Array;
    }
  ).extract_image(
    bytes,
    path,
    typeof maxZipEntryBytes === 'number' && maxZipEntryBytes > 0
      ? BigInt(maxZipEntryBytes)
      : undefined,
  );
}

function toUint8(buffer: ArrayBuffer | Uint8Array | Buffer): Uint8Array {
  return buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer as ArrayBuffer);
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  const error = new Error('PPTX presentation session was aborted');
  error.name = 'AbortError';
  throw error;
}

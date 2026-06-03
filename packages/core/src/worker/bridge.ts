/**
 * Request/response correlation over a Web Worker.
 *
 * Each format package (pptx/docx/xlsx) drives a WASM parser in a Worker. The
 * naive pattern — attach a one-shot `message` listener keyed only on the
 * response `type` — breaks under concurrency: with two requests in flight, the
 * first arriving response of a matching type resolves the wrong promise (or, if
 * an unrelated message arrives, the listener detaches without resolving and the
 * promise hangs forever).
 *
 * `WorkerBridge` fixes this by assigning every request a monotonic id and
 * resolving against a pending-callback map keyed by that id — the proven
 * pattern. It is wire-protocol agnostic: the discriminant field (`kind` vs
 * `type`), the error shape, and any unsolicited messages (e.g. an init `ready`
 * handshake) are described by the {@link WorkerBridgeOptions} callbacks, so all
 * three packages can share one correlation mechanism without standardizing
 * their message envelopes.
 */

/** The subset of the DOM `Worker` interface the bridge depends on. Keeping it
 *  structural lets tests substitute an in-memory fake. */
export interface WorkerLike {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  addEventListener(type: 'message', listener: (e: MessageEvent) => void): void;
  removeEventListener(type: 'message', listener: (e: MessageEvent) => void): void;
  terminate(): void;
}

export interface WorkerBridgeOptions<TRes> {
  /**
   * Extract the correlation id from a response. Return `undefined` for
   * unsolicited messages that do not answer a request (e.g. a `ready`
   * handshake); those are routed to {@link onUnsolicited} instead.
   */
  readonly correlate: (response: TRes) => number | undefined;
  /**
   * Extract an error message from a response, or `undefined` when the response
   * is a success. When defined, the matching request rejects with `Error(msg)`.
   */
  readonly toError?: (response: TRes) => string | undefined;
  /** Called for every message that does not correlate to a pending request. */
  readonly onUnsolicited?: (response: TRes) => void;
}

export class WorkerBridge<TRes = unknown> {
  private readonly _worker: WorkerLike;
  private readonly _opts: WorkerBridgeOptions<TRes>;
  private readonly _pending = new Map<
    number,
    { resolve: (r: TRes) => void; reject: (e: Error) => void }
  >();
  private _nextId = 1;

  constructor(worker: WorkerLike, opts: WorkerBridgeOptions<TRes>) {
    this._worker = worker;
    this._opts = opts;
    this._worker.addEventListener('message', this._handle);
  }

  private _handle = (e: MessageEvent<TRes>): void => {
    const res = e.data;
    const id = this._opts.correlate(res);
    if (id === undefined) {
      this._opts.onUnsolicited?.(res);
      return;
    }
    const cb = this._pending.get(id);
    if (!cb) return; // unknown / already-settled id: ignore (never hang another request)
    this._pending.delete(id);
    const err = this._opts.toError?.(res);
    if (err !== undefined) cb.reject(new Error(err));
    else cb.resolve(res);
  };

  /** Allocate the next correlation id. Useful when the caller must embed the id
   *  in a transferable-bearing message it builds itself. */
  nextId(): number {
    return this._nextId++;
  }

  /**
   * Send a correlated request and resolve with its matching response. `build`
   * receives the freshly allocated id so it can embed it in the message.
   */
  request(build: (id: number) => unknown, transfer?: Transferable[]): Promise<TRes> {
    const id = this._nextId++;
    return new Promise<TRes>((resolve, reject) => {
      this._pending.set(id, { resolve, reject });
      this._worker.postMessage(build(id), transfer);
    });
  }

  /** Fire-and-forget message with no correlation (e.g. the `init` message). */
  post(message: unknown, transfer?: Transferable[]): void {
    this._worker.postMessage(message, transfer);
  }

  /** Terminate the worker and reject every still-pending request. */
  terminate(): void {
    this._worker.removeEventListener('message', this._handle);
    this._worker.terminate();
    for (const cb of this._pending.values()) {
      cb.reject(new Error('Worker terminated'));
    }
    this._pending.clear();
  }
}

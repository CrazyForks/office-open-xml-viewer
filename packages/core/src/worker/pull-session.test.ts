import { describe, expect, it, vi } from 'vitest';
import { OoxmlResourceLimitError } from '../errors/ooxml-error.js';
import { WorkerBridge, type WorkerLike } from './bridge.js';
import { serializeWorkerError } from './error-wire.js';
import {
  BoundedPullSession,
  PullSessionHost,
  PullSessionHostCoordinator,
  PULL_SESSION_PROTOCOL,
  type PullSessionCommand,
  type PullSessionHostDriver,
  type PullSessionHostOptions,
  type PullSessionIdentity,
  type PullSessionKey,
  type PullSessionPost,
  type PullSessionResponse,
} from './pull-session.js';

type Payload = { value: string };
type Response = PullSessionResponse<Payload, string>;

class FakeWorker implements WorkerLike {
  posted: PullSessionCommand<string>[] = [];
  terminated = 0;
  private listeners = new Set<(event: MessageEvent) => void>();
  onPost?: (command: PullSessionCommand<string>) => void;
  postMessage(message: unknown): void {
    const command = message as PullSessionCommand<string>;
    this.posted.push(command);
    this.onPost?.(command);
  }
  addEventListener(type: 'message' | 'messageerror' | 'error', listener: (e: never) => void): void {
    if (type === 'message') this.listeners.add(listener as (event: MessageEvent) => void);
  }
  removeEventListener(type: 'message' | 'messageerror' | 'error', listener: (e: never) => void): void {
    if (type === 'message') this.listeners.delete(listener as (event: MessageEvent) => void);
  }
  terminate(): void {
    this.terminated++;
  }
  respond(data: Response): void {
    for (const listener of this.listeners) listener({ data } as MessageEvent);
  }
}

/** Test adapter that captures dispatch output only after the post callback succeeds. */
class TestPullSessionHost<
  TPayload,
  TSessionId extends PullSessionKey = PullSessionKey,
> extends PullSessionHost<TPayload, TSessionId> {
  constructor(
    options: Omit<PullSessionHostOptions<TPayload, TSessionId>, 'coordinator' | 'driver'> & {
      coordinator?: PullSessionHostCoordinator;
      driver: Omit<PullSessionHostDriver<TPayload>, 'measureChunk'> & {
        measureChunk?: PullSessionHostDriver<TPayload>['measureChunk'];
      };
    },
  ) {
    const driver = options.driver;
    super({
      ...options,
      coordinator: options.coordinator ?? new PullSessionHostCoordinator(),
      driver: {
        ...driver,
        pull: async (credit) => {
          const chunk = await driver.pull(credit);
          return chunk.leaseId !== undefined && chunk.retainedBytes === undefined
            ? { ...chunk, retainedBytes: chunk.byteLength }
            : chunk;
        },
        measureChunk: driver.measureChunk ?? ((chunk) => chunk.byteLength),
      },
    });
  }

  async handle(command: PullSessionCommand<TSessionId>): Promise<PullSessionPost<TPayload, TSessionId>> {
    let output: PullSessionPost<TPayload, TSessionId> | undefined;
    await this.dispatch(command, (response, transfer) => {
      output = { response, transfer };
    });
    if (!output) throw new Error('dispatch did not post');
    return output;
  }
}

const identity: PullSessionIdentity<string> = { sessionId: 'session', operationId: 7, generation: 3 };
const usage = {
  archiveEntryCount: 2,
  declaredInflatedBytes: 10,
  distinctInflatedBytes: 8,
  operationInflatedBytes: 4,
};

type WithoutEnvelope<T> = T extends unknown
  ? Omit<T, keyof PullSessionIdentity<string> | 'requestId' | 'protocol'>
  : never;

function response(
  request: PullSessionCommand<string>,
  value: WithoutEnvelope<Response>,
): Response {
  return { ...value, protocol: PULL_SESSION_PROTOCOL, ...identity, requestId: request.requestId } as Response;
}

function setupClient(disposeTransferred = vi.fn(), timeoutMs?: number, cancelGraceMs?: number) {
  const worker = new FakeWorker();
  const bridge = new WorkerBridge<Response>(worker, {
    correlate: (message) => message.requestId,
  });
  const session = new BoundedPullSession(bridge, {
    ...identity,
    maxByteCredit: 64,
    timeoutMs,
    cancelGraceMs,
    disposeTransferred,
  });
  return { worker, bridge, session, disposeTransferred };
}

function command<TKind extends PullSessionCommand['kind']>(
  kind: TKind,
  fields: Omit<
    Extract<PullSessionCommand<string>, { kind: TKind }>,
    keyof typeof identity | 'kind' | 'protocol'
  >,
): Extract<PullSessionCommand<string>, { kind: TKind }> {
  return { kind, protocol: PULL_SESSION_PROTOCOL, ...identity, ...fields } as Extract<
    PullSessionCommand<string>,
    { kind: TKind }
  >;
}

describe('BoundedPullSession client', () => {
  it('validates bounded credit before sending', async () => {
    const { worker, session } = setupClient();
    await expect(session.pull(0)).rejects.toThrow(/positive/i);
    await expect(session.pull(65)).rejects.toThrow(/maximum/i);
    expect(worker.posted).toEqual([]);
  });

  it('waits for correlated ACK acceptance before permitting the next pull', async () => {
    const { worker, session } = setupClient();
    const firstPromise = session.pull(16);
    const firstRequest = worker.posted[0];
    worker.respond(
      response(firstRequest, {
        kind: 'chunk', sequence: 0, byteLength: 4, done: false, payload: { value: 'one' }, leaseId: 11,
      }),
    );
    const first = await firstPromise;
    const ackPromise = first.ack();
    expect(first.ack()).toBe(ackPromise);
    await expect(session.pull(16)).rejects.toThrow(/acknowledge/i);
    const ackRequest = worker.posted[1];
    expect(ackRequest).toMatchObject({ kind: 'ack', sequence: 0, operationId: 7, generation: 3 });
    worker.respond(response(ackRequest, { kind: 'accepted', command: 'ack' }));
    await ackPromise;
    expect(worker.posted.filter((item) => item.kind === 'ack')).toHaveLength(1);

    void session.pull(8);
    expect(worker.posted[2]).toMatchObject({ kind: 'pull', sequence: 1, byteCredit: 8 });
  });

  it('supports a final data-bearing chunk which remains blocked until acknowledged', async () => {
    const { worker, session } = setupClient();
    const pending = session.pull(16);
    const pull = worker.posted[0];
    worker.respond(
      response(pull, {
        kind: 'chunk', sequence: 0, byteLength: 3, done: true, payload: { value: 'end' }, leaseId: 12,
      }),
    );
    const final = await pending;
    expect(final).toMatchObject({ done: true, payload: { value: 'end' } });
    const ack = final.ack();
    const ackRequest = worker.posted[1];
    worker.respond(response(ackRequest, { kind: 'accepted', command: 'ack' }));
    await ack;
    await expect(session.pull(16)).rejects.toThrow(/closed/i);
  });

  it('separates transferred payload disposal from correlated numeric lease release', async () => {
    const { worker, session, disposeTransferred } = setupClient();
    const pending = session.pull(16);
    const pull = worker.posted[0];
    worker.respond(
      response(pull, {
        kind: 'chunk', sequence: 0, byteLength: 2, done: false, payload: { value: 'x' }, leaseId: 44,
      }),
    );
    const chunk = await pending;
    chunk.disposeTransferred();
    chunk.disposeTransferred();
    expect(disposeTransferred).toHaveBeenCalledOnce();
    expect(worker.posted).toHaveLength(1);

    const release = chunk.release();
    const releaseRequest = worker.posted[1];
    expect(releaseRequest).toMatchObject({ kind: 'release', leaseId: 44 });
    worker.respond(response(releaseRequest, { kind: 'accepted', command: 'release' }));
    await release;
    await chunk.release();
    expect(worker.posted.filter((item) => item.kind === 'release')).toHaveLength(1);
  });

  it('disposes a stale-generation response without exposing it', async () => {
    const { worker, session, disposeTransferred } = setupClient();
    const pending = session.pull(8);
    const pull = worker.posted[0];
    worker.respond({
      protocol: PULL_SESSION_PROTOCOL,
      kind: 'chunk',
      sessionId: identity.sessionId,
      operationId: identity.operationId,
      generation: identity.generation - 1,
      requestId: pull.requestId,
      sequence: 0,
      byteLength: 2,
      done: false,
      payload: { value: 'stale' },
    });
    await expect(pending).rejects.toThrow(/stale/i);
    expect(disposeTransferred).toHaveBeenCalledWith({ value: 'stale' });
    expect(worker.posted.filter((item) => item.kind === 'cancel')).toHaveLength(1);
  });

  it('deduplicates cancel and close commands', async () => {
    const canceled = setupClient();
    const firstCancel = canceled.session.cancel('abort');
    const secondCancel = canceled.session.cancel('abort');
    expect(secondCancel).toBe(firstCancel);
    expect(canceled.session.close()).toBe(firstCancel);
    const cancelRequest = canceled.worker.posted[0];
    canceled.worker.respond(response(cancelRequest, { kind: 'accepted', command: 'cancel' }));
    await Promise.all([firstCancel, secondCancel]);
    expect(canceled.worker.posted).toHaveLength(1);

    const closed = setupClient();
    const firstClose = closed.session.close();
    const secondClose = closed.session.close();
    expect(secondClose).toBe(firstClose);
    expect(closed.session.cancel()).toBe(firstClose);
    const closeRequest = closed.worker.posted[0];
    closed.worker.respond(response(closeRequest, { kind: 'accepted', command: 'close' }));
    await Promise.all([firstClose, secondClose]);
    expect(closed.worker.posted).toHaveLength(1);
  });

  it('converges timeout and abort on correlated cancellation', async () => {
    vi.useFakeTimers();
    try {
      const timed = setupClient(vi.fn(), 10);
      const pending = timed.session.pull(8);
      const rejection = expect(pending).rejects.toThrow(/timed out/i);
      await vi.advanceTimersByTimeAsync(10);
      await rejection;
      expect(timed.worker.posted.at(-1)).toMatchObject({ kind: 'cancel', reason: 'timeout' });

      const aborted = setupClient();
      const controller = new AbortController();
      const abortPending = aborted.session.pull(8, { signal: controller.signal });
      controller.abort();
      await expect(abortPending).rejects.toMatchObject({ name: 'AbortError' });
      expect(aborted.worker.posted.at(-1)).toMatchObject({ kind: 'cancel', reason: 'abort' });
    } finally {
      vi.useRealTimers();
    }
  });

  it('owns and disposes a late timed-out response through its per-request hook', async () => {
    vi.useFakeTimers();
    try {
      const { worker, bridge, session, disposeTransferred } = setupClient(vi.fn(), 10);
      const pending = session.pull(8);
      const pull = worker.posted[0];
      const rejection = expect(pending).rejects.toThrow(/timed out/i);
      await vi.advanceTimersByTimeAsync(10);
      await rejection;
      expect(bridge.orphanedRequestCount).toBe(1);
      worker.respond(
        response(pull, {
          kind: 'chunk', sequence: 0, byteLength: 1, done: false, payload: { value: 'late' },
        }),
      );
      expect(disposeTransferred).toHaveBeenCalledWith({ value: 'late' });
      expect(bridge.orphanedRequestCount).toBe(0);
    } finally {
      vi.useRealTimers();
    }
  });

  it('forgets timeout tombstones after ordered lifecycle acceptance', async () => {
    vi.useFakeTimers();
    try {
      const { worker, bridge, session } = setupClient(vi.fn(), 10);
      const pending = session.pull(8);
      const rejection = expect(pending).rejects.toThrow(/timed out/i);
      await vi.advanceTimersByTimeAsync(10);
      await rejection;
      expect(bridge.orphanedRequestCount).toBe(1);
      const cancel = worker.posted.at(-1) as PullSessionCommand<string>;
      worker.respond(response(cancel, { kind: 'accepted', command: 'cancel' }));
      await Promise.resolve();
      await Promise.resolve();
      expect(bridge.orphanedRequestCount).toBe(0);
    } finally {
      vi.useRealTimers();
    }
  });

  it('terminates the bridge when lifecycle acceptance exceeds its grace period', async () => {
    vi.useFakeTimers();
    try {
      const { worker, session } = setupClient(vi.fn(), undefined, 25);
      const pending = session.cancel('abort');
      const rejection = expect(pending).rejects.toThrow(/accept lifecycle|terminated/i);
      await vi.advanceTimersByTimeAsync(25);
      await rejection;
      expect(worker.terminated).toBe(1);
      expect(session.close()).toBe(pending);
    } finally {
      vi.useRealTimers();
    }
  });

  it('uses lifecycle grace independently of the ordinary request timeout', async () => {
    vi.useFakeTimers();
    try {
      const { worker, session } = setupClient(vi.fn(), 10, 40);
      const pending = session.close();
      const rejection = expect(pending).rejects.toThrow(/lifecycle|terminated/i);
      await vi.advanceTimersByTimeAsync(10);
      expect(worker.terminated).toBe(0);
      await vi.advanceTimersByTimeAsync(30);
      await rejection;
      expect(worker.terminated).toBe(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it('cancels when correlated ACK acceptance fails', async () => {
    const { worker, session } = setupClient();
    const pending = session.pull(8);
    const pull = worker.posted[0];
    worker.respond(
      response(pull, {
        kind: 'chunk', sequence: 0, byteLength: 1, done: false, payload: { value: 'x' },
      }),
    );
    const chunk = await pending;
    const ack = chunk.ack();
    const ackRequest = worker.posted[1];
    worker.respond(response(ackRequest, { kind: 'error', error: serializeWorkerError(new Error('ack failed')) }));
    await expect(ack).rejects.toThrow('ack failed');
    expect(worker.posted.at(-1)).toMatchObject({ kind: 'cancel', reason: 'request-error' });
    await expect(session.pull(8)).rejects.toThrow(/closed/i);
  });

  it('cancels when correlated worker-lease release fails', async () => {
    const { worker, session } = setupClient();
    const pending = session.pull(8);
    const pull = worker.posted[0];
    worker.respond(
      response(pull, {
        kind: 'chunk', sequence: 0, byteLength: 1, done: false, payload: { value: 'x' }, leaseId: 5,
      }),
    );
    const chunk = await pending;
    const release = chunk.release();
    const releaseRequest = worker.posted[1];
    worker.respond(
      response(releaseRequest, {
        kind: 'error', error: serializeWorkerError(new Error('release failed')),
      }),
    );
    await expect(release).rejects.toThrow('release failed');
    expect(worker.posted.at(-1)).toMatchObject({ kind: 'cancel', reason: 'request-error' });
    await expect(session.pull(8)).rejects.toThrow(/closed/i);
  });
});

describe('PullSessionHost', () => {
  it('serializes pulls and enforces host credit and one-unacked-chunk backpressure', async () => {
    let active = 0;
    let maximumActive = 0;
    const driver: PullSessionHostDriver<Payload> = {
      pull: async (credit) => {
        active++;
        maximumActive = Math.max(maximumActive, active);
        await Promise.resolve();
        active--;
        return { payload: { value: String(credit) }, byteLength: credit, done: false };
      },
      measureChunk: (chunk) => chunk.byteLength,
    };
    const host = new TestPullSessionHost({ ...identity, maxByteCredit: 16, driver });
    const first = host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 8 }));
    const second = host.handle(command('pull', { requestId: 2, sequence: 0, byteCredit: 8 }));
    expect((await first).response.kind).toBe('chunk');
    await expect(second).resolves.toMatchObject({
      response: { kind: 'error', error: { message: expect.stringMatching(/acknowledged/) } },
    });
    expect(maximumActive).toBe(1);
    const creditHost = new TestPullSessionHost({ ...identity, maxByteCredit: 16, driver });
    await expect(
      creditHost.handle(command('pull', { requestId: 3, sequence: 0, byteCredit: 17 })),
    ).resolves.toMatchObject({
      response: { kind: 'error', error: { message: expect.stringMatching(/maximum/) } },
    });
  });

  it('enforces measured bytes instead of trusting driver-reported byteLength', async () => {
    const releaseLease = vi.fn();
    const disposeInvalidChunk = vi.fn();
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 16,
      driver: {
        pull: () => ({ payload: { value: 'oversized' }, byteLength: 1, done: false, leaseId: 6 }),
        measureChunk: () => 9,
        releaseLease,
        disposeInvalidChunk,
      },
    });
    await expect(
      host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 8 })),
    ).resolves.toMatchObject({
      response: { kind: 'error', error: { message: expect.stringMatching(/credit/) } },
    });
    expect(releaseLease).toHaveBeenCalledWith(6);
    expect(disposeInvalidChunk).toHaveBeenCalledOnce();
  });

  it('returns a validated transfer list separately from the response envelope', async () => {
    const buffer = new ArrayBuffer(4);
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 16,
      driver: {
        pull: () => ({ payload: { value: 'binary' }, byteLength: 4, done: false, transfer: [buffer] }),
      },
    });
    const post = await host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 4 }));
    expect(post.response).toMatchObject({ kind: 'chunk', payload: { value: 'binary' } });
    expect(post.transfer).toEqual([buffer]);
  });

  it('rejects measurement below the host-counted ArrayBuffer transfer lower bound', async () => {
    const buffer = new ArrayBuffer(64);
    const disposeInvalidChunk = vi.fn();
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 64,
      driver: {
        pull: () => ({ payload: { value: 'underreported' }, byteLength: 1, done: false, transfer: [buffer] }),
        measureChunk: () => 1,
        disposeInvalidChunk,
      },
    });
    await expect(
      host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 1 })),
    ).resolves.toMatchObject({
      response: { kind: 'error', error: { message: expect.stringMatching(/ArrayBuffer/) } },
    });
    expect(disposeInvalidChunk).toHaveBeenCalledOnce();
  });

  it('attaches optional resource usage checkpoints to chunk and control responses', async () => {
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 16,
      driver: {
        pull: () => ({ payload: { value: 'x' }, byteLength: 1, done: false }),
        resourceUsage: () => usage,
      },
    });
    const pulled = await host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }));
    expect(pulled.response).toMatchObject({ kind: 'chunk', usage });
    const acked = await host.handle(command('ack', { requestId: 2, sequence: 0 }));
    expect(acked.response).toMatchObject({ kind: 'accepted', command: 'ack', usage });
  });

  it('idempotently accepts duplicate ack, release, cancel, and close', async () => {
    const acknowledge = vi.fn();
    const releaseLease = vi.fn();
    const cancel = vi.fn();
    const close = vi.fn();
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 16,
      driver: {
        pull: () => ({ payload: { value: 'x' }, byteLength: 1, done: false, leaseId: 9 }),
        acknowledge,
        releaseLease,
        cancel,
        close,
      },
    });
    await host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 8 }));
    await host.handle(command('ack', { requestId: 2, sequence: 0 }));
    await host.handle(command('ack', { requestId: 3, sequence: 0 }));
    await host.handle(command('release', { requestId: 4, leaseId: 1 }));
    await host.handle(command('release', { requestId: 5, leaseId: 1 }));
    await host.handle(command('cancel', { requestId: 6, reason: 'abort' }));
    await host.handle(command('cancel', { requestId: 7, reason: 'abort' }));
    await host.handle(command('close', { requestId: 8 }));
    await host.handle(command('close', { requestId: 9 }));
    expect(acknowledge).toHaveBeenCalledTimes(1);
    expect(releaseLease).toHaveBeenCalledTimes(1);
    expect(cancel).toHaveBeenCalledTimes(1);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('maps monotonic wire lease IDs so a delayed release cannot hit a reused driver ID', async () => {
    const releaseLease = vi.fn();
    const chunks = [
      { payload: { value: 'one' }, byteLength: 1, done: false, leaseId: 7, retainedBytes: 1 },
      { payload: { value: 'two' }, byteLength: 1, done: false, leaseId: 7, retainedBytes: 1 },
    ];
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 8,
      driver: { pull: () => chunks.shift() as (typeof chunks)[number], releaseLease },
    });
    const first = await host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }));
    expect(first.response).toMatchObject({ kind: 'chunk', leaseId: 1 });
    await host.handle(command('ack', { requestId: 2, sequence: 0 }));
    await host.handle(command('release', { requestId: 3, leaseId: 1 }));
    const second = await host.handle(command('pull', { requestId: 4, sequence: 1, byteCredit: 1 }));
    expect(second.response).toMatchObject({ kind: 'chunk', leaseId: 2 });
    await host.handle(command('release', { requestId: 5, leaseId: 1 }));
    expect(releaseLease).toHaveBeenCalledTimes(1);
    await host.handle(command('release', { requestId: 6, leaseId: 2 }));
    expect(releaseLease).toHaveBeenCalledTimes(2);
    expect(releaseLease).toHaveBeenLastCalledWith(7);
  });

  it('cancels the operation without individually releasing an active duplicate driver lease', async () => {
    const releaseLease = vi.fn();
    const cancel = vi.fn();
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 8,
      driver: {
        pull: vi
          .fn()
          .mockReturnValueOnce({ payload: { value: 'one' }, byteLength: 1, done: false, leaseId: 7, retainedBytes: 1 })
          .mockReturnValueOnce({ payload: { value: 'duplicate' }, byteLength: 1, done: false, leaseId: 7, retainedBytes: 1 }),
        releaseLease,
        cancel,
      },
    });
    await host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }));
    await host.handle(command('ack', { requestId: 2, sequence: 0 }));
    await expect(
      host.handle(command('pull', { requestId: 3, sequence: 1, byteCredit: 1 })),
    ).resolves.toMatchObject({ response: { kind: 'error', error: { message: expect.stringMatching(/duplicate/) } } });
    expect(releaseLease).toHaveBeenCalledTimes(1);
    expect(releaseLease).toHaveBeenCalledWith(7);
    expect(cancel).toHaveBeenCalledOnce();
  });

  it('rejects a new lease when the generation wire ID space is exhausted', async () => {
    const releaseLease = vi.fn();
    const chunks = [
      { payload: { value: 'max' }, byteLength: 1, done: false, leaseId: 1, retainedBytes: 1 },
      { payload: { value: 'overflow' }, byteLength: 1, done: false, leaseId: 2, retainedBytes: 1 },
    ];
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 8,
      wireLeaseIdStart: Number.MAX_SAFE_INTEGER,
      driver: { pull: () => chunks.shift() as (typeof chunks)[number], releaseLease },
    });
    await host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }));
    await host.handle(command('ack', { requestId: 2, sequence: 0 }));
    await host.handle(command('release', { requestId: 3, leaseId: Number.MAX_SAFE_INTEGER }));
    await expect(
      host.handle(command('pull', { requestId: 4, sequence: 1, byteCredit: 1 })),
    ).resolves.toMatchObject({
      response: { kind: 'error', error: { message: expect.stringMatching(/space exhausted/) } },
    });
    expect(releaseLease).toHaveBeenLastCalledWith(2);
  });

  it('requires an ACK for a final payload and then refuses another pull', async () => {
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 16,
      driver: { pull: () => ({ payload: { value: 'last bytes' }, byteLength: 10, done: true }) },
    });
    await expect(
      host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 10 })),
    ).resolves.toMatchObject({
      response: { kind: 'chunk', done: true, payload: { value: 'last bytes' } },
    });
    await expect(
      host.handle(command('pull', { requestId: 2, sequence: 0, byteCredit: 10 })),
    ).resolves.toMatchObject({ response: { kind: 'error' } });
    await expect(host.handle(command('ack', { requestId: 3, sequence: 0 }))).resolves.toMatchObject({
      response: { kind: 'accepted', command: 'ack' },
    });
    await expect(
      host.handle(command('pull', { requestId: 4, sequence: 1, byteCredit: 10 })),
    ).resolves.toMatchObject({
      response: { kind: 'error', error: { message: expect.stringMatching(/closed/) } },
    });
  });

  it('close releases worker-retained leases even after a final ACK', async () => {
    const releaseLease = vi.fn();
    const close = vi.fn();
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 16,
      driver: {
        pull: () => ({ payload: { value: 'last' }, byteLength: 4, done: true, leaseId: 30 }),
        releaseLease,
        close,
      },
    });
    await host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 8 }));
    await host.handle(command('ack', { requestId: 2, sequence: 0 }));
    await host.handle(command('close', { requestId: 3 }));
    await host.handle(command('close', { requestId: 4 }));
    expect(releaseLease).toHaveBeenCalledOnce();
    expect(releaseLease).toHaveBeenCalledWith(30);
    expect(close).toHaveBeenCalledOnce();
  });

  it('retries failed cancel cleanup without repeating successful hooks', async () => {
    const releaseLease = vi
      .fn<(leaseId: number) => Promise<void>>()
      .mockRejectedValueOnce(new Error('release failed'))
      .mockResolvedValue(undefined);
    const cancel = vi.fn();
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 16,
      driver: {
        pull: () => ({ payload: { value: 'held' }, byteLength: 1, done: true, leaseId: 40 }),
        releaseLease,
        cancel,
      },
    });
    await host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }));
    await host.handle(command('ack', { requestId: 2, sequence: 0 }));
    await expect(host.handle(command('cancel', { requestId: 3, reason: 'abort' }))).resolves.toMatchObject({
      response: { kind: 'error', error: { message: 'release failed' } },
    });
    await expect(host.handle(command('cancel', { requestId: 4, reason: 'abort' }))).resolves.toMatchObject({
      response: { kind: 'accepted', command: 'cancel' },
    });
    expect(releaseLease).toHaveBeenCalledTimes(2);
    expect(cancel).toHaveBeenCalledOnce();
  });

  it('retries a rejected close hook after retaining completed cleanup state', async () => {
    const releaseLease = vi.fn();
    const close = vi
      .fn<() => Promise<void>>()
      .mockRejectedValueOnce(new Error('close failed'))
      .mockResolvedValue(undefined);
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 16,
      driver: {
        pull: () => ({ payload: { value: 'held' }, byteLength: 1, done: true, leaseId: 41 }),
        releaseLease,
        close,
      },
    });
    await host.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }));
    await host.handle(command('ack', { requestId: 2, sequence: 0 }));
    await expect(host.handle(command('close', { requestId: 3 }))).resolves.toMatchObject({
      response: { kind: 'error', error: { message: 'close failed' } },
    });
    await expect(host.handle(command('close', { requestId: 4 }))).resolves.toMatchObject({
      response: { kind: 'accepted', command: 'close' },
    });
    expect(releaseLease).toHaveBeenCalledOnce();
    expect(close).toHaveBeenCalledTimes(2);
  });

  it('rejects stale operation generations before invoking the driver', async () => {
    const pull = vi.fn(() => ({ payload: { value: 'x' }, byteLength: 1, done: false }));
    const host = new TestPullSessionHost({ ...identity, maxByteCredit: 16, driver: { pull } });
    await expect(
      host.handle({ ...command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }), generation: 2 }),
    ).resolves.toMatchObject({
      response: { kind: 'error', error: { message: expect.stringMatching(/stale/) } },
    });
    expect(pull).not.toHaveBeenCalled();
  });
});

describe('PullSessionHostCoordinator', () => {
  it('runs fatal cleanup for multiple hosts sequentially', async () => {
    const coordinator = new PullSessionHostCoordinator();
    let active = 0;
    let maximumActive = 0;
    const cleanup = vi.fn(async () => {
      active++;
      maximumActive = Math.max(maximumActive, active);
      await Promise.resolve();
      active--;
    });
    const sibling = (operationId: number) =>
      new TestPullSessionHost({
        ...identity,
        operationId,
        maxByteCredit: 8,
        coordinator,
        driver: { pull: () => ({ payload: { value: 'x' }, byteLength: 1, done: false }), cancel: cleanup },
      });
    sibling(1);
    sibling(2);
    const fatalError = new OoxmlResourceLimitError('fatal serialization', {
      stage: 'decompression',
      violation: {
        format: 'docx', operation: 'parse', resource: 'archive-entry', part: 'word/a.xml',
        metric: 'actual-inflated-bytes', limit: 1, observed: 2, configurable: true, usage,
      },
    });
    const failing = new TestPullSessionHost<Payload, string>({
      ...identity,
      operationId: 3,
      maxByteCredit: 8,
      coordinator,
      driver: { pull: () => { throw fatalError; }, cancel: cleanup },
    });
    await failing.handle({
      ...command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }), operationId: 3,
    });
    expect(cleanup).toHaveBeenCalledTimes(3);
    expect(maximumActive).toBe(1);
  });

  it('queues a host registered during poison behind the active cleanup', async () => {
    const coordinator = new PullSessionHostCoordinator();
    const events: string[] = [];
    let releaseFirst!: () => void;
    const firstCleanup = new Promise<void>((resolve) => {
      releaseFirst = resolve;
    });
    new TestPullSessionHost({
      ...identity,
      operationId: 1,
      maxByteCredit: 8,
      coordinator,
      driver: {
        pull: () => ({ payload: { value: 'x' }, byteLength: 1, done: false }),
        cancel: async () => {
          events.push('first:start');
          await firstCleanup;
          events.push('first:end');
        },
      },
    });
    const fatalError = new OoxmlResourceLimitError('late registration fatal', {
      stage: 'decompression',
      violation: {
        format: 'docx', operation: 'parse', resource: 'archive-entry', part: 'word/b.xml',
        metric: 'actual-inflated-bytes', limit: 1, observed: 2, configurable: true, usage,
      },
    });
    const failing = new TestPullSessionHost<Payload, string>({
      ...identity,
      operationId: 2,
      maxByteCredit: 8,
      coordinator,
      driver: { pull: () => { throw fatalError; }, cancel: () => { events.push('failing'); } },
    });
    const poisoning = failing.handle({
      ...command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }), operationId: 2,
    });
    const queuedCommand = failing.dispatch(
      { ...command('cancel', { requestId: 2, reason: 'request-error' }), operationId: 2 },
      () => { events.push('command'); },
    );
    while (!events.includes('first:start')) await Promise.resolve();
    new TestPullSessionHost({
      ...identity,
      operationId: 3,
      maxByteCredit: 8,
      coordinator,
      driver: {
        pull: () => ({ payload: { value: 'late' }, byteLength: 1, done: false }),
        cancel: () => { events.push('late'); },
      },
    });
    expect(events).toEqual(['first:start']);
    releaseFirst();
    await poisoning;
    await queuedCommand;
    expect(events.indexOf('late')).toBeGreaterThan(events.indexOf('first:end'));
    expect(events).toEqual(['first:start', 'first:end', 'failing', 'late', 'command']);
  });

  it('unregisters completed hosts after final ACK or their last lease release', async () => {
    const withoutLease = new PullSessionHostCoordinator();
    const plain = new TestPullSessionHost({
      ...identity, maxByteCredit: 8, coordinator: withoutLease,
      driver: { pull: () => ({ payload: { value: 'done' }, byteLength: 1, done: true }) },
    });
    await plain.handle(command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }));
    await plain.handle(command('ack', { requestId: 2, sequence: 0 }));
    expect(withoutLease.registeredHostCount).toBe(0);

    const withLease = new PullSessionHostCoordinator();
    const leased = new TestPullSessionHost({
      ...identity, maxByteCredit: 8, coordinator: withLease,
      driver: { pull: () => ({ payload: { value: 'done' }, byteLength: 1, done: true, leaseId: 3, retainedBytes: 1 }) },
    });
    await leased.handle(command('pull', { requestId: 3, sequence: 0, byteCredit: 1 }));
    await leased.handle(command('ack', { requestId: 4, sequence: 0 }));
    expect(withLease.registeredHostCount).toBe(1);
    await leased.handle(command('release', { requestId: 5, leaseId: 1 }));
    expect(withLease.registeredHostCount).toBe(0);
  });

  it('accounts retained lease bytes as well as count', () => {
    const coordinator = new PullSessionHostCoordinator({ maxRetainedCount: 2, maxRetainedBytes: 4 });
    const first = Symbol('first');
    const second = Symbol('second');
    coordinator.retainLease(first, 1, 4);
    expect(() => coordinator.retainLease(second, 2, 1)).toThrow(/lease bytes/i);
    coordinator.releaseLease(first, 1);
    expect(() => coordinator.retainLease(second, 2, 1)).not.toThrow();
  });

  it('governs retained numeric lease count and bytes across operation hosts', async () => {
    const coordinator = new PullSessionHostCoordinator({ maxRetainedCount: 1, maxRetainedBytes: 4 });
    const releaseSecond = vi.fn();
    const first = new TestPullSessionHost({
      ...identity, operationId: 1, maxByteCredit: 8, coordinator,
      driver: { pull: () => ({ payload: { value: 'a' }, byteLength: 1, done: false, leaseId: 1, retainedBytes: 4 }) },
    });
    const second = new TestPullSessionHost({
      ...identity, operationId: 2, maxByteCredit: 8, coordinator,
      driver: {
        pull: () => ({ payload: { value: 'b' }, byteLength: 1, done: false, leaseId: 2, retainedBytes: 1 }),
        releaseLease: releaseSecond,
      },
    });
    await first.handle({ ...command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }), operationId: 1 });
    await first.handle({ ...command('ack', { requestId: 2, sequence: 0 }), operationId: 1 });
    await expect(
      second.handle({ ...command('pull', { requestId: 3, sequence: 0, byteCredit: 1 }), operationId: 2 }),
    ).resolves.toMatchObject({ response: { kind: 'error', error: { message: expect.stringMatching(/lease count/) } } });
    expect(releaseSecond).toHaveBeenCalledWith(2);
    await first.handle({ ...command('release', { requestId: 4, leaseId: 1 }), operationId: 1 });
    await expect(
      second.handle({ ...command('pull', { requestId: 5, sequence: 0, byteCredit: 1 }), operationId: 2 }),
    ).resolves.toMatchObject({ response: { kind: 'chunk' } });
  });

  it('poisons sibling operations on the first typed resource error and replays it', async () => {
    const coordinator = new PullSessionHostCoordinator();
    const releaseLease = vi.fn();
    const cancelHolder = vi.fn();
    const cancelFailing = vi.fn();
    const holder = new TestPullSessionHost({
      ...identity, operationId: 1, maxByteCredit: 8, coordinator,
      driver: {
        pull: () => ({ payload: { value: 'held' }, byteLength: 1, done: false, leaseId: 8, retainedBytes: 3 }),
        releaseLease,
        cancel: cancelHolder,
      },
    });
    const fatalError = new OoxmlResourceLimitError('package limit', {
      stage: 'decompression',
      violation: {
        format: 'docx', operation: 'parse', resource: 'archive-entry', part: 'word/a.xml',
        metric: 'actual-inflated-bytes', limit: 3, observed: 4, configurable: true, usage,
      },
    });
    const failing = new TestPullSessionHost<Payload, string>({
      ...identity, operationId: 2, maxByteCredit: 8, coordinator,
      driver: { pull: () => { throw fatalError; }, cancel: cancelFailing },
    });
    await holder.handle({ ...command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }), operationId: 1 });
    await holder.handle({ ...command('ack', { requestId: 2, sequence: 0 }), operationId: 1 });
    const fatal = await failing.handle({
      ...command('pull', { requestId: 3, sequence: 0, byteCredit: 1 }), operationId: 2,
    });
    expect(fatal.response).toMatchObject({
      kind: 'error', error: { code: 'ooxml-resource-limit', resourceLimit: fatalError.details },
    });
    expect(releaseLease).toHaveBeenCalledWith(8);
    expect(cancelHolder).toHaveBeenCalledOnce();
    expect(cancelFailing).toHaveBeenCalledOnce();
    const replay = await holder.handle({
      ...command('pull', { requestId: 4, sequence: 1, byteCredit: 1 }), operationId: 1,
    });
    expect(replay.response).toMatchObject({
      kind: 'error', error: { code: 'ooxml-resource-limit', message: 'package limit' },
    });
    const cleanup = await failing.handle({
      ...command('cancel', { requestId: 5, reason: 'request-error' }), operationId: 2,
    });
    expect(cleanup.response).toMatchObject({ kind: 'accepted', command: 'cancel' });
  });

  it('keeps failed fatal cleanup registered and retries it on lifecycle control', async () => {
    const coordinator = new PullSessionHostCoordinator();
    const releaseLease = vi
      .fn<(id: number) => Promise<void>>()
      .mockRejectedValueOnce(new Error('temporary release failure'))
      .mockResolvedValue(undefined);
    const holder = new TestPullSessionHost({
      ...identity, operationId: 1, maxByteCredit: 8, coordinator,
      driver: {
        pull: () => ({ payload: { value: 'held' }, byteLength: 1, done: false, leaseId: 9, retainedBytes: 1 }),
        releaseLease,
      },
    });
    const fatalError = new OoxmlResourceLimitError('fatal', {
      stage: 'decompression',
      violation: {
        format: 'docx', operation: 'parse', resource: 'archive-entry', part: 'word/a.xml',
        metric: 'actual-inflated-bytes', limit: 1, observed: 2, configurable: true, usage,
      },
    });
    const failing = new TestPullSessionHost<Payload, string>({
      ...identity, operationId: 2, maxByteCredit: 8, coordinator,
      driver: { pull: () => { throw fatalError; } },
    });
    await holder.handle({ ...command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }), operationId: 1 });
    await holder.handle({ ...command('ack', { requestId: 2, sequence: 0 }), operationId: 1 });
    await failing.handle({ ...command('pull', { requestId: 3, sequence: 0, byteCredit: 1 }), operationId: 2 });
    expect(coordinator.registeredHostCount).toBeGreaterThan(0);
    await expect(
      holder.handle({ ...command('cancel', { requestId: 4, reason: 'request-error' }), operationId: 1 }),
    ).resolves.toMatchObject({ response: { kind: 'accepted', command: 'cancel' } });
    expect(releaseLease).toHaveBeenCalledTimes(2);
  });


  it('rolls back transfer, lease, gate, and operation ownership when postMessage throws', async () => {
    const coordinator = new PullSessionHostCoordinator();
    const buffer = new ArrayBuffer(4);
    const disposeInvalidChunk = vi.fn();
    const releaseLease = vi.fn();
    const cancel = vi.fn();
    const owner = new PullSessionHost({
      ...identity,
      operationId: 1,
      maxByteCredit: 8,
      coordinator,
      driver: {
        pull: () => ({
          payload: { value: 'unpostable' },
          byteLength: 4,
          done: false,
          leaseId: 50,
          transfer: [buffer],
          retainedBytes: 4,
        }),
        measureChunk: (chunk) => chunk.byteLength,
        disposeInvalidChunk,
        releaseLease,
        cancel,
      },
    });
    const next = new TestPullSessionHost({
      ...identity,
      operationId: 2,
      maxByteCredit: 8,
      coordinator,
      driver: { pull: () => ({ payload: { value: 'next' }, byteLength: 1, done: false }) },
    });

    await expect(
      owner.dispatch(
        { ...command('pull', { requestId: 1, sequence: 0, byteCredit: 4 }), operationId: 1 },
        () => {
          throw new DOMException('could not clone', 'DataCloneError');
        },
      ),
    ).rejects.toMatchObject({ name: 'DataCloneError' });
    expect(disposeInvalidChunk).toHaveBeenCalledOnce();
    expect(disposeInvalidChunk).toHaveBeenCalledWith(expect.objectContaining({ transfer: [buffer] }));
    expect(releaseLease).toHaveBeenCalledWith(50);
    expect(cancel).toHaveBeenCalledOnce();
    await expect(
      next.handle({ ...command('pull', { requestId: 2, sequence: 0, byteCredit: 1 }), operationId: 2 }),
    ).resolves.toMatchObject({ response: { kind: 'chunk', payload: { value: 'next' } } });

    const posted: Response[] = [];
    await owner.dispatch(
      { ...command('cancel', { requestId: 3, reason: 'request-error' }), operationId: 1 },
      (response) => posted.push(response),
    );
    await owner.dispatch(
      { ...command('cancel', { requestId: 4, reason: 'request-error' }), operationId: 1 },
      (response) => posted.push(response),
    );
    expect(posted).toHaveLength(2);
    expect(cancel).toHaveBeenCalledOnce();
    expect(releaseLease).toHaveBeenCalledOnce();
    expect(disposeInvalidChunk).toHaveBeenCalledOnce();
  });

  it('admits only one produced/unacked chunk across concurrent operation hosts', async () => {
    const coordinator = new PullSessionHostCoordinator();
    let resolveFirst!: (chunk: { payload: Payload; byteLength: number; done: boolean }) => void;
    const firstChunk = new Promise<{ payload: Payload; byteLength: number; done: boolean }>((resolve) => {
      resolveFirst = resolve;
    });
    const first = new TestPullSessionHost({
      ...identity,
      operationId: 1,
      maxByteCredit: 8,
      coordinator,
      driver: { pull: () => firstChunk },
    });
    const secondPull = vi.fn(() => ({ payload: { value: 'two' }, byteLength: 1, done: false }));
    const second = new TestPullSessionHost({
      ...identity,
      operationId: 2,
      maxByteCredit: 8,
      coordinator,
      driver: { pull: secondPull },
    });

    const firstPending = first.handle({
      ...command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }), operationId: 1,
    });
    await Promise.resolve();
    const blockedSecond = second.handle({
      ...command('pull', { requestId: 2, sequence: 0, byteCredit: 1 }), operationId: 2,
    });
    expect(secondPull).not.toHaveBeenCalled();
    resolveFirst({ payload: { value: 'one' }, byteLength: 1, done: false });
    await firstPending;
    await expect(blockedSecond).resolves.toMatchObject({
      response: { kind: 'error', error: { message: expect.stringMatching(/another operation/) } },
    });
    expect(secondPull).not.toHaveBeenCalled();
    await first.handle({ ...command('ack', { requestId: 3, sequence: 0 }), operationId: 1 });
    await expect(
      second.handle({ ...command('pull', { requestId: 4, sequence: 0, byteCredit: 1 }), operationId: 2 }),
    ).resolves.toMatchObject({ response: { kind: 'chunk', payload: { value: 'two' } } });
  });

  it('releases the shared gate on driver failure but retains it across ACK failure until cancel', async () => {
    const coordinator = new PullSessionHostCoordinator();
    const failing = new TestPullSessionHost<Payload, string>({
      ...identity,
      operationId: 1,
      maxByteCredit: 8,
      coordinator,
      driver: { pull: () => { throw new Error('driver failed'); } },
    });
    const acknowledge = vi.fn(() => { throw new Error('ack failed'); });
    const holder = new TestPullSessionHost({
      ...identity,
      operationId: 2,
      maxByteCredit: 8,
      coordinator,
      driver: {
        pull: () => ({ payload: { value: 'held' }, byteLength: 1, done: false }),
        acknowledge,
      },
    });
    const waiting = new TestPullSessionHost({
      ...identity,
      operationId: 3,
      maxByteCredit: 8,
      coordinator,
      driver: { pull: () => ({ payload: { value: 'next' }, byteLength: 1, done: false }) },
    });

    await expect(
      failing.handle({ ...command('pull', { requestId: 1, sequence: 0, byteCredit: 1 }), operationId: 1 }),
    ).resolves.toMatchObject({ response: { kind: 'error', error: { message: 'driver failed' } } });
    await expect(
      holder.handle({ ...command('pull', { requestId: 2, sequence: 0, byteCredit: 1 }), operationId: 2 }),
    ).resolves.toMatchObject({ response: { kind: 'chunk' } });
    await expect(
      holder.handle({ ...command('ack', { requestId: 3, sequence: 0 }), operationId: 2 }),
    ).resolves.toMatchObject({ response: { kind: 'error', error: { message: 'ack failed' } } });
    await expect(
      waiting.handle({ ...command('pull', { requestId: 4, sequence: 0, byteCredit: 1 }), operationId: 3 }),
    ).resolves.toMatchObject({ response: { kind: 'error' } });
    await holder.handle({
      ...command('cancel', { requestId: 5, reason: 'request-error' }), operationId: 2,
    });
    await expect(
      waiting.handle({ ...command('pull', { requestId: 6, sequence: 0, byteCredit: 1 }), operationId: 3 }),
    ).resolves.toMatchObject({ response: { kind: 'chunk', payload: { value: 'next' } } });
  });
});

describe('client/host protocol parity', () => {
  it('accepts stale-generation lifecycle as a no-op without terminating the shared worker', async () => {
    const cancel = vi.fn();
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 8,
      driver: { pull: () => ({ payload: { value: 'current' }, byteLength: 1, done: false }), cancel },
    });
    const worker = new FakeWorker();
    worker.onPost = (posted) => {
      void host.handle(posted).then((post) => worker.respond(post.response));
    };
    const bridge = new WorkerBridge<Response>(worker, { correlate: (item) => item.requestId });
    const stale = new BoundedPullSession(bridge, {
      ...identity,
      generation: identity.generation - 1,
      maxByteCredit: 8,
      cancelGraceMs: 20,
    });
    await stale.cancel('abort');
    expect(worker.terminated).toBe(0);
    expect(cancel).not.toHaveBeenCalled();

    const mismatchedWorker = new FakeWorker();
    mismatchedWorker.onPost = (posted) => {
      void host.handle(posted).then((post) => mismatchedWorker.respond(post.response));
    };
    const mismatchedBridge = new WorkerBridge<Response>(mismatchedWorker, {
      correlate: (item) => item.requestId,
    });
    const mismatched = new BoundedPullSession(mismatchedBridge, {
      ...identity,
      sessionId: 'old-session',
      generation: identity.generation - 1,
      maxByteCredit: 8,
      cancelGraceMs: 20,
    });
    await mismatched.close();
    expect(mismatchedWorker.terminated).toBe(0);
    expect(cancel).not.toHaveBeenCalled();
  });

  it('round-trips pull, ACK, lease release, and final data through the same envelopes', async () => {
    const releaseLease = vi.fn();
    const chunks = [
      { payload: { value: 'a' }, byteLength: 1, done: false, leaseId: 21 },
      { payload: { value: 'z' }, byteLength: 1, done: true, leaseId: 22 },
    ];
    const host = new TestPullSessionHost({
      ...identity,
      maxByteCredit: 8,
      driver: {
        pull: () => chunks.shift() as (typeof chunks)[number],
        releaseLease,
        resourceUsage: () => usage,
      },
    });
    const worker = new FakeWorker();
    worker.onPost = (posted) => {
      void host.handle(posted).then((result) => worker.respond(result.response));
    };
    const bridge = new WorkerBridge<Response>(worker, { correlate: (item) => item.requestId });
    const client = new BoundedPullSession(bridge, { ...identity, maxByteCredit: 8 });

    const first = await client.pull(8);
    expect(first.usage).toEqual(usage);
    expect(client.usageCheckpoint).toEqual(usage);
    await first.ack();
    await first.release();
    const final = await client.pull(8);
    expect(final).toMatchObject({ payload: { value: 'z' }, done: true });
    await final.ack();
    await final.release();
    expect(releaseLease).toHaveBeenCalledWith(21);
    expect(releaseLease).toHaveBeenCalledWith(22);
  });

  it('preserves a typed resource-limit error across host serialization and client deserialization', async () => {
    const limitError = new OoxmlResourceLimitError('entry too large', {
      stage: 'decompression',
      violation: {
        format: 'docx',
        operation: 'parse',
        resource: 'archive-entry',
        part: 'word/document.xml',
        metric: 'actual-inflated-bytes',
        limit: 3,
        observed: 4,
        configurable: true,
        usage,
      },
    });
    const host = new TestPullSessionHost<Payload, string>({
      ...identity,
      maxByteCredit: 8,
      driver: {
        pull: () => {
          throw limitError;
        },
        resourceUsage: () => usage,
      },
    });
    const worker = new FakeWorker();
    worker.onPost = (posted) => {
      void host.handle(posted).then((post) => worker.respond(post.response));
    };
    const bridge = new WorkerBridge<Response>(worker, { correlate: (item) => item.requestId });
    const client = new BoundedPullSession(bridge, { ...identity, maxByteCredit: 8 });

    const rejection = client.pull(8).catch((error: unknown) => error);
    const error = await rejection;
    expect(error).toBeInstanceOf(OoxmlResourceLimitError);
    expect(error).toMatchObject({ code: 'ooxml-resource-limit' });
    expect((error as OoxmlResourceLimitError).details.violation).toMatchObject({
      part: 'word/document.xml',
      observed: 4,
    });
    expect(client.usageCheckpoint).toEqual(usage);
  });
});

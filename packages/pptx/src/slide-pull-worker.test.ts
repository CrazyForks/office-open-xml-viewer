import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import {
  BoundedPullSession,
  HARD_MAX_PPTX_SLIDE_JSON_BYTES,
  PULL_SESSION_PROTOCOL,
  PULL_SESSION_INSUFFICIENT_CREDIT_CODE,
  WorkerBridge,
  type PullSessionCommand,
  type PullSessionResponse,
  type WorkerLike,
} from '@silurus/ooxml-core/worker';
import { describe, expect, it, vi } from 'vitest';
import { SlidePullWorker } from './slide-pull-worker.js';

const identity = { sessionId: 4, operationId: 9, generation: 2 } as const;
const usage = {
  archiveEntryCount: 1,
  declaredInflatedBytes: 2,
  distinctInflatedBytes: 3,
  operationInflatedBytes: 4,
};

class LoopbackWorker implements WorkerLike {
  readonly posted: PullSessionCommand<number>[] = [];
  terminated = 0;
  onPost?: (command: PullSessionCommand<number>) => void;
  private readonly listeners = new Set<(event: MessageEvent) => void>();

  postMessage(message: unknown): void {
    const command = message as PullSessionCommand<number>;
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

  respond(response: PullSessionResponse<ArrayBuffer, number>): void {
    for (const listener of this.listeners) listener({ data: response } as MessageEvent);
  }
}

function bytes(value: unknown): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(value));
}

function insufficientCredit(requiredBytes: number, offeredBytes: number): Error {
  return new Error(`OOXML_INSUFFICIENT_CREDIT:${JSON.stringify({
    code: PULL_SESSION_INSUFFICIENT_CREDIT_CODE,
    requiredBytes,
    offeredBytes,
  })}`);
}

function makeArchive(payload = bytes({ index: 0, slideNumber: 1, background: null, elements: [] })) {
  return {
    pull_slide: vi.fn((
      _slideIndex: number,
      _operationId: number,
      _generation: number,
      _byteCredit: number,
    ) => payload),
    slide_cursor_resource_usage: vi.fn(() => bytes(usage)),
    acknowledge_slide: vi.fn(),
    cancel_slide: vi.fn<() => void | Promise<void>>(),
    close_presentation_session: vi.fn<() => void | Promise<void>>(),
  };
}

async function openWorker(
  worker: SlidePullWorker,
  slideIndex = 0,
  value: { sessionId: number; operationId: number; generation: number } = identity,
): Promise<void> {
  worker.reserveOpen(value);
  await worker.open(slideIndex, value);
}

function command(
  requestId: number,
  body:
    | { kind: 'pull'; sequence: number; byteCredit: number }
    | { kind: 'ack'; sequence: number }
    | { kind: 'release'; leaseId: number }
    | { kind: 'cancel'; reason: 'request-error' }
    | { kind: 'close' },
  value: { sessionId: number; operationId: number; generation: number } = identity,
): PullSessionCommand<number> {
  return { protocol: PULL_SESSION_PROTOCOL, requestId, ...value, ...body };
}

describe('SlidePullWorker', () => {
  it('pulls one indivisible slide and ACKs Rust only after consumer acceptance', async () => {
    const archive = makeArchive();
    const order: string[] = [];
    archive.acknowledge_slide.mockImplementation(() => { order.push('rust-ack'); });
    const accept = vi.fn((_slideIndex: number, _slide: unknown, _usage: unknown) => ({
      rollback: () => order.push('rollback'),
      commit: () => order.push('commit'),
    }));
    const worker = new SlidePullWorker(() => archive, (...args) => {
      order.push('accept');
      return accept(...args);
    });
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(worker, 3);

    await worker.dispatch(command(1, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), (response) => replies.push(response));
    expect(replies[0]).toMatchObject({ kind: 'chunk', done: true, sequence: 0 });
    expect(archive.pull_slide).toHaveBeenCalledWith(
      3, identity.operationId, identity.generation, HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    );
    expect(archive.acknowledge_slide).not.toHaveBeenCalled();

    await worker.dispatch(command(2, { kind: 'ack', sequence: 0 }), (response) => replies.push(response));
    expect(order).toEqual(['accept', 'rust-ack', 'commit']);
    expect(accept).toHaveBeenCalledWith(3, expect.objectContaining({ index: 0 }), usage);
  });

  it('retries the same Rust prepared unit after insufficient byte credit', async () => {
    const payload = bytes({ index: 0, data: 'retained' });
    let prepared = payload;
    const archive = makeArchive(payload);
    archive.pull_slide.mockImplementation((_index, _operation, _generation, credit) => {
      if (credit < prepared.byteLength) {
        throw insufficientCredit(prepared.byteLength, credit);
      }
      const result = prepared;
      prepared = new Uint8Array();
      return result;
    });
    const worker = new SlidePullWorker(() => archive);
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(worker);
    await worker.dispatch(command(1, { kind: 'pull', sequence: 0, byteCredit: 1 }),
      (response) => replies.push(response));
    expect(replies.at(-1)).toMatchObject({
      kind: 'error', error: { code: PULL_SESSION_INSUFFICIENT_CREDIT_CODE },
    });
    await worker.dispatch(command(2, { kind: 'pull', sequence: 0, byteCredit: payload.byteLength }),
      (response) => replies.push(response));
    expect(replies.at(-1)).toMatchObject({ kind: 'chunk', byteLength: payload.byteLength });
    expect(archive.pull_slide).toHaveBeenCalledTimes(2);
    await worker.dispatch(command(3, { kind: 'cancel', reason: 'request-error' }), () => undefined);
  });

  it('preserves the Rust retry contract through BoundedPullSession', async () => {
    const payload = bytes({ index: 0, data: 'indivisible' });
    const archive = makeArchive(payload);
    archive.pull_slide.mockImplementation((_index, _operation, _generation, credit) => {
      if (credit < payload.byteLength) {
        throw insufficientCredit(payload.byteLength, credit);
      }
      return payload;
    });
    const driver = new SlidePullWorker(() => archive);
    await openWorker(driver);
    const transport = new LoopbackWorker();
    const bridge = new WorkerBridge<PullSessionResponse<ArrayBuffer, number>>(transport, {
      correlate: (response) => response.requestId,
    });
    transport.onPost = (request) => {
      void driver.dispatch(request, (response) => transport.respond(response));
    };
    const client = new BoundedPullSession(bridge, {
      ...identity,
      maxByteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    });

    await expect(client.pull(1)).rejects.toMatchObject({
      name: 'PullSessionInsufficientCreditError', code: PULL_SESSION_INSUFFICIENT_CREDIT_CODE,
    });
    expect(transport.posted).toHaveLength(1);
    const chunk = await client.pull(payload.byteLength);
    expect(chunk).toMatchObject({ sequence: 0, byteLength: payload.byteLength, done: true });
    await chunk.ack();
    expect(archive.acknowledge_slide).toHaveBeenCalledOnce();
    expect(transport.posted.some((item) => item.kind === 'cancel')).toBe(false);
  });

  it('does not classify merely similar credit messages as recoverable', async () => {
    const archive = makeArchive();
    archive.pull_slide.mockImplementation(() => {
      throw new Error('slide unit requires 0002 bytes but credit is 1');
    });
    const worker = new SlidePullWorker(() => archive);
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(worker);
    await worker.dispatch(command(1, { kind: 'pull', sequence: 0, byteCredit: 1 }),
      (response) => replies.push(response));
    expect(replies[0]).toMatchObject({ kind: 'error' });
    if (replies[0]?.kind !== 'error') throw new Error('expected error');
    expect(replies[0].error.code).not.toBe(PULL_SESSION_INSUFFICIENT_CREDIT_CODE);
    await worker.dispatch(command(2, { kind: 'cancel', reason: 'request-error' }), () => undefined);
  });

  it('cancels the provisional Rust slide and releases the package FIFO', async () => {
    const archive = makeArchive();
    const worker = new SlidePullWorker(() => archive);
    await openWorker(worker);
    await worker.dispatch(command(1, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), () => undefined);
    const sibling = vi.fn();
    const queued = worker.run(sibling);
    await Promise.resolve();
    expect(sibling).not.toHaveBeenCalled();
    await worker.dispatch(command(2, { kind: 'cancel', reason: 'request-error' }), () => undefined);
    await queued;
    expect(archive.cancel_slide).toHaveBeenCalledOnce();
    expect(sibling).toHaveBeenCalledOnce();
  });

  it('keeps Rust provisional when acceptance fails before ACK', async () => {
    const archive = makeArchive();
    const worker = new SlidePullWorker(() => archive, () => { throw new Error('cache rejected slide'); });
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(worker);
    await worker.dispatch(command(1, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), (response) => replies.push(response));
    await worker.dispatch(command(2, { kind: 'ack', sequence: 0 }),
      (response) => replies.push(response));
    expect(replies.at(-1)).toMatchObject({ kind: 'error' });
    expect(archive.acknowledge_slide).not.toHaveBeenCalled();
    await worker.dispatch(command(3, { kind: 'cancel', reason: 'request-error' }), () => undefined);
    expect(archive.cancel_slide).toHaveBeenCalledOnce();
  });

  it('rolls back accepted consumer state when the Rust ACK fails', async () => {
    const archive = makeArchive();
    archive.acknowledge_slide.mockImplementation(() => { throw new Error('ack failed'); });
    let retained = false;
    const worker = new SlidePullWorker(() => archive, () => {
      retained = true;
      return { rollback: () => { retained = false; } };
    });
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(worker);
    await worker.dispatch(command(1, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), (response) => replies.push(response));
    await worker.dispatch(command(2, { kind: 'ack', sequence: 0 }),
      (response) => replies.push(response));
    expect(retained).toBe(false);
    expect(replies.at(-1)).toMatchObject({ kind: 'error' });
    await worker.dispatch(command(3, { kind: 'cancel', reason: 'request-error' }), () => undefined);
  });

  it('uses an exact ArrayBuffer span and a transfer list', async () => {
    const backing = new Uint8Array(64);
    const encoded = bytes({ index: 0 });
    backing.set(encoded, 13);
    const view = backing.subarray(13, 13 + encoded.byteLength);
    const archive = makeArchive(view);
    const worker = new SlidePullWorker(() => archive);
    let response: PullSessionResponse<ArrayBuffer, number> | undefined;
    let transfer: Transferable[] | undefined;
    await openWorker(worker);
    await worker.dispatch(command(1, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), (value, moved) => { response = value; transfer = moved; });
    expect(response).toMatchObject({ kind: 'chunk', byteLength: encoded.byteLength });
    if (!response || response.kind !== 'chunk') throw new Error('expected chunk');
    expect(response.payload.byteLength).toBe(encoded.byteLength);
    expect(transfer).toEqual([response.payload]);
    await worker.dispatch(command(2, { kind: 'cancel', reason: 'request-error' }), () => undefined);
  });

  it('transfers and detaches the fast-path ArrayBuffer instead of cloning it', async () => {
    const payload = bytes({ index: 0, data: 'transfer-owned' });
    const archive = makeArchive(payload);
    const worker = new SlidePullWorker(() => archive);
    await openWorker(worker);
    await worker.dispatch(command(1, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), (response, transfer) => {
      if (response.kind !== 'chunk' || !transfer) throw new Error('expected transferable chunk');
      const cloned = structuredClone(response, { transfer });
      expect(cloned.payload.byteLength).toBeGreaterThan(0);
    });
    expect(payload.buffer.byteLength).toBe(0);
    await worker.dispatch(command(2, { kind: 'cancel', reason: 'request-error' }), () => undefined);
  });

  it('contains a failed data post and converges package cleanup', async () => {
    const archive = makeArchive();
    const worker = new SlidePullWorker(() => archive);
    const fallback: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(worker);
    await expect(worker.dispatchSafely(command(1, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), (response) => {
      if (response.kind === 'chunk') throw new Error('postMessage failed');
      fallback.push(response);
    })).resolves.toBeUndefined();
    expect(archive.cancel_slide).toHaveBeenCalledOnce();
    expect(fallback[0]).toMatchObject({ kind: 'error' });
    await expect(worker.run(() => 'ready')).resolves.toBe('ready');
  });

  it('honors cancel while open is queued and rejects stale pending identity', async () => {
    const archive = makeArchive();
    const worker = new SlidePullWorker(() => archive);
    let release!: () => void;
    const blocker = worker.run(() => new Promise<void>((resolve) => { release = resolve; }));
    await Promise.resolve();
    worker.reserveOpen(identity);
    const opening = worker.open(0, identity);
    const stale = { ...identity, generation: identity.generation + 1 };
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await worker.dispatch(command(1, { kind: 'close' }, stale), (response) => replies.push(response));
    expect(replies[0]).toMatchObject({ kind: 'error', error: { code: 'ooxml-stale-lifecycle' } });
    await worker.dispatch(command(2, { kind: 'cancel', reason: 'request-error' }),
      (response) => replies.push(response));
    expect(replies[1]).toMatchObject({ kind: 'accepted', command: 'cancel' });
    release();
    await blocker;
    await expect(opening).rejects.toThrow('open was canceled');
    expect(archive.pull_slide).not.toHaveBeenCalled();
  });

  it('closes an opened operation when both opened-response posts fail', async () => {
    const archive = makeArchive();
    const worker = new SlidePullWorker(() => archive);
    await openWorker(worker);
    await expect(worker.postOpenedSafely(identity,
      () => { throw new Error('opened post failed'); },
      () => { throw new Error('fallback post failed'); },
    )).resolves.toBeUndefined();
    expect(archive.cancel_slide).toHaveBeenCalledOnce();
    await expect(worker.run(() => 'ready')).resolves.toBe('ready');
  });

  it('latches the first resource failure and replays it to sibling package work', async () => {
    const first = new OoxmlResourceLimitError('slide JSON limit', {
      stage: 'serialization',
      violation: {
        format: 'pptx', operation: 'slide-cursor', resource: 'slide-json', metric: 'bytes',
        limit: 10, observed: 11, configurable: false, usage,
      },
    });
    const archive = makeArchive();
    archive.pull_slide.mockImplementation(() => { throw first; });
    const worker = new SlidePullWorker(() => archive);
    const poisonedGeneration = worker.coordinator;
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(worker);
    await worker.dispatch(command(1, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), (response) => replies.push(response));
    expect(replies[0]).toMatchObject({ kind: 'error', error: { code: 'ooxml-resource-limit' } });
    expect(poisonedGeneration.fatalError).toMatchObject({ code: 'ooxml-resource-limit' });
    const sibling = vi.fn();
    await expect(worker.run(sibling)).rejects.toBe(first);
    expect(sibling).not.toHaveBeenCalled();

    await worker.reset();
    expect(worker.coordinator).not.toBe(poisonedGeneration);
    expect(worker.coordinator.fatalError).toBeUndefined();
    archive.pull_slide.mockImplementation((_index, _operation, _generation, _credit) =>
      bytes({ index: 0, slideNumber: 1, background: null, elements: [] }));
    const fresh = { sessionId: 8, operationId: 10, generation: 3 };
    await openWorker(worker, 0, fresh);
    const freshReplies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await worker.dispatch(command(2, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }, fresh), (response) => freshReplies.push(response));
    expect(freshReplies.at(-1)).toMatchObject({ kind: 'chunk' });
    await worker.dispatch(command(3, { kind: 'ack', sequence: 0 }, fresh),
      (response) => freshReplies.push(response));
    expect(freshReplies.at(-1)).toMatchObject({ kind: 'accepted', command: 'ack' });
  });

  it('latches and replays a raw Rust resource-limit envelope', async () => {
    const raw = `OOXML_RESOURCE_LIMIT:${JSON.stringify({
      code: 'ooxml-resource-limit',
      details: {
        stage: 'serialization',
        violation: {
          format: 'pptx', operation: 'slide-cursor', resource: 'slide-json', metric: 'bytes',
          limit: 10, observed: 11, configurable: false, usage,
        },
      },
    })}`;
    const archive = makeArchive();
    archive.pull_slide.mockImplementation(() => { throw new Error(raw); });
    const worker = new SlidePullWorker(() => archive);
    await openWorker(worker);
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await worker.dispatch(command(1, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), (response) => replies.push(response));
    expect(replies[0]).toMatchObject({ kind: 'error', error: { code: 'ooxml-resource-limit' } });
    await expect(worker.run(() => 'must not run')).rejects.toBeInstanceOf(OoxmlResourceLimitError);
  });

  it('keeps Rust unACKed when usage bytes are malformed and remains cancelable', async () => {
    const archive = makeArchive();
    archive.slide_cursor_resource_usage.mockImplementation(() => bytes({ archiveEntryCount: -1 }));
    const worker = new SlidePullWorker(() => archive);
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(worker);
    await worker.dispatch(command(1, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), (response) => replies.push(response));
    expect(replies[0]).toMatchObject({ kind: 'error' });
    expect(archive.acknowledge_slide).not.toHaveBeenCalled();
    await worker.dispatch(command(2, { kind: 'cancel', reason: 'request-error' }),
      (response) => replies.push(response));
    expect(archive.cancel_slide).toHaveBeenCalledOnce();
    await expect(worker.run(() => 'clean')).resolves.toBe('clean');
  });

  it('reset closes pending/active operations, presentation state, and clears poison', async () => {
    const archive = makeArchive();
    const worker = new SlidePullWorker(() => archive);
    await expect(worker.run(() => {
      throw new OoxmlResourceLimitError('limit', {
        stage: 'worker',
        violation: {
          format: 'pptx', operation: 'worker', resource: 'worker-model', metric: 'bytes',
          limit: 1, observed: 2, configurable: false, usage,
        },
      });
    })).rejects.toBeInstanceOf(OoxmlResourceLimitError);
    await worker.reset();
    expect(archive.close_presentation_session).toHaveBeenCalledOnce();
    await expect(worker.run(() => 'fresh')).resolves.toBe('fresh');

    await openWorker(worker);
    await worker.reset();
    expect(archive.cancel_slide).toHaveBeenCalledOnce();
    expect(archive.close_presentation_session).toHaveBeenCalledTimes(2);

    const pending = { ...identity, sessionId: 5 };
    worker.reserveOpen(pending);
    await worker.reset();
    expect(worker.pendingOpenCount).toBe(0);
    expect(archive.close_presentation_session).toHaveBeenCalledTimes(3);
    await expect(worker.open(0, pending)).rejects.toThrow('reservation is stale or missing');
  });

  it('erects a synchronous generation barrier while an active close is delayed', async () => {
    const archive = makeArchive();
    let releaseClose!: () => void;
    archive.cancel_slide.mockImplementation(() => new Promise<void>((resolve) => {
      releaseClose = resolve;
    }));
    const worker = new SlidePullWorker(() => archive);
    await openWorker(worker);
    const alreadyReserved = { sessionId: 5, operationId: 10, generation: 2 };
    worker.reserveOpen(alreadyReserved);

    const resetting = worker.reset();
    expect(() => worker.reserveOpen({ sessionId: 6, operationId: 11, generation: 3 }))
      .toThrow('reset is in progress');
    await expect(worker.open(1, alreadyReserved)).rejects.toThrow('reset is in progress');
    const escaped = vi.fn();
    await expect(worker.run(escaped)).rejects.toThrow('reset is in progress');
    expect(escaped).not.toHaveBeenCalled();
    expect(archive.close_presentation_session).not.toHaveBeenCalled();

    releaseClose();
    await resetting;
    expect(worker.pendingOpenCount).toBe(0);
    expect(archive.cancel_slide).toHaveBeenCalledOnce();
    expect(archive.close_presentation_session).toHaveBeenCalledOnce();

    archive.cancel_slide.mockImplementation(() => undefined);
    const fresh = { sessionId: 7, operationId: 12, generation: 3 };
    await openWorker(worker, 0, fresh);
    await worker.reset();
  });

  it('does not forward late commands during or after reset to either archive generation', async () => {
    const oldArchive = makeArchive();
    let releaseClose!: () => void;
    oldArchive.cancel_slide.mockImplementation(() => new Promise<void>((resolve) => {
      releaseClose = resolve;
    }));
    const freshArchive = makeArchive();
    let currentArchive = oldArchive;
    const worker = new SlidePullWorker(() => currentArchive);
    await openWorker(worker);

    const resetting = worker.reset();
    await Promise.resolve();
    await Promise.resolve();
    expect(oldArchive.cancel_slide).toHaveBeenCalledOnce();
    const during: PullSessionResponse<ArrayBuffer, number>[] = [];
    await worker.dispatch(command(10, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), (response) => during.push(response));
    await worker.dispatch(command(11, { kind: 'ack', sequence: 0 }),
      (response) => during.push(response));
    await worker.dispatch(command(12, { kind: 'release', leaseId: 1 }),
      (response) => during.push(response));
    await worker.dispatch(command(13, { kind: 'cancel', reason: 'request-error' }),
      (response) => during.push(response));
    await worker.dispatch(command(14, { kind: 'close' }, { ...identity, generation: 3 }),
      (response) => during.push(response));
    expect(during.slice(0, 3)).toEqual([
      expect.objectContaining({ kind: 'error', error: expect.objectContaining({ code: 'ooxml-pull-resetting' }) }),
      expect.objectContaining({ kind: 'error', error: expect.objectContaining({ code: 'ooxml-pull-resetting' }) }),
      expect.objectContaining({ kind: 'error', error: expect.objectContaining({ code: 'ooxml-pull-resetting' }) }),
    ]);
    expect(during[3]).toMatchObject({ kind: 'accepted', command: 'cancel' });
    expect(during[4]).toMatchObject({ kind: 'error', error: { code: 'ooxml-stale-lifecycle' } });
    expect(oldArchive.pull_slide).not.toHaveBeenCalled();
    expect(oldArchive.acknowledge_slide).not.toHaveBeenCalled();
    expect(oldArchive.cancel_slide).toHaveBeenCalledOnce();

    releaseClose();
    await resetting;
    currentArchive = freshArchive;
    const after: PullSessionResponse<ArrayBuffer, number>[] = [];
    await worker.dispatch(command(20, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), (response) => after.push(response));
    await worker.dispatch(command(21, { kind: 'ack', sequence: 0 }),
      (response) => after.push(response));
    await worker.dispatch(command(22, { kind: 'cancel', reason: 'request-error' }),
      (response) => after.push(response));
    expect(after[0]).toMatchObject({ kind: 'error' });
    expect(after[1]).toMatchObject({ kind: 'error' });
    expect(after[2]).toMatchObject({ kind: 'accepted', command: 'cancel' });
    expect(freshArchive.pull_slide).not.toHaveBeenCalled();
    expect(freshArchive.acknowledge_slide).not.toHaveBeenCalled();
    expect(freshArchive.cancel_slide).not.toHaveBeenCalled();
  });

  it('coalesces concurrent reset calls onto one teardown barrier', async () => {
    const archive = makeArchive();
    let releaseClose!: () => void;
    archive.cancel_slide.mockImplementation(() => new Promise<void>((resolve) => {
      releaseClose = resolve;
    }));
    const worker = new SlidePullWorker(() => archive);
    await openWorker(worker);

    const first = worker.reset();
    const second = worker.reset();
    expect(second).toBe(first);
    await Promise.resolve();
    await Promise.resolve();
    expect(archive.cancel_slide).toHaveBeenCalledOnce();
    releaseClose();
    await Promise.all([first, second]);
    expect(archive.cancel_slide).toHaveBeenCalledOnce();
    expect(archive.close_presentation_session).toHaveBeenCalledOnce();
  });

  it('keeps a failed reset closed until a retry completes successfully', async () => {
    const archive = makeArchive();
    archive.close_presentation_session
      .mockImplementationOnce(() => { throw new Error('close presentation failed'); })
      .mockImplementation(() => undefined);
    const worker = new SlidePullWorker(() => archive);
    const poisonedGeneration = worker.coordinator;
    const fatal = new OoxmlResourceLimitError('worker projection limit', {
      stage: 'worker',
      violation: {
        format: 'pptx', operation: 'worker', resource: 'worker-model', metric: 'bytes',
        limit: 1, observed: 2, configurable: false, usage,
      },
    });
    await expect(worker.run(() => { throw fatal; })).rejects.toBe(fatal);
    await expect(worker.reset()).rejects.toThrow('close presentation failed');
    expect(worker.coordinator).toBe(poisonedGeneration);

    expect(() => worker.reserveOpen(identity)).toThrow('reset failed');
    await expect(worker.open(0, identity)).rejects.toThrow('reset failed');
    const ordinary = vi.fn();
    await expect(worker.run(ordinary)).rejects.toThrow('reset failed');
    expect(ordinary).not.toHaveBeenCalled();
    const failedResponses: PullSessionResponse<ArrayBuffer, number>[] = [];
    await worker.dispatch(command(1, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }), (response) => failedResponses.push(response));
    await worker.dispatch(command(2, { kind: 'close' }),
      (response) => failedResponses.push(response));
    expect(failedResponses[0]).toMatchObject({
      kind: 'error', error: { code: 'ooxml-pull-reset-failed' },
    });
    expect(failedResponses[1]).toMatchObject({ kind: 'accepted', command: 'close' });
    expect(archive.pull_slide).not.toHaveBeenCalled();
    expect(archive.cancel_slide).not.toHaveBeenCalled();

    await worker.reset();
    expect(worker.coordinator).not.toBe(poisonedGeneration);
    expect(archive.close_presentation_session).toHaveBeenCalledTimes(2);
    await expect(worker.run(() => 'fresh')).resolves.toBe('fresh');
    await openWorker(worker);
    await worker.reset();
  });

  it('rejects non-positive identities synchronously and is not publicly exported', async () => {
    const worker = new SlidePullWorker(() => makeArchive());
    expect(() => worker.reserveOpen({ ...identity, operationId: 0 })).toThrow('positive');
    const publicModule = await import('./index.js');
    expect('SlidePullWorker' in publicModule).toBe(false);
  });
});

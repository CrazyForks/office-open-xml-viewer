import { describe, it, expect } from 'vitest';
import { WorkerBridge, type WorkerLike } from './bridge.js';

/** Minimal in-memory Worker stand-in. Records posted messages and lets a test
 *  deliver responses synchronously via {@link respond}. */
class FakeWorker implements WorkerLike {
  posted: unknown[] = [];
  transfers: (Transferable[] | undefined)[] = [];
  terminated = false;
  private listeners = new Set<(e: MessageEvent) => void>();

  postMessage(message: unknown, transfer?: Transferable[]): void {
    this.posted.push(message);
    this.transfers.push(transfer);
  }
  addEventListener(_type: 'message', listener: (e: MessageEvent) => void): void {
    this.listeners.add(listener);
  }
  removeEventListener(_type: 'message', listener: (e: MessageEvent) => void): void {
    this.listeners.delete(listener);
  }
  terminate(): void {
    this.terminated = true;
  }
  /** Deliver a message to all listeners, like the worker posting back. */
  respond(data: unknown): void {
    for (const l of this.listeners) l({ data } as MessageEvent);
  }
}

interface Res {
  id?: number;
  kind: 'ready' | 'ok' | 'error';
  value?: string;
  message?: string;
}

function makeBridge(worker: FakeWorker, onUnsolicited?: (r: Res) => void) {
  return new WorkerBridge<Res>(worker, {
    correlate: (r) => r.id,
    toError: (r) => (r.kind === 'error' ? (r.message ?? 'error') : undefined),
    onUnsolicited,
  });
}

describe('WorkerBridge', () => {
  it('correlates a response to its request by id', async () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const p = bridge.request((id) => ({ kind: 'parse', id }));
    const sentId = (w.posted[0] as { id: number }).id;
    w.respond({ id: sentId, kind: 'ok', value: 'A' });
    await expect(p).resolves.toMatchObject({ value: 'A' });
  });

  it('routes concurrent responses to the matching request, even out of order', async () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const p1 = bridge.request((id) => ({ kind: 'parse', id }));
    const p2 = bridge.request((id) => ({ kind: 'parse', id }));
    const id1 = (w.posted[0] as { id: number }).id;
    const id2 = (w.posted[1] as { id: number }).id;
    expect(id1).not.toBe(id2);
    // Respond to the second request first.
    w.respond({ id: id2, kind: 'ok', value: 'second' });
    w.respond({ id: id1, kind: 'ok', value: 'first' });
    await expect(p1).resolves.toMatchObject({ value: 'first' });
    await expect(p2).resolves.toMatchObject({ value: 'second' });
  });

  it('rejects only the matching request on an error response', async () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const p1 = bridge.request((id) => ({ kind: 'parse', id }));
    const p2 = bridge.request((id) => ({ kind: 'parse', id }));
    const id1 = (w.posted[0] as { id: number }).id;
    const id2 = (w.posted[1] as { id: number }).id;
    w.respond({ id: id1, kind: 'error', message: 'boom' });
    w.respond({ id: id2, kind: 'ok', value: 'fine' });
    await expect(p1).rejects.toThrow('boom');
    await expect(p2).resolves.toMatchObject({ value: 'fine' });
  });

  it('does not resolve or hang the wrong request when an unknown id arrives', async () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const p = bridge.request((id) => ({ kind: 'parse', id }));
    const id = (w.posted[0] as { id: number }).id;
    w.respond({ id: 9999, kind: 'ok', value: 'stray' }); // unknown id: ignored
    w.respond({ id, kind: 'ok', value: 'real' });
    await expect(p).resolves.toMatchObject({ value: 'real' });
  });

  it('forwards unsolicited messages (no id) to onUnsolicited instead of a pending request', async () => {
    const w = new FakeWorker();
    const seen: Res[] = [];
    const bridge = makeBridge(w, (r) => seen.push(r));
    const p = bridge.request((id) => ({ kind: 'parse', id }));
    w.respond({ kind: 'ready' }); // no id
    const id = (w.posted[0] as { id: number }).id;
    w.respond({ id, kind: 'ok', value: 'done' });
    await expect(p).resolves.toMatchObject({ value: 'done' });
    expect(seen).toEqual([{ kind: 'ready' }]);
  });

  it('rejects all pending requests when terminated', async () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const p = bridge.request((id) => ({ kind: 'parse', id }));
    bridge.terminate();
    expect(w.terminated).toBe(true);
    await expect(p).rejects.toThrow(/terminated/i);
  });

  it('passes the transfer list through to postMessage', () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const buf = new ArrayBuffer(8);
    bridge.request((id) => ({ kind: 'parse', id, buffer: buf }), [buf]);
    expect(w.transfers[0]).toEqual([buf]);
  });

  it('post() sends a fire-and-forget message without allocating an id', () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    bridge.post({ kind: 'init', wasmUrl: 'x' });
    expect(w.posted[0]).toEqual({ kind: 'init', wasmUrl: 'x' });
    // The next request should still start from id 1.
    bridge.request((id) => ({ kind: 'parse', id }));
    expect((w.posted[1] as { id: number }).id).toBe(1);
  });

  it('ignores a duplicate/late response after the request already settled', async () => {
    const w = new FakeWorker();
    const bridge = makeBridge(w);
    const p = bridge.request((id) => ({ kind: 'parse', id }));
    const id = (w.posted[0] as { id: number }).id;
    w.respond({ id, kind: 'ok', value: 'first' });
    await expect(p).resolves.toMatchObject({ value: 'first' });
    // A stray duplicate must not throw or affect anything.
    expect(() => w.respond({ id, kind: 'ok', value: 'dup' })).not.toThrow();
  });
});

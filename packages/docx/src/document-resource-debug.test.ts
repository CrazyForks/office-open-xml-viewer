import { afterEach, describe, expect, it, vi } from 'vitest';
import type { OoxmlResourceUsageSnapshot, WorkerLike } from '@silurus/ooxml-core';
import { DocxDocument } from './document';

class SilentWorker implements WorkerLike {
  static instances: SilentWorker[] = [];
  terminated = false;

  constructor() {
    SilentWorker.instances.push(this);
  }

  postMessage(): void {}
  addEventListener(): void {}
  removeEventListener(): void {}
  terminate(): void { this.terminated = true; }
}

const globals = globalThis as Record<string, unknown>;
const originals = {
  Worker: globals.Worker,
  location: globals.location,
};

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
  globals.Worker = originals.Worker;
  globals.location = originals.location;
  SilentWorker.instances = [];
});

function installLoadHarness(): void {
  globals.Worker = SilentWorker;
  globals.location = { href: 'http://localhost/' };
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(
    DocxDocument.prototype as unknown as {
      _parse(
        buffer: ArrayBuffer,
        resourcePolicy: object,
        useGoogleFonts?: boolean,
        timeoutMs?: number,
      ): Promise<void>;
    },
    '_parse',
  ).mockResolvedValue(undefined);
}

describe('DocxDocument resource diagnostics', () => {
  it('uses workerTimeoutMs and ignores a rejected final usage probe', async () => {
    installLoadHarness();
    const probe = vi.spyOn(
      DocxDocument.prototype as unknown as {
        _resourceUsage(timeoutMs: number): Promise<OoxmlResourceUsageSnapshot>;
      },
      '_resourceUsage',
    ).mockRejectedValue(new Error('diagnostic probe failed'));

    const document = await DocxDocument.load(new ArrayBuffer(0), {
      debug: true,
      workerTimeoutMs: 17,
    });

    expect(probe).toHaveBeenCalledWith(17);
    expect(SilentWorker.instances[0]?.terminated).toBe(false);
    document.destroy();
  });

  it('bounds a no-response final usage probe without changing load success', async () => {
    vi.useFakeTimers();
    installLoadHarness();

    const pending = DocxDocument.load(new ArrayBuffer(0), {
      debug: true,
      workerTimeoutMs: 25,
    });
    await vi.advanceTimersByTimeAsync(24);
    let settled = false;
    void pending.then(() => { settled = true; });
    await Promise.resolve();
    expect(settled).toBe(false);

    await vi.advanceTimersByTimeAsync(1);
    const document = await pending;
    expect(SilentWorker.instances[0]?.terminated).toBe(false);
    document.destroy();
  });

  it('uses a finite diagnostic timeout when workerTimeoutMs is omitted', async () => {
    vi.useFakeTimers();
    installLoadHarness();

    const pending = DocxDocument.load(new ArrayBuffer(0), { debug: true });
    await vi.advanceTimersByTimeAsync(999);
    let settled = false;
    void pending.then(() => { settled = true; });
    await Promise.resolve();
    expect(settled).toBe(false);

    await vi.advanceTimersByTimeAsync(1);
    const document = await pending;
    expect(SilentWorker.instances[0]?.terminated).toBe(false);
    document.destroy();
  });
});

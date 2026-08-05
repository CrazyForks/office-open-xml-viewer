import { describe, expect, it, vi } from 'vitest';
import { WasmTrapError } from '@silurus/ooxml-core/worker';
import { WasmRuntimeGenerationHost } from '@silurus/ooxml-core/internal/wasm-runtime-generation';

type Archive = { value(): number; free(): void };

function runtimeHost() {
  const initSync = vi.fn();
  const reinit = vi.fn(async () => undefined);
  const host = new WasmRuntimeGenerationHost<Archive>(
    { initSync, reinit },
    new WebAssembly.Module(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0])),
  );
  return { host, initSync, reinit };
}

describe('WasmRuntimeGenerationHost', () => {
  it('initializes once and permits concurrent archive handles in one generation', async () => {
    const { host, initSync, reinit } = runtimeHost();
    const first = await host.open(() => ({ value: () => 1, free: vi.fn() }));
    const second = await host.open(() => ({ value: () => 2, free: vi.fn() }));

    expect(first.proxy.value()).toBe(1);
    expect(second.proxy.value()).toBe(2);
    expect(initSync).toHaveBeenCalledOnce();
    expect(reinit).not.toHaveBeenCalled();
  });

  it('poisons every live handle and reinitializes once for concurrent next opens', async () => {
    const { host, reinit } = runtimeHost();
    const trapped = await host.open(() => ({
      value: () => { throw new WebAssembly.RuntimeError('unreachable'); },
      free: vi.fn(),
    }));
    const siblingFree = vi.fn();
    const sibling = await host.open(() => ({ value: () => 2, free: siblingFree }));

    expect(() => trapped.proxy.value()).toThrow(WasmTrapError);
    expect(() => sibling.proxy.value()).toThrow(WasmTrapError);
    sibling.close((archive) => archive.free());
    expect(siblingFree).not.toHaveBeenCalled();

    const [next, concurrent] = await Promise.all([
      host.open(() => ({ value: () => 3, free: vi.fn() })),
      host.open(() => ({ value: () => 4, free: vi.fn() })),
    ]);
    expect(reinit).toHaveBeenCalledOnce();
    expect(next.proxy.value()).toBe(3);
    expect(concurrent.proxy.value()).toBe(4);
  });

  it('frees a healthy archive exactly once', async () => {
    const { host } = runtimeHost();
    const free = vi.fn();
    const handle = await host.open(() => ({ value: () => 1, free }));

    handle.close((archive) => archive.free());
    handle.close((archive) => archive.free());
    expect(free).toHaveBeenCalledOnce();
    expect(() => handle.proxy.value()).toThrow(/discarded runtime generation/);
  });
});

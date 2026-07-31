import { describe, expect, it, vi } from 'vitest';
import {
  FindPopupController,
  type FindableViewer,
  type FindPopupElements,
} from './findPopup.js';

class FakeTarget {
  private listeners = new Map<string, Array<(event: Event) => void>>();

  addEventListener(type: string, listener: EventListenerOrEventListenerObject): void {
    const fn = typeof listener === 'function' ? listener : listener.handleEvent.bind(listener);
    this.listeners.set(type, [...(this.listeners.get(type) ?? []), fn]);
  }

  removeEventListener(type: string, listener: EventListenerOrEventListenerObject): void {
    const current = this.listeners.get(type) ?? [];
    this.listeners.set(type, current.filter((fn) => fn !== listener));
  }

  dispatch(type: string, event: Event): void {
    for (const listener of this.listeners.get(type) ?? []) listener(event);
  }
}

function button(): HTMLButtonElement & FakeTarget {
  return Object.assign(new FakeTarget(), { disabled: false }) as HTMLButtonElement & FakeTarget;
}

function setup() {
  const keyTarget = new FakeTarget();
  const input = Object.assign(new FakeTarget(), {
    value: '',
    focus: vi.fn(),
    select: vi.fn(),
  }) as unknown as HTMLInputElement & FakeTarget;
  const elements = {
    root: { hidden: true },
    input,
    status: { textContent: '' },
    previous: button(),
    next: button(),
    close: button(),
  } as unknown as FindPopupElements;
  const viewer: FindableViewer = {
    findText: vi.fn(async () => [{ matchIndex: 0 }, { matchIndex: 1 }]),
    findNext: vi
      .fn()
      .mockResolvedValueOnce({ matchIndex: 0 })
      .mockResolvedValueOnce({ matchIndex: 1 }),
    findPrev: vi.fn(async () => ({ matchIndex: 0 })),
    clearFind: vi.fn(),
  };
  const controller = new FindPopupController(elements, keyTarget);
  controller.setViewer(viewer);
  return { controller, elements, input, keyTarget, viewer };
}

function key(key: string, init: Partial<KeyboardEvent> = {}): KeyboardEvent {
  return {
    key,
    ctrlKey: false,
    metaKey: false,
    altKey: false,
    shiftKey: false,
    preventDefault: vi.fn(),
    stopPropagation: vi.fn(),
    ...init,
  } as unknown as KeyboardEvent;
}

async function flushController(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
  await Promise.resolve();
}

describe('VS Code find popup', () => {
  it.each([
    { ctrlKey: true },
    { metaKey: true },
  ])('opens for Ctrl/Cmd+F and focuses the query', (modifier) => {
    const { elements, input, keyTarget } = setup();
    const event = key('f', modifier);

    keyTarget.dispatch('keydown', event);

    expect(elements.root.hidden).toBe(false);
    expect(input.focus).toHaveBeenCalledOnce();
    expect(input.select).toHaveBeenCalledOnce();
    expect(event.preventDefault).toHaveBeenCalledOnce();
  });

  it('searches case-insensitively and Enter/Shift+Enter navigate', async () => {
    const { controller, elements, input, keyTarget, viewer } = setup();
    controller.open();
    input.value = 'ALPHA';

    await controller.search();
    expect(viewer.findText).toHaveBeenCalledWith('ALPHA', { caseSensitive: false });
    expect(elements.status.textContent).toBe('1 of 2');

    keyTarget.dispatch('keydown', key('Enter'));
    await flushController();
    expect(viewer.findNext).toHaveBeenCalledTimes(2);

    keyTarget.dispatch('keydown', key('Enter', { shiftKey: true }));
    await flushController();
    expect(viewer.findPrev).toHaveBeenCalledOnce();
  });

  it.each(['previous', 'next', 'close'] as const)(
    'leaves Enter on the %s button to native button activation',
    async (buttonName) => {
      const { controller, elements, input, keyTarget, viewer } = setup();
      controller.open();
      input.value = 'ALPHA';
      await controller.search();
      const event = key('Enter', { target: elements[buttonName] });

      keyTarget.dispatch('keydown', event);
      await flushController();

      expect(event.preventDefault).not.toHaveBeenCalled();
      expect(event.stopPropagation).not.toHaveBeenCalled();
      expect(viewer.findNext).toHaveBeenCalledTimes(1);
      expect(viewer.findPrev).not.toHaveBeenCalled();
      expect(elements.root.hidden).toBe(false);
    },
  );

  it.each(['Escape', 'close'])('closes with %s and clears highlights', (action) => {
    const { controller, elements, keyTarget, viewer } = setup();
    controller.open();

    if (action === 'Escape') keyTarget.dispatch('keydown', key('Escape'));
    else (elements.close as unknown as FakeTarget).dispatch('click', {} as Event);

    expect(elements.root.hidden).toBe(true);
    expect(viewer.clearFind).toHaveBeenCalled();
  });

  it('reports an empty result set and disables navigation', async () => {
    const { controller, elements, input, viewer } = setup();
    vi.mocked(viewer.findText).mockResolvedValueOnce([]);
    controller.open();
    input.value = 'missing';

    await controller.search();

    expect(elements.status.textContent).toBe('No results');
    expect(elements.previous.disabled).toBe(true);
    expect(elements.next.disabled).toBe(true);
  });
});

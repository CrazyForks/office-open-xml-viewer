import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  BROWSER_CONSOLE_TUI_STYLE,
  emitConsoleTui,
} from './console-tui.js';

describe('emitConsoleTui', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it('pins fixed-width typography in browser consoles without setting colors', () => {
    const consoleLog = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    vi.stubGlobal('window', {});
    vi.stubGlobal('process', undefined);

    emitConsoleTui('┌─┐\n│A│\n└─┘');

    expect(consoleLog).toHaveBeenCalledWith(
      '%c┌─┐\n│A│\n└─┘',
      BROWSER_CONSOLE_TUI_STYLE,
    );
    expect(BROWSER_CONSOLE_TUI_STYLE).toBe([
      'font-family: "SFMono-Regular", Menlo, Monaco, Consolas, monospace',
      'font-size: 12px',
      'line-height: 1.35',
      'font-variant-ligatures: none',
      'letter-spacing: 0',
      'white-space: pre',
    ].join(';'));
    expect(BROWSER_CONSOLE_TUI_STYLE).not.toMatch(/(?:^|;)\s*(?:color|background)(?:-|:)/);
  });

  it('keeps Node consoles plain when a DOM shim defines window', () => {
    const consoleLog = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    vi.stubGlobal('window', {});
    vi.stubGlobal('process', { release: { name: 'node' } });

    emitConsoleTui('┌─┐');

    expect(consoleLog).toHaveBeenCalledWith('┌─┐');
  });

  it('keeps Worker consoles plain and single-argument', () => {
    const consoleLog = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    vi.stubGlobal('window', undefined);
    vi.stubGlobal('self', {});
    vi.stubGlobal('process', undefined);

    emitConsoleTui('┌─┐');

    expect(consoleLog).toHaveBeenCalledWith('┌─┐');
  });
});

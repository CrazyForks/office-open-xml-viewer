/** Typography-only CSS for fixed-cell console output in browser DevTools. */
export const BROWSER_CONSOLE_TUI_STYLE = [
  'font-family: "SFMono-Regular", Menlo, Monaco, Consolas, monospace',
  'font-size: 12px',
  'line-height: 1.35',
  'font-variant-ligatures: none',
  'letter-spacing: 0',
  'white-space: pre',
].join(';');

/**
 * Writes a preformatted TUI without imposing foreground or background colors.
 * Browser DevTools receive fixed-width typography; Node and workers keep a
 * plain, single-argument log entry.
 */
export function emitConsoleTui(output: string): void {
  if (isBrowserWindowRuntime()) {
    console.log(`%c${output}`, BROWSER_CONSOLE_TUI_STYLE);
    return;
  }
  console.log(output);
}

function isBrowserWindowRuntime(): boolean {
  const runtime = globalThis as typeof globalThis & {
    process?: { release?: { name?: string } };
  };
  return typeof window !== 'undefined'
    && runtime.process?.release?.name !== 'node';
}

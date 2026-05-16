import type { Meta, StoryObj } from '@storybook/html';
import {
  pptxToMarkdown,
  docxToMarkdown,
  xlsxToMarkdown,
  initPptxFromBytes,
  initDocxFromBytes,
  initXlsxFromBytes,
} from './index';

// Vite-side `?url` imports of the wasm-pack output — the same artifacts
// the viewer worker.ts files use, so we reuse them here without a second
// copy of the binary.
import pptxWasmUrl from '../../pptx/src/wasm/pptx_parser_bg.wasm?url';
import docxWasmUrl from '../../docx/src/wasm/docx_parser_bg.wasm?url';
import xlsxWasmUrl from '../../xlsx/src/wasm/xlsx_parser_bg.wasm?url';

type Args = Record<string, never>;

const meta: Meta<Args> = {
  title: 'Markdown',
};
export default meta;

type Story = StoryObj<Args>;

// Per-format WASM init is sync once the bytes are in hand. We cache the
// fetch+init promise so loading several files of the same format only
// pays the .wasm download once.
const initOnce = (() => {
  const cache = new Map<string, Promise<void>>();
  return (key: string, url: string, init: (bytes: Uint8Array) => void) => {
    let p = cache.get(key);
    if (!p) {
      p = fetch(url)
        .then((r) => r.arrayBuffer())
        .then((buf) => init(new Uint8Array(buf)));
      cache.set(key, p);
    }
    return p;
  };
})();

interface FormatConfig {
  title: string;
  demoUrl: string;
  accept: string;
  wasmUrl: string;
  initKey: string;
  init: (bytes: Uint8Array) => void;
  convert: (buf: ArrayBuffer | Uint8Array) => string;
}

function buildStory(cfg: FormatConfig): Story {
  return {
    name: cfg.title,
    render() {
      const root = document.createElement('div');
      root.style.cssText = 'font-family:sans-serif;padding:16px;';

      const toolbar = document.createElement('div');
      toolbar.style.cssText = 'display:flex;gap:8px;align-items:center;margin-bottom:10px;flex-wrap:wrap;';

      const demoBtn = document.createElement('button');
      demoBtn.textContent = `Load demo`;
      const fileInput = document.createElement('input');
      fileInput.type = 'file';
      fileInput.accept = cfg.accept;
      const status = document.createElement('span');
      status.style.cssText = 'font-size:13px;color:#444;';
      toolbar.append(demoBtn, fileInput, status);
      root.append(toolbar);

      const stats = document.createElement('div');
      stats.style.cssText = 'font-size:12px;color:#666;margin-bottom:8px;min-height:18px;';
      root.append(stats);

      const pre = document.createElement('pre');
      pre.style.cssText =
        'font-size:12px;line-height:1.4;max-height:600px;overflow:auto;' +
        'background:#1e1e1e;color:#d4d4d4;padding:12px;border-radius:4px;white-space:pre-wrap;word-break:break-word;';
      pre.textContent = `Pick a ${cfg.accept} file or click "Load demo".`;
      root.append(pre);

      const run = async (buf: ArrayBuffer, label: string) => {
        status.textContent = 'Parsing…';
        try {
          await initOnce(cfg.initKey, cfg.wasmUrl, cfg.init);
          const t0 = performance.now();
          const md = cfg.convert(buf);
          const elapsed = performance.now() - t0;
          const inKB = (buf.byteLength / 1024).toFixed(1);
          const outKB = (new TextEncoder().encode(md).byteLength / 1024).toFixed(1);
          stats.textContent = `${label} — ${inKB} KB → ${outKB} KB markdown (${(buf.byteLength / md.length).toFixed(1)}× compression, ${elapsed.toFixed(0)} ms)`;
          pre.textContent = md;
          status.textContent = 'OK';
        } catch (err) {
          status.textContent = `Error: ${(err as Error).message}`;
        }
      };

      demoBtn.addEventListener('click', async () => {
        const r = await fetch(cfg.demoUrl);
        const buf = await r.arrayBuffer();
        await run(buf, cfg.demoUrl.split('/').pop() ?? 'demo');
      });
      fileInput.addEventListener('change', async () => {
        const f = fileInput.files?.[0];
        if (!f) return;
        const buf = await f.arrayBuffer();
        await run(buf, f.name);
      });

      return root;
    },
  };
}

export const Pptx = buildStory({
  title: 'PPTX → Markdown',
  demoUrl: `${import.meta.env.BASE_URL}pptx/demo/sample-1.pptx`,
  accept: '.pptx',
  wasmUrl: pptxWasmUrl,
  initKey: 'pptx',
  init: initPptxFromBytes,
  convert: pptxToMarkdown,
});

export const Docx = buildStory({
  title: 'DOCX → Markdown',
  demoUrl: `${import.meta.env.BASE_URL}docx/demo/sample-1.docx`,
  accept: '.docx',
  wasmUrl: docxWasmUrl,
  initKey: 'docx',
  init: initDocxFromBytes,
  convert: docxToMarkdown,
});

export const Xlsx = buildStory({
  title: 'XLSX → Markdown',
  demoUrl: `${import.meta.env.BASE_URL}xlsx/demo/sample-1.xlsx`,
  accept: '.xlsx',
  wasmUrl: xlsxWasmUrl,
  initKey: 'xlsx',
  init: initXlsxFromBytes,
  convert: xlsxToMarkdown,
});

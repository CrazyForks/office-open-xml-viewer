/**
 * Webview bootstrap script.
 *
 * Runs inside the VSCode Webview iframe. Receives the file bytes via the
 * `ooxml-init` message and instantiates the appropriate viewer:
 *   - docx / pptx: scroll-stacked render of every page / slide with a transparent
 *     text layer for native selection (PDF.js-style).
 *   - xlsx: XlsxViewer (sheet-based, no scroll stack)
 */

declare const __OOXML_FILE_TYPE__: 'docx' | 'xlsx' | 'pptx';

declare function acquireVsCodeApi(): {
  postMessage(msg: unknown): void;
  getState(): unknown;
  setState(state: unknown): void;
};

import { XlsxViewer, type CellRange } from '@silurus/ooxml-xlsx';
import { DocxDocument, buildDocxTextLayer, type DocxTextRunInfo } from '@silurus/ooxml-docx';
import { PptxPresentation, buildPptxTextLayer, type PptxTextRunInfo } from '@silurus/ooxml-pptx';
import { svgExtents } from '@silurus/ooxml-core';
// Side-effect import: bundles the self-contained MathJax + STIX Two Math engine
// into the webview and sets globalThis.__ooxmlStix2. The library renders OMML
// equations only when handed a `math` engine; its built-in engine loads lazily
// by injecting a <script>, which this webview's nonce CSP blocks — so we bundle
// the engine inline instead and pass the adapter below to the viewers.
import '@silurus/ooxml-core/mathjax-stix2';

const math = {
  loadMathJax: async (): Promise<void> => {
    /* engine bundled by the import above; globalThis.__ooxmlStix2 is set */
  },
  mathMLToSvg: async (mathml: string) => {
    if (!__ooxmlStix2) throw new Error('Math engine failed to initialize');
    const svg = __ooxmlStix2.mathml2svg(mathml);
    return { svg, ...svgExtents(svg) };
  },
};

const vscodeApi = acquireVsCodeApi();
const fileType = __OOXML_FILE_TYPE__;

const statusEl = document.getElementById('status')!;
const viewerContainer = document.getElementById('viewer-container')!;

function showError(msg: string): void {
  statusEl.dataset.state = 'error';
  statusEl.textContent = msg;
  statusEl.style.display = '';
}

function hideStatus(): void {
  statusEl.style.display = 'none';
}

// Notify extension host that the webview script is ready to receive messages.
vscodeApi.postMessage({ type: 'webview-ready' });

window.addEventListener('message', async (event: MessageEvent) => {
  const msg = event.data;
  if (msg.type !== 'ooxml-init') return;

  // Opt-in flag forwarded from the extension host (gated by the
  // `ooxmlViewer.useGoogleFonts` setting AND workspace trust). When false the
  // viewers never touch the network — the matching CSP keeps the webview offline.
  const useGoogleFonts: boolean = msg.useGoogleFonts === true;

  let buffer: ArrayBuffer;
  try {
    const res = await fetch(msg.url);
    if (!res.ok) throw new Error(`Fetch failed: ${res.status} ${res.statusText}`);
    buffer = await res.arrayBuffer();
  } catch (err) {
    showError(`Error: ${err instanceof Error ? err.message : String(err)}`);
    return;
  }

  try {
    if (fileType === 'docx') {
      await initDocx(buffer, useGoogleFonts);
    } else if (fileType === 'xlsx') {
      await initXlsx(buffer, useGoogleFonts);
    } else if (fileType === 'pptx') {
      await initPptx(buffer, useGoogleFonts);
    }
  } catch (err) {
    showError(`Error: ${err instanceof Error ? err.message : String(err)}`);
  }
});

// ── XLSX ─────────────────────────────────────────────────────────────────────

async function initXlsx(buffer: ArrayBuffer, useGoogleFonts: boolean): Promise<void> {
  const container = document.createElement('div');
  container.style.cssText = 'width:100%;height:100vh;';
  viewerContainer.appendChild(container);

  const viewer = new XlsxViewer(container, {
    math,
    useGoogleFonts,
    onError(err) {
      showError(`Error: ${err.message}`);
    },
    onSelectionChange(sel: CellRange | null) {
      if (!sel) return;
      vscodeApi.postMessage({ type: 'selection', fileType: 'xlsx', selection: sel });
    },
  });

  await viewer.load(buffer);
  hideStatus();

  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
      const sel = viewer.selection;
      if (!sel) return;
      vscodeApi.postMessage({ type: 'copy-request', fileType: 'xlsx', selection: sel });
    }
  });
}

// ── DOCX (scroll view) ───────────────────────────────────────────────────────
//
// The text-selection overlay is built by the shared, public buildDocxTextLayer
// (imported above) rather than a local duplicate: this webview's `.page-canvas`
// is `width:100%` of `.page-wrapper` (itself capped by an inline `max-width`),
// so it IS responsively scaled by the editor pane's width, exactly the pattern
// the shared builder now guards (an overlay pinned to literal px would overflow
// `.page-wrapper` under a narrow pane, pushing scroll onto an ancestor — the same
// bug fixed for PptxViewer/DocxViewer). buildDocxTextLayer takes the page's
// intended CSS box (px, numbers) as the `%` denominators and leaves the
// `.text-layer` container's own `width:100%;height:100%` (set by CSS class)
// untouched.

async function initDocx(buffer: ArrayBuffer, useGoogleFonts: boolean): Promise<void> {
  const doc = await DocxDocument.load(buffer, { math, useGoogleFonts });

  const stack = document.createElement('div');
  stack.className = 'page-stack';
  viewerContainer.appendChild(stack);

  const widthPx = Math.min(window.innerWidth - 64, 900);

  for (let i = 0; i < doc.pageCount; i++) {
    const wrapper = document.createElement('div');
    wrapper.className = 'page-wrapper';
    wrapper.style.maxWidth = `${widthPx}px`;

    const canvas = document.createElement('canvas');
    canvas.className = 'page-canvas';

    const textLayer = document.createElement('div');
    textLayer.className = 'text-layer';

    wrapper.append(canvas, textLayer);
    stack.appendChild(wrapper);

    const runs: DocxTextRunInfo[] = [];
    await doc.renderPage(canvas, i, { width: widthPx, onTextRun: (r) => runs.push(r) });
    const cssHeight = parseFloat(canvas.style.height) || canvas.height;
    buildDocxTextLayer(textLayer, runs, widthPx, cssHeight);
  }

  hideStatus();
}

// ── PPTX (scroll view) ───────────────────────────────────────────────────────
//
// Same rationale as the docx section above: the shared, public
// buildPptxTextLayer (imported above) replaces the former local duplicate,
// which pinned `.text-layer`'s width/height and every shape frame / span to
// literal px — the same responsive-overflow bug fixed for PptxViewer, and a
// live one here too (`.page-canvas` is `width:100%` of `.page-wrapper`, capped
// by an inline `max-width`, so it IS responsively scaled by the editor pane).

async function initPptx(buffer: ArrayBuffer, useGoogleFonts: boolean): Promise<void> {
  const pres = await PptxPresentation.load(buffer, { math, useGoogleFonts });

  const stack = document.createElement('div');
  stack.className = 'page-stack';
  viewerContainer.appendChild(stack);

  const widthPx = Math.min(window.innerWidth - 64, 960);
  const cssHeight = pres.slideWidth > 0
    ? Math.round((pres.slideHeight * widthPx) / pres.slideWidth)
    : Math.round((widthPx * 9) / 16);

  for (let i = 0; i < pres.slideCount; i++) {
    const wrapper = document.createElement('div');
    wrapper.className = 'page-wrapper';
    wrapper.style.maxWidth = `${widthPx}px`;

    const canvas = document.createElement('canvas');
    canvas.className = 'page-canvas';

    const textLayer = document.createElement('div');
    textLayer.className = 'text-layer';

    wrapper.append(canvas, textLayer);
    stack.appendChild(wrapper);

    const runs: PptxTextRunInfo[] = [];
    await pres.presentSlide(canvas, i, { width: widthPx, onTextRun: (r) => runs.push(r) });
    buildPptxTextLayer(textLayer, runs, widthPx, cssHeight);
  }

  hideStatus();
}

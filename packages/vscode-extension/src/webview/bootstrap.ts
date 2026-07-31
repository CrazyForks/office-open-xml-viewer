/**
 * Webview bootstrap script.
 *
 * Runs inside the VSCode Webview iframe. Receives the file bytes via the
 * `ooxml-init` message and instantiates the appropriate viewer:
 *   - docx / pptx: their virtualized continuous-scroll viewers.
 *   - xlsx: its sheet-based scrolling viewer.
 *
 * All three satisfy the shared ZoomableViewer contract, which drives the
 * extension toolbar's zoom-out / zoom-in controls.
 */

declare const __OOXML_FILE_TYPE__: 'docx' | 'xlsx' | 'pptx';

declare function acquireVsCodeApi(): {
  postMessage(msg: unknown): void;
  getState(): unknown;
  setState(state: unknown): void;
};

import { XlsxViewer, type CellRange } from '@silurus/ooxml-xlsx';
import { DocxScrollViewer } from '@silurus/ooxml-docx';
import { PptxScrollViewer } from '@silurus/ooxml-pptx';
import { svgExtents, type ZoomableViewer } from '@silurus/ooxml-core';
import { loadAtContainerFit, loadAtNaturalScale } from './naturalScale';
import { captureUnhandledWheelZoom } from './wheelZoomFallback';
import { FindPopupController, type FindableViewer } from './findPopup';
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
const zoomOutButton = document.getElementById('zoom-out') as HTMLButtonElement;
const zoomInButton = document.getElementById('zoom-in') as HTMLButtonElement;
const zoomLabel = document.getElementById('zoom-label')!;
const findHighlightColors = {
  match: 'var(--vscode-ooxmlViewer-findMatchBackground, rgba(255, 214, 0, 0.42))',
  active: 'var(--vscode-ooxmlViewer-findActiveMatchBackground, rgba(255, 140, 0, 0.55))',
};
let activeViewer: ZoomableViewer | null = null;

function showError(msg: string): void {
  statusEl.dataset.state = 'error';
  statusEl.textContent = msg;
  statusEl.style.display = '';
}

function hideStatus(): void {
  statusEl.style.display = 'none';
}

function updateZoomLabel(scale: number): void {
  zoomLabel.textContent = `${Math.round(scale * 100)}%`;
}

const findPopup = new FindPopupController(
  {
    root: document.getElementById('find-popup')!,
    input: document.getElementById('find-input') as HTMLInputElement,
    status: document.getElementById('find-status')!,
    previous: document.getElementById('find-previous') as HTMLButtonElement,
    next: document.getElementById('find-next') as HTMLButtonElement,
    close: document.getElementById('find-close') as HTMLButtonElement,
  },
  window,
  (err) => showError(`Error: ${err instanceof Error ? err.message : String(err)}`),
);

function bindZoomViewer(viewer: ZoomableViewer & FindableViewer): void {
  activeViewer = viewer;
  findPopup.setViewer(viewer);
  zoomOutButton.disabled = false;
  zoomInButton.disabled = false;
  updateZoomLabel(viewer.getScale());
}

async function runZoom(action: 'zoomIn' | 'zoomOut'): Promise<void> {
  if (!activeViewer) return;
  try {
    await activeViewer[action]();
    updateZoomLabel(activeViewer.getScale());
  } catch (err) {
    showError(`Error: ${err instanceof Error ? err.message : String(err)}`);
  }
}

zoomOutButton.addEventListener('click', () => {
  void runZoom('zoomOut');
});
zoomInButton.addEventListener('click', () => {
  void runZoom('zoomIn');
});

// Chromium normally sends a trackpad pinch as Ctrl+wheel to the viewer-owned
// scroll host. VS Code's webview can target the frame instead, so capture it at
// the window and fall back only when the viewer's own handler did not run.
window.addEventListener(
  'wheel',
  (event) => {
    captureUnhandledWheelZoom(event, activeViewer, undefined, (err) => {
      showError(`Error: ${err instanceof Error ? err.message : String(err)}`);
    });
  },
  { capture: true, passive: false },
);

// Notify extension host that the webview script is ready to receive messages.
vscodeApi.postMessage({ type: 'webview-ready' });

window.addEventListener('message', async (event: MessageEvent) => {
  const msg = event.data;
  if (msg.type === 'show-find') {
    findPopup.open();
    return;
  }
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
  let loadFailed = false;
  const viewer = new XlsxViewer(viewerContainer, {
    math,
    useGoogleFonts,
    showZoomSlider: false,
    findHighlightColors,
    onScaleChange: updateZoomLabel,
    onError(err) {
      loadFailed = true;
      showError(`Error: ${err.message}`);
    },
    onSelectionChange(sel: CellRange | null) {
      if (!sel) return;
      vscodeApi.postMessage({ type: 'selection', fileType: 'xlsx', selection: sel });
    },
  });

  await viewer.load(buffer);
  if (loadFailed) return;
  bindZoomViewer(viewer);
  hideStatus();

  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
      const sel = viewer.selection;
      if (!sel) return;
      vscodeApi.postMessage({ type: 'copy-request', fileType: 'xlsx', selection: sel });
    }
  });
}

// ── DOCX (virtualized continuous scroll) ─────────────────────────────────────

async function initDocx(buffer: ArrayBuffer, useGoogleFonts: boolean): Promise<void> {
  let loadFailed = false;
  const viewer = new DocxScrollViewer(viewerContainer, {
    math,
    useGoogleFonts,
    enableTextSelection: true,
    findHighlightColors,
    refitOnResize: false,
    background: 'var(--vscode-editor-background)',
    onScaleChange: updateZoomLabel,
    onError(err) {
      loadFailed = true;
      showError(`Error: ${err.message}`);
    },
  });
  await loadAtNaturalScale(viewer, buffer);
  if (loadFailed) return;
  bindZoomViewer(viewer);
  hideStatus();
}

// ── PPTX (virtualized continuous scroll) ─────────────────────────────────────

async function initPptx(buffer: ArrayBuffer, useGoogleFonts: boolean): Promise<void> {
  let loadFailed = false;
  const viewer = new PptxScrollViewer(viewerContainer, {
    math,
    useGoogleFonts,
    enableTextSelection: true,
    findHighlightColors,
    enableMediaPlayback: true,
    mediaOverscan: 1,
    refitOnResize: true,
    background: 'var(--vscode-editor-background)',
    onScaleChange: updateZoomLabel,
    onError(err) {
      loadFailed = true;
      showError(`Error: ${err.message}`);
    },
  });
  // Presentations default to the available editor width. The viewer subtracts
  // its built-in desk gutters, leaving a small margin without horizontal scroll.
  await loadAtContainerFit(viewer, buffer);
  if (loadFailed) return;
  bindZoomViewer(viewer);
  hideStatus();
}

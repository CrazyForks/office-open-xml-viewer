// "Try yours" — render a user-supplied file entirely in the browser. The file
// is read with FileReader/arrayBuffer and parsed by the WASM engines; it never
// leaves the page (no upload, no server).
import { PptxPresentation } from '@silurus/ooxml-pptx';
import { DocxDocument } from '@silurus/ooxml-docx';
import { XlsxViewer } from '@silurus/ooxml-xlsx';
import { loadMathJax, mathMLToSvg } from '../../../packages/core/src/math/engine';

// Opt-in OMML equation engine — enabled here so user-supplied docx/pptx with
// equations render. (In the published library this is `@silurus/ooxml/math`.)
const math = { loadMathJax, mathMLToSvg };

const DPR = () => Math.min(typeof window !== 'undefined' ? window.devicePixelRatio : 1, 2);

// Render slides at the width they actually display at (`.lv-page` max-width),
// not larger. Rendering bigger than the display only shrinks the interactive
// media controls (drawn at fixed px in canvas space) when the canvas is
// CSS-downscaled to fit — making them look tiny next to the slide. dpr keeps
// the backing store crisp on HiDPI.
const SLIDE_W = 880;

type SlideHandle = Awaited<ReturnType<PptxPresentation['presentSlide']>>;

// Disposes the previous render's live resources (interactive slide handles +
// their IntersectionObserver) so audio/video stops and RAF loops are released
// when a new file is loaded.
let activeCleanup: (() => void) | null = null;

export interface RenderResult {
  format: 'docx' | 'xlsx' | 'pptx';
  units: number; // pages / slides; 0 for xlsx (sheet-based)
  unitLabel: string;
}

export async function renderFile(stage: HTMLElement, file: File): Promise<RenderResult> {
  const ext = file.name.split('.').pop()?.toLowerCase();
  if (ext !== 'docx' && ext !== 'xlsx' && ext !== 'pptx') {
    throw new Error('Unsupported file — choose a .docx, .xlsx or .pptx file.');
  }
  // Tear down any live handles from a previous file before replacing the stage.
  activeCleanup?.();
  activeCleanup = null;

  const buffer = await file.arrayBuffer();
  stage.innerHTML = '';

  if (ext === 'xlsx') {
    const host = document.createElement('div');
    host.className = 'lv-xlsx';
    stage.appendChild(host);
    const viewer = new XlsxViewer(host, { useGoogleFonts: true, showZoomSlider: true, math });
    await viewer.load(buffer);
    return { format: 'xlsx', units: 0, unitLabel: 'sheet' };
  }

  const sc = document.createElement('div');
  sc.className = 'lv-scroll';
  stage.appendChild(sc);

  if (ext === 'pptx') {
    const deck = await PptxPresentation.load(buffer, { useGoogleFonts: true, math });
    const canvases: HTMLCanvasElement[] = [];
    for (let i = 0; i < deck.slideCount; i++) {
      const c = document.createElement('canvas');
      c.className = 'lv-page';
      c.dataset.slide = String(i);
      sc.appendChild(c);
      await deck.renderSlide(c, i, { width: SLIDE_W, dpr: DPR() });
      canvases.push(c);
    }

    // Audio/video playback: upgrade slides to an interactive PresentationHandle
    // (click-to-play media + scrubber) — but only while they are on-screen, so
    // the per-handle RAF loop and decoded media stay bounded no matter how many
    // slides the deck has. Off-screen slides keep their static base render.
    const handles = new Map<number, SlideHandle>();
    const pending = new Set<number>();
    const visible = new Set<number>();
    const io = new IntersectionObserver(
      (entries) => {
        for (const e of entries) {
          const c = e.target as HTMLCanvasElement;
          const idx = Number(c.dataset.slide);
          if (e.isIntersecting) {
            visible.add(idx);
            if (handles.has(idx) || pending.has(idx)) continue;
            pending.add(idx);
            deck
              .presentSlide(c, idx, { width: SLIDE_W, dpr: DPR() })
              .then((h) => {
                pending.delete(idx);
                if (visible.has(idx)) handles.set(idx, h);
                else h.destroy(); // scrolled away before it resolved
              })
              .catch(() => pending.delete(idx));
          } else {
            visible.delete(idx);
            const h = handles.get(idx);
            if (h) {
              h.destroy(); // stops media + RAF; the static base frame remains
              handles.delete(idx);
            }
          }
        }
      },
      { root: sc, rootMargin: '150px 0px' },
    );
    canvases.forEach((c) => io.observe(c));
    activeCleanup = () => {
      io.disconnect();
      handles.forEach((h) => h.destroy());
      handles.clear();
      pending.clear();
      visible.clear();
    };

    return { format: 'pptx', units: deck.slideCount, unitLabel: 'slide' };
  }

  const doc = await DocxDocument.load(buffer, { useGoogleFonts: true, math });
  for (let i = 0; i < doc.pageCount; i++) {
    const c = document.createElement('canvas');
    c.className = 'lv-page';
    sc.appendChild(c);
    await doc.renderPage(c, i, { width: 1000, dpr: DPR() });
  }
  return { format: 'docx', units: doc.pageCount, unitLabel: 'page' };
}

// Hot standby: warm each WASM engine (and the Google Fonts CSS) on an idle tick
// so the user's first real file parses without paying the cold cost of fetching
// + compiling the parser binaries. Each `renderFile` spawns a fresh inline
// worker that re-fetches its `*_parser_bg.wasm`; pre-loading the bundled demo of
// every format primes the browser's HTTP/code cache for that binary, then the
// throwaway engines are released. Fire-and-forget, errors swallowed — a failed
// warm-up just means the first real parse is as slow as before, never broken.
let warmed = false;
export function prewarmEngines(): void {
  if (warmed || typeof window === 'undefined') return;
  warmed = true;
  // Respect Data Saver / metered connections — don't spend bandwidth warming.
  const conn = (navigator as Navigator & { connection?: { saveData?: boolean } }).connection;
  if (conn?.saveData) return;

  const base = import.meta.env.BASE_URL;
  const sample = (f: string) => `${base}samples/${f}`.replace(/([^:])\/\/+/g, '$1/');

  const run = (): void => {
    void PptxPresentation.load(sample('sample-1.pptx'), { useGoogleFonts: true })
      .then((d) => d.destroy())
      .catch(() => {});
    void DocxDocument.load(sample('sample-1.docx'), { useGoogleFonts: true })
      .then((d) => d.destroy())
      .catch(() => {});
    // XlsxViewer needs a container; mount into a detached node never added to the
    // DOM, then dispose. The parse + one render warms the xlsx WASM engine.
    const host = document.createElement('div');
    const v = new XlsxViewer(host, { useGoogleFonts: true });
    void v.load(sample('sample-1.xlsx')).then(() => v.destroy()).catch(() => v.destroy());
  };

  const ric = (window as Window & { requestIdleCallback?: (cb: () => void, o?: { timeout: number }) => void })
    .requestIdleCallback;
  if (ric) ric(run, { timeout: 2500 });
  else setTimeout(run, 600);
}

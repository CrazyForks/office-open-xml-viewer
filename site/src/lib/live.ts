// Client-side mounts for the three real viewers, used by the Live Showcase.
// Each mount returns a `destroy()` so the showcase can swap formats cleanly.
import { PptxViewer } from '@silurus/ooxml-pptx';
import { DocxViewer } from '@silurus/ooxml-docx';
import { XlsxViewer } from '@silurus/ooxml-xlsx';

export type LiveController = { destroy: () => void };

function controlBar(): { bar: HTMLDivElement; prev: HTMLButtonElement; next: HTMLButtonElement; info: HTMLSpanElement } {
  const bar = document.createElement('div');
  bar.className = 'lv-bar';
  const prev = document.createElement('button');
  prev.className = 'lv-btn';
  prev.textContent = '‹';
  prev.disabled = true;
  const next = document.createElement('button');
  next.className = 'lv-btn';
  next.textContent = '›';
  next.disabled = true;
  const info = document.createElement('span');
  info.className = 'lv-info';
  info.textContent = 'Loading…';
  bar.append(prev, info, next);
  return { bar, prev, next, info };
}

function stage(fill = false): HTMLDivElement {
  const s = document.createElement('div');
  s.className = fill ? 'lv-stage lv-stage--fill' : 'lv-stage';
  return s;
}

export function mountPptx(root: HTMLElement, url: string, width = 1280): LiveController {
  root.innerHTML = '';
  const { bar, prev, next, info } = controlBar();
  const st = stage(true);
  const canvas = document.createElement('canvas');
  st.appendChild(canvas);
  root.append(bar, st);

  const viewer = new PptxViewer(canvas, {
    width,
    useGoogleFonts: true,
    enableTextSelection: true,
    onSlideChange: (idx, total) => {
      info.textContent = `Slide ${idx + 1} / ${total}`;
      prev.disabled = idx === 0;
      next.disabled = idx === total - 1;
    },
    onError: (err) => { info.textContent = `Error: ${err.message}`; },
  });
  prev.addEventListener('click', () => void viewer.prevSlide());
  next.addEventListener('click', () => void viewer.nextSlide());
  viewer.load(url).catch((e: unknown) => { info.textContent = msg(e); });

  return { destroy: () => { root.innerHTML = ''; } };
}

export function mountDocx(root: HTMLElement, url: string, width = 760): LiveController {
  root.innerHTML = '';
  const { bar, prev, next, info } = controlBar();
  const st = stage();
  const canvas = document.createElement('canvas');
  st.appendChild(canvas);
  root.append(bar, st);

  const viewer = new DocxViewer(canvas, {
    width,
    dpr: window.devicePixelRatio,
    useGoogleFonts: true,
    enableTextSelection: true,
  });
  const sync = () => {
    const total = viewer.pageCount;
    info.textContent = total ? `Page ${viewer.currentPage + 1} / ${total}` : 'Loading…';
    prev.disabled = viewer.currentPage <= 0;
    next.disabled = viewer.currentPage >= total - 1;
  };
  prev.addEventListener('click', () => { void viewer.prevPage().then(sync); });
  next.addEventListener('click', () => { void viewer.nextPage().then(sync); });
  viewer.load(url).then(sync).catch((e: unknown) => { info.textContent = msg(e); });

  return { destroy: () => { root.innerHTML = ''; } };
}

export function mountXlsx(root: HTMLElement, url: string): LiveController {
  root.innerHTML = '';
  const host = document.createElement('div');
  host.className = 'lv-xlsx';
  root.append(host);

  const viewer = new XlsxViewer(host, {
    useGoogleFonts: true,
    showZoomSlider: true,
    onError: (err: Error) => { host.setAttribute('data-error', err.message); },
  });
  viewer.load(url).catch(() => { /* surfaced via onError */ });

  return { destroy: () => { root.innerHTML = ''; } };
}

function msg(e: unknown): string {
  return `Failed: ${e instanceof Error ? e.message : String(e)}`;
}

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
  const buffer = await file.arrayBuffer();
  stage.innerHTML = '';

  if (ext === 'xlsx') {
    const host = document.createElement('div');
    host.className = 'lv-xlsx';
    stage.appendChild(host);
    const viewer = new XlsxViewer(host, { useGoogleFonts: true, showZoomSlider: true });
    await viewer.load(buffer);
    return { format: 'xlsx', units: 0, unitLabel: 'sheet' };
  }

  const sc = document.createElement('div');
  sc.className = 'lv-scroll';
  stage.appendChild(sc);

  if (ext === 'pptx') {
    const deck = await PptxPresentation.load(buffer, { useGoogleFonts: true });
    for (let i = 0; i < deck.slideCount; i++) {
      const c = document.createElement('canvas');
      c.className = 'lv-page';
      sc.appendChild(c);
      await deck.renderSlide(c, i, { width: 1280, dpr: DPR(), math });
    }
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

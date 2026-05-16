/**
 * Shared multi-page export primitives. Each format package re-exports a
 * format-specific wrapper that fills in the renderer; the heavy lifting
 * (offscreen canvas, PNG encoding, multi-page PDF assembly via pdf-lib)
 * lives here so the three viewers stay in sync.
 *
 * pdf-lib is a lazy import — bundles that don't call `pagesToPdfBlob`
 * never pull it in.
 */

export interface PageBitmap {
  /** 0-based index in the source document. */
  index: number;
  /** PNG bytes for the rendered page / slide. */
  blob: Blob;
  /** Pixel dimensions of the PNG (multiplied by dpr already). */
  pixelWidth: number;
  pixelHeight: number;
  /** Page size in points (PDF units = 1/72 inch). Used when assembling PDFs. */
  pointWidth: number;
  pointHeight: number;
}

export interface RenderPageToCanvasContext {
  /** Total number of pages / slides / sheets to render. */
  pageCount: number;
  /** Draw a page onto a caller-supplied canvas. Implementations should
   *  size the canvas appropriately for the requested width / dpr. */
  renderPage: (canvas: HTMLCanvasElement, pageIndex: number, opts: { width: number; dpr: number }) => Promise<void>;
  /** Convert a logical width (px in the rendered output) to PDF points
   *  for the page size of the given page. */
  pageSizeInPoints: (pageIndex: number) => { widthPt: number; heightPt: number };
}

export interface ExportPngOptions {
  /** Output width in CSS pixels (height derived from the page aspect ratio). */
  width?: number;
  /** Device pixel ratio. Default 2. */
  dpr?: number;
}

/** Render a single page to a PNG `Blob`. */
export async function renderPageToPng(
  ctx: RenderPageToCanvasContext,
  pageIndex: number,
  opts: ExportPngOptions = {},
): Promise<PageBitmap> {
  const width = opts.width ?? 1280;
  const dpr = opts.dpr ?? 2;
  const canvas = createOffscreen();
  await ctx.renderPage(canvas, pageIndex, { width, dpr });
  const blob = await canvasToPngBlob(canvas);
  const { widthPt, heightPt } = ctx.pageSizeInPoints(pageIndex);
  return {
    index: pageIndex,
    blob,
    pixelWidth: canvas.width,
    pixelHeight: canvas.height,
    pointWidth: widthPt,
    pointHeight: heightPt,
  };
}

/** Render every page as PNG. */
export async function renderAllPagesToPng(
  ctx: RenderPageToCanvasContext,
  opts: ExportPngOptions = {},
): Promise<PageBitmap[]> {
  const out: PageBitmap[] = [];
  for (let i = 0; i < ctx.pageCount; i++) {
    out.push(await renderPageToPng(ctx, i, opts));
  }
  return out;
}

/** Stitch a list of rendered pages into a single PDF Blob via `pdf-lib`. */
export async function pagesToPdfBlob(pages: PageBitmap[]): Promise<Blob> {
  const { PDFDocument } = await import('pdf-lib');
  const pdf = await PDFDocument.create();
  for (const page of pages) {
    const bytes = new Uint8Array(await page.blob.arrayBuffer());
    const image = await pdf.embedPng(bytes);
    const pdfPage = pdf.addPage([page.pointWidth, page.pointHeight]);
    pdfPage.drawImage(image, {
      x: 0,
      y: 0,
      width: page.pointWidth,
      height: page.pointHeight,
    });
  }
  const pdfBytes = await pdf.save();
  return new Blob([pdfBytes as BlobPart], { type: 'application/pdf' });
}

function createOffscreen(): HTMLCanvasElement {
  // Always pick an HTMLCanvasElement (not OffscreenCanvas) so the renderers
  // can `instanceof`-check it and apply CSS sizing logic uniformly.
  const canvas = document.createElement('canvas');
  return canvas;
}

function canvasToPngBlob(canvas: HTMLCanvasElement): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error('canvas.toBlob returned null — encoder failed'));
        return;
      }
      resolve(blob);
    }, 'image/png');
  });
}

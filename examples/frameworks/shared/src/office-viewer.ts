import type { DocxViewer, DocxViewerOptions } from '@silurus/ooxml/docx';
import type { PptxViewer, PptxViewerOptions } from '@silurus/ooxml/pptx';
import type { XlsxViewer, XlsxViewerOptions } from '@silurus/ooxml/xlsx';

export type OfficeFormat = 'docx' | 'xlsx' | 'pptx';
export type OfficeSource = string | ArrayBuffer;

export interface OfficeViewerByFormat {
  docx: DocxViewer;
  xlsx: XlsxViewer;
  pptx: PptxViewer;
}

export interface OfficeViewerOptionsByFormat {
  docx: DocxViewerOptions;
  xlsx: XlsxViewerOptions;
  pptx: PptxViewerOptions;
}

export interface OfficeViewerTargetByFormat {
  docx: HTMLCanvasElement;
  xlsx: HTMLElement;
  pptx: HTMLCanvasElement;
}

export interface OfficeViewerConfig<F extends OfficeFormat> {
  format: F;
  source: OfficeSource;
  options?: OfficeViewerOptionsByFormat[F];
}

export interface MountOfficeViewerConfig<F extends OfficeFormat> extends OfficeViewerConfig<F> {
  target: OfficeViewerTargetByFormat[F];
}

export type OfficeViewerInstance = OfficeViewerByFormat[OfficeFormat];

function requireCanvas(target: HTMLElement, format: 'docx' | 'pptx'): HTMLCanvasElement {
  if (target instanceof HTMLCanvasElement) return target;
  throw new TypeError(`${format.toUpperCase()} viewers must be mounted on a canvas element.`);
}

export async function createOfficeViewer<F extends OfficeFormat>(
  format: F,
  target: OfficeViewerTargetByFormat[F],
  options?: OfficeViewerOptionsByFormat[F],
): Promise<OfficeViewerByFormat[F]> {
  switch (format) {
    case 'docx': {
      const { DocxViewer } = await import('@silurus/ooxml/docx');
      return new DocxViewer(
        requireCanvas(target, 'docx'),
        options as DocxViewerOptions,
      ) as OfficeViewerByFormat[F];
    }
    case 'xlsx': {
      const { XlsxViewer } = await import('@silurus/ooxml/xlsx');
      return new XlsxViewer(
        target,
        options as XlsxViewerOptions,
      ) as OfficeViewerByFormat[F];
    }
    case 'pptx': {
      const { PptxViewer } = await import('@silurus/ooxml/pptx');
      return new PptxViewer(
        requireCanvas(target, 'pptx'),
        options as PptxViewerOptions,
      ) as OfficeViewerByFormat[F];
    }
  }
}

export async function mountOfficeViewer<F extends OfficeFormat>({
  format,
  target,
  source,
  options,
}: MountOfficeViewerConfig<F>): Promise<OfficeViewerByFormat[F]> {
  const viewer = await createOfficeViewer(format, target, options);

  try {
    await viewer.load(source);
    return viewer;
  } catch (error) {
    viewer.destroy();
    throw error;
  }
}

export {
  parsePptx,
  openPptxPresentation,
  extractMedia as extractPptxMedia,
  extractImage as extractPptxImage,
  type OpenPptxPresentationOptions,
  type PptxPresentationSession,
  type PptxSessionRenderOptions,
} from './pptx';
export {
  openDocxDocument,
  parseDocx,
  type DocxDocumentSession,
  type DocxPageRenderOptions,
  type DocxRenderedPage,
  type OpenDocxDocumentOptions,
} from './docx';
export {
  openXlsxWorkbook,
  parseXlsx,
  parseSheet as parseXlsxSheet,
  parseXlsxAllSheets,
  type OpenXlsxWorkbookOptions,
  type XlsxWorkbookSession,
  type XlsxWorksheetRowChunk,
} from './xlsx';
export {
  renderSlideNode,
  makeSourceBufferFetchImage,
  installImageBitmapShim,
  installOffscreenCanvasShim,
  type NodeCanvasLike,
  type NodeCanvasFactory,
  type NodeImageLike,
} from './render';
export type { OoxmlNodeSessionOptions } from './session-options';
export type {
  OoxmlResourceMetrics,
  OoxmlResourceMetricsCheckpoint,
  OoxmlResourcePolicySnapshot,
} from '@silurus/ooxml-core';
export {
  OoxmlDecodedImageLimitError,
  OoxmlResourceLimitError,
  isOoxmlDecodedImageLimitError,
  type OoxmlDecodedImageLimitMetric,
  type OoxmlResourceLimit,
  type OoxmlResourceLimitErrorDetails,
  type OoxmlResourceLimits,
  type OoxmlResourceUsageSnapshot,
} from '@silurus/ooxml-core';

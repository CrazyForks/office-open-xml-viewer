/** Canonical XLSX workbook/worksheet coordinator entry point consumed by Node. */
export { resolveSharedStringRows } from '../shared-strings.js';
export {
  WorksheetPullWorker,
  XLSX_WORKSHEET_PULL_BYTES,
  type WorksheetWireChunk,
} from '../worksheet-pull-worker.js';
export {
  acquireXlsxNodeSession,
  type XlsxNodeAcquisition,
  type XlsxNodeAcquisitionOptions,
  type XlsxNodeArchive,
} from './node-acquisition.js';

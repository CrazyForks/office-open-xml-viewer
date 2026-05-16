export type {
  Change,
  ChangeOp,
  ChangeLocation,
  BBox,
  DiffResult,
  Format,
} from './types';
export { changesAtSlide, changesAtPage, changesAtSheet } from './types';

export { diffPptx, bboxesForSlide } from './pptx';
export { diffDocx } from './docx';
export { diffXlsx, type XlsxDiffInput } from './xlsx';

export { alignSequences, type SequenceAlignment } from './util/sequence';
export { deepEqual } from './util/equal';

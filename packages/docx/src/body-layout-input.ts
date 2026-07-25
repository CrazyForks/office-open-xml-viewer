import { bodyLayoutAcquisitionInput } from './parser-model.js';
import { projectBodyLayoutInput, type BodyLayoutInput } from './layout/body-layout-input.js';
import type { DocxDocumentModel } from './types.js';

export function createBodyLayoutInput(document: DocxDocumentModel): BodyLayoutInput {
  return projectBodyLayoutInput(bodyLayoutAcquisitionInput(document));
}

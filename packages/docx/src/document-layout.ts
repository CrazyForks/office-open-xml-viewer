/** Canonical immutable body page production for the DOCX renderer. */
import type { DocxDocumentModel } from './types.js';
import type { LayoutOptions } from './layout/options.js';
import type { LayoutServices } from './layout/types.js';
import { productionDocumentInput } from './layout/resources.js';
import { layoutDocumentInput } from './layout/document.js';
import { createLayoutServices } from './layout-runtime.js';

export function layoutDocument(
  doc: DocxDocumentModel,
  services: LayoutServices = createLayoutServices(doc),
  options?: LayoutOptions,
) {
  return layoutDocumentInput(productionDocumentInput(doc).bodyLayoutInput, services, options);
}
export type {
  DocumentLayout,
  LayoutPage,
  ParagraphLayout,
  TableLayout,
} from './layout/types.js';

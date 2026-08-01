/** Canonical immutable body page production for the DOCX renderer. */
import type { DocxDocumentModel } from './types.js';
import type { LayoutOptions } from './layout/options.js';
import type { LayoutServices } from './layout/types.js';
import { layoutDocumentInput } from './layout/document.js';
import {
  isLayoutSourceStore,
  type LayoutSourceStore,
} from './layout/layout-source-store.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { layoutSourceStoreOf } from './layout/runtime-state.js';
import { createLayoutServices } from './layout-runtime.js';

export function layoutDocument(
  input: DocxDocumentModel | LayoutSourceStore,
  services: LayoutServices = createLayoutServices(input),
  options?: LayoutOptions,
) {
  const source = isLayoutSourceStore(input) ? input : layoutSourceStore(input);
  const retainedSource = layoutSourceStoreOf(services);
  if (retainedSource && retainedSource !== source) {
    throw new Error('Layout services belong to a different document source');
  }
  return layoutDocumentInput(source.bodyLayoutInput, services, options);
}
export type {
  DocumentLayout,
  LayoutPage,
  ParagraphLayout,
  TableLayout,
} from './layout/types.js';

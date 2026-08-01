import type { LayoutServices } from './layout/types.js';
import { attachDocumentLayoutVariants } from './layout/document-layout-variants.js';
import type { LayoutVariantStore } from './layout/variant-store.js';
import { layoutDocument } from './document-layout.js';
import type { LayoutSourceStore } from './layout/layout-source-store.js';
import { layoutSourceStoreOf } from './layout/runtime-state.js';

export interface RetainedRenderWorkerDocumentLayout {
  readonly layoutServices: LayoutServices;
  readonly layoutVariants: LayoutVariantStore;
  readonly defaultCurrentDateMs: number;
}

/**
 * The worker keeps one document-scoped service graph and variant store. Keeping
 * their construction behind this pure seam lets parity tests execute the same
 * ownership wiring without importing the worker's WASM and `self` side effects.
 */
export function retainRenderWorkerDocumentLayout(
  source: LayoutSourceStore,
  layoutServices: LayoutServices,
  defaultCurrentDateMs: number,
): RetainedRenderWorkerDocumentLayout {
  const retainedSource = layoutSourceStoreOf(layoutServices);
  if (retainedSource && retainedSource !== source) {
    throw new Error('Layout services belong to a different document source');
  }
  const variants = attachDocumentLayoutVariants({
    source,
    services: layoutServices,
    defaultCurrentDateMs,
    buildLayout: (options) => layoutDocument(source, layoutServices, options),
  });
  return Object.freeze({
    layoutServices,
    layoutVariants: variants.store,
    defaultCurrentDateMs,
  });
}

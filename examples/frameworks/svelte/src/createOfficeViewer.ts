import { writable, type Readable } from 'svelte/store';
import type { Action } from 'svelte/action';

export type OfficeSource = string | ArrayBuffer;
export type OfficeViewerStatus = 'idle' | 'loading' | 'ready' | 'error';

export interface OfficeViewerHandle {
  load: (source: OfficeSource) => Promise<unknown>;
  destroy: () => void;
  getScale: () => number;
  setScale: (scale: number) => void;
  zoomIn: () => void;
  zoomOut: () => void;
  fitWidth: () => void;
  fitPage: () => void;
}

export type OfficeViewerFactory<V extends OfficeViewerHandle> = (
  container: HTMLElement,
) => V | Promise<V>;

export interface OfficeViewerConfig<V extends OfficeViewerHandle> {
  source: OfficeSource;
  createViewer: OfficeViewerFactory<V>;
}

export interface OfficeViewerAction<V extends OfficeViewerHandle> {
  action: Action<HTMLElement, OfficeViewerConfig<V>>;
  status: Readable<OfficeViewerStatus>;
  error: Readable<Error | null>;
  reload: () => void;
  getScale: () => number | undefined;
  setScale: (scale: number) => void;
  zoomIn: () => void;
  zoomOut: () => void;
  fitWidth: () => void;
  fitPage: () => void;
}

export function createOfficeViewer<V extends OfficeViewerHandle>(): OfficeViewerAction<V> {
  const status = writable<OfficeViewerStatus>('idle');
  const error = writable<Error | null>(null);
  const state = {
    node: null as HTMLElement | null,
    config: null as OfficeViewerConfig<V> | null,
    viewer: null as V | null,
    controller: null as AbortController | null,
  };

  const destroyViewer = () => {
    state.controller?.abort();
    state.controller = null;
    const viewer = state.viewer;
    state.viewer = null;
    viewer?.destroy();
  };

  const mount = () => {
    destroyViewer();
    const { node, config } = state;
    if (!node || !config) {
      status.set('idle');
      return;
    }

    const controller = new AbortController();
    state.controller = controller;
    status.set('loading');
    error.set(null);
    Promise.resolve(config.createViewer(node))
      .then(async (viewer) => {
        if (controller.signal.aborted) {
          viewer.destroy();
          return null;
        }
        try {
          await viewer.load(config.source);
          return viewer;
        } catch (reason) {
          viewer.destroy();
          throw reason;
        }
      })
      .then((viewer) => {
        if (!viewer || controller.signal.aborted) {
          viewer?.destroy();
          return;
        }
        state.viewer = viewer;
        status.set('ready');
      })
      .catch((reason: unknown) => {
        if (controller.signal.aborted) return;
        error.set(reason instanceof Error ? reason : new Error(String(reason)));
        status.set('error');
      });
  };

  const action: Action<HTMLElement, OfficeViewerConfig<V>> = (element, initialConfig) => {
    state.node = element;
    state.config = initialConfig;
    mount();

    return {
      update(nextConfig) {
        state.config = nextConfig;
        mount();
      },
      destroy() {
        destroyViewer();
        state.node = null;
        state.config = null;
        status.set('idle');
      },
    };
  };

  return {
    action,
    status,
    error,
    reload: mount,
    getScale: () => state.viewer?.getScale(),
    setScale: (scale) => state.viewer?.setScale(scale),
    zoomIn: () => state.viewer?.zoomIn(),
    zoomOut: () => state.viewer?.zoomOut(),
    fitWidth: () => state.viewer?.fitWidth(),
    fitPage: () => state.viewer?.fitPage(),
  };
}

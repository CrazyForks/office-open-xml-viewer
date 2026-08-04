import { writable, type Readable } from 'svelte/store';
import type { Action } from 'svelte/action';
import {
  mountOfficeViewer,
  type OfficeFormat,
  type OfficeViewerByFormat,
  type OfficeViewerConfig,
  type OfficeViewerTargetByFormat,
} from '@ooxml-framework-examples/shared';

export type OfficeViewerStatus = 'idle' | 'loading' | 'ready' | 'error';

export interface OfficeViewerAction<F extends OfficeFormat> {
  action: Action<OfficeViewerTargetByFormat[F], OfficeViewerConfig<F>>;
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

export function createOfficeViewer<F extends OfficeFormat>(): OfficeViewerAction<F> {
  const status = writable<OfficeViewerStatus>('idle');
  const error = writable<Error | null>(null);
  let node: OfficeViewerTargetByFormat[F] | null = null;
  let config: OfficeViewerConfig<F> | null = null;
  let viewer: OfficeViewerByFormat[F] | null = null;
  let generation = 0;

  function destroyViewer(): void {
    generation += 1;
    viewer?.destroy();
    viewer = null;
  }

  function mount(): void {
    destroyViewer();
    if (!node || !config) {
      status.set('idle');
      return;
    }

    const currentGeneration = generation;
    status.set('loading');
    error.set(null);
    void mountOfficeViewer({ ...config, target: node }).then((nextViewer) => {
      if (generation !== currentGeneration) {
        nextViewer.destroy();
        return;
      }
      viewer = nextViewer;
      status.set('ready');
    }).catch((reason: unknown) => {
      if (generation !== currentGeneration) return;
      error.set(reason instanceof Error ? reason : new Error(String(reason)));
      status.set('error');
    });
  }

  const action: Action<OfficeViewerTargetByFormat[F], OfficeViewerConfig<F>> = (element, initialConfig) => {
    node = element;
    config = initialConfig;
    mount();

    return {
      update(nextConfig) {
        config = nextConfig;
        mount();
      },
      destroy() {
        destroyViewer();
        node = null;
        config = null;
        status.set('idle');
      },
    };
  };

  return {
    action,
    status,
    error,
    reload: mount,
    getScale: () => viewer?.getScale(),
    setScale: (scale) => viewer?.setScale(scale),
    zoomIn: () => viewer?.zoomIn(),
    zoomOut: () => viewer?.zoomOut(),
    fitWidth: () => viewer?.fitWidth(),
    fitPage: () => viewer?.fitPage(),
  };
}

import { createEffect, createSignal, onCleanup, type Accessor } from 'solid-js';

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

export interface CreateOfficeViewerOptions<V extends OfficeViewerHandle> {
  target: Accessor<HTMLElement | null>;
  source: Accessor<OfficeSource>;
  createViewer: Accessor<OfficeViewerFactory<V>>;
}

export function createOfficeViewer<V extends OfficeViewerHandle>(config: CreateOfficeViewerOptions<V>) {
  const [status, setStatus] = createSignal<OfficeViewerStatus>('idle');
  const [error, setError] = createSignal<Error | null>(null);
  const [reloadVersion, setReloadVersion] = createSignal(0);
  const [viewer, setViewer] = createSignal<V | null>(null);

  createEffect(() => {
    reloadVersion();
    const container = config.target();
    if (!container) {
      setStatus('idle');
      return;
    }

    const controller = new AbortController();
    const createViewer = config.createViewer();
    const source = config.source();
    setStatus('loading');
    setError(null);

    Promise.resolve(createViewer(container))
      .then(async (nextViewer) => {
        if (controller.signal.aborted) {
          nextViewer.destroy();
          return null;
        }
        try {
          await nextViewer.load(source);
          return nextViewer;
        } catch (reason) {
          nextViewer.destroy();
          throw reason;
        }
      })
      .then((nextViewer) => {
        if (!nextViewer || controller.signal.aborted) {
          nextViewer?.destroy();
          return;
        }
        setViewer(() => nextViewer);
        setStatus('ready');
      })
      .catch((reason: unknown) => {
        if (controller.signal.aborted) return;
        setError(reason instanceof Error ? reason : new Error(String(reason)));
        setStatus('error');
      });

    onCleanup(() => {
      controller.abort();
      const currentViewer = viewer();
      setViewer(null);
      currentViewer?.destroy();
    });
  });

  return {
    status,
    error,
    reload: () => setReloadVersion((version) => version + 1),
    getScale: () => viewer()?.getScale(),
    setScale: (scale: number) => viewer()?.setScale(scale),
    zoomIn: () => viewer()?.zoomIn(),
    zoomOut: () => viewer()?.zoomOut(),
    fitWidth: () => viewer()?.fitWidth(),
    fitPage: () => viewer()?.fitPage(),
  };
}

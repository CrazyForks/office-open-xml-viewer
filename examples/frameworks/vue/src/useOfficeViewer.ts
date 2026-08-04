import { readonly, ref, shallowRef, toValue, watchEffect, type MaybeRefOrGetter, type ShallowRef } from 'vue';

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

export interface UseOfficeViewerOptions<V extends OfficeViewerHandle> {
  target: Readonly<ShallowRef<HTMLElement | null>>;
  source: MaybeRefOrGetter<OfficeSource>;
  createViewer: OfficeViewerFactory<V>;
}

export function useOfficeViewer<V extends OfficeViewerHandle>(config: UseOfficeViewerOptions<V>) {
  const viewer = shallowRef<V | null>(null);
  const status = ref<OfficeViewerStatus>('idle');
  const error = shallowRef<Error | null>(null);
  const reloadVersion = ref(0);

  watchEffect((onCleanup) => {
    reloadVersion.value;
    const container = config.target.value;
    if (!container) {
      status.value = 'idle';
      return;
    }

    const controller = new AbortController();
    status.value = 'loading';
    error.value = null;

    Promise.resolve(config.createViewer(container))
      .then(async (nextViewer) => {
        if (controller.signal.aborted) {
          nextViewer.destroy();
          return null;
        }
        try {
          await nextViewer.load(toValue(config.source));
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
        viewer.value = nextViewer;
        status.value = 'ready';
      })
      .catch((reason: unknown) => {
        if (controller.signal.aborted) return;
        error.value = reason instanceof Error ? reason : new Error(String(reason));
        status.value = 'error';
      });

    onCleanup(() => {
      controller.abort();
      const currentViewer = viewer.value;
      viewer.value = null;
      currentViewer?.destroy();
    });
  });

  return {
    status: readonly(status),
    error: readonly(error),
    reload: () => { reloadVersion.value += 1; },
    getScale: () => viewer.value?.getScale(),
    setScale: (scale: number) => viewer.value?.setScale(scale),
    zoomIn: () => viewer.value?.zoomIn(),
    zoomOut: () => viewer.value?.zoomOut(),
    fitWidth: () => viewer.value?.fitWidth(),
    fitPage: () => viewer.value?.fitPage(),
  };
}

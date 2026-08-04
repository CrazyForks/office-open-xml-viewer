import { readonly, ref, shallowRef, toValue, watchEffect, type MaybeRefOrGetter, type ShallowRef } from 'vue';
import {
  mountOfficeViewer,
  type OfficeFormat,
  type OfficeSource,
  type OfficeViewerByFormat,
  type OfficeViewerOptionsByFormat,
  type OfficeViewerTargetByFormat,
} from '@ooxml-framework-examples/shared';

export type OfficeViewerStatus = 'idle' | 'loading' | 'ready' | 'error';

export interface UseOfficeViewerOptions<F extends OfficeFormat> {
  target: Readonly<ShallowRef<OfficeViewerTargetByFormat[F] | null>>;
  format: MaybeRefOrGetter<F>;
  source: MaybeRefOrGetter<OfficeSource>;
  viewerOptions?: MaybeRefOrGetter<OfficeViewerOptionsByFormat[F] | undefined>;
}

export function useOfficeViewer<F extends OfficeFormat>(config: UseOfficeViewerOptions<F>) {
  const viewer = shallowRef<OfficeViewerByFormat[F] | null>(null);
  const status = ref<OfficeViewerStatus>('idle');
  const error = shallowRef<Error | null>(null);
  const reloadVersion = ref(0);

  watchEffect((onCleanup) => {
    reloadVersion.value;
    const target = config.target.value;
    if (!target) {
      status.value = 'idle';
      return;
    }

    let active = true;
    let mountedViewer: OfficeViewerByFormat[F] | null = null;
    status.value = 'loading';
    error.value = null;

    void mountOfficeViewer({
      format: toValue(config.format),
      target,
      source: toValue(config.source),
      options: config.viewerOptions ? toValue(config.viewerOptions) : undefined,
    }).then((nextViewer) => {
      if (!active) {
        nextViewer.destroy();
        return;
      }
      mountedViewer = nextViewer;
      viewer.value = nextViewer;
      status.value = 'ready';
    }).catch((reason: unknown) => {
      if (!active) return;
      error.value = reason instanceof Error ? reason : new Error(String(reason));
      status.value = 'error';
    });

    onCleanup(() => {
      active = false;
      if (viewer.value === mountedViewer) viewer.value = null;
      mountedViewer?.destroy();
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

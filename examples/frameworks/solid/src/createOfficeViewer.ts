import { createEffect, createSignal, onCleanup, type Accessor } from 'solid-js';
import {
  mountOfficeViewer,
  type OfficeFormat,
  type OfficeSource,
  type OfficeViewerByFormat,
  type OfficeViewerOptionsByFormat,
  type OfficeViewerTargetByFormat,
} from '@ooxml-framework-examples/shared';

export type OfficeViewerStatus = 'idle' | 'loading' | 'ready' | 'error';
type MaybeAccessor<T> = T | Accessor<T>;

export interface CreateOfficeViewerOptions<F extends OfficeFormat> {
  target: Accessor<OfficeViewerTargetByFormat[F] | null>;
  format: MaybeAccessor<F>;
  source: MaybeAccessor<OfficeSource>;
  viewerOptions?: MaybeAccessor<OfficeViewerOptionsByFormat[F] | undefined>;
}

function read<T>(value: MaybeAccessor<T>): T {
  return typeof value === 'function' ? (value as Accessor<T>)() : value;
}

export function createOfficeViewer<F extends OfficeFormat>(config: CreateOfficeViewerOptions<F>) {
  const [status, setStatus] = createSignal<OfficeViewerStatus>('idle');
  const [error, setError] = createSignal<Error | null>(null);
  const [reloadVersion, setReloadVersion] = createSignal(0);
  let viewer: OfficeViewerByFormat[F] | null = null;

  createEffect(() => {
    reloadVersion();
    const target = config.target();
    if (!target) {
      setStatus('idle');
      return;
    }

    let active = true;
    let mountedViewer: OfficeViewerByFormat[F] | null = null;
    setStatus('loading');
    setError(null);

    void mountOfficeViewer({
      format: read(config.format),
      target,
      source: read(config.source),
      options: config.viewerOptions === undefined ? undefined : read(config.viewerOptions),
    }).then((nextViewer) => {
      if (!active) {
        nextViewer.destroy();
        return;
      }
      mountedViewer = nextViewer;
      viewer = nextViewer;
      setStatus('ready');
    }).catch((reason: unknown) => {
      if (!active) return;
      setError(reason instanceof Error ? reason : new Error(String(reason)));
      setStatus('error');
    });

    onCleanup(() => {
      active = false;
      if (viewer === mountedViewer) viewer = null;
      mountedViewer?.destroy();
    });
  });

  return {
    status,
    error,
    reload: () => setReloadVersion((version) => version + 1),
    getScale: () => viewer?.getScale(),
    setScale: (scale: number) => viewer?.setScale(scale),
    zoomIn: () => viewer?.zoomIn(),
    zoomOut: () => viewer?.zoomOut(),
    fitWidth: () => viewer?.fitWidth(),
    fitPage: () => viewer?.fitPage(),
  };
}

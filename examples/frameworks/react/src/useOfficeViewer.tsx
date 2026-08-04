import {
  useCallback,
  useEffect,
  useRef,
  useState,
  type HTMLAttributes,
  type ReactElement,
} from 'react';

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
  source: OfficeSource;
  /** Keep this factory referentially stable with useCallback. */
  createViewer: OfficeViewerFactory<V>;
}

export interface UseOfficeViewerResult {
  status: OfficeViewerStatus;
  error: Error | null;
  renderOfficeViewer: (props?: Omit<HTMLAttributes<HTMLDivElement>, 'ref'>) => ReactElement;
  reload: () => void;
  getScale: () => number | undefined;
  setScale: (scale: number) => void;
  zoomIn: () => void;
  zoomOut: () => void;
  fitWidth: () => void;
  fitPage: () => void;
}

export function useOfficeViewer<V extends OfficeViewerHandle>({
  source,
  createViewer,
}: UseOfficeViewerOptions<V>): UseOfficeViewerResult {
  const mountRef = useRef<HTMLDivElement>(null);
  const viewerRef = useRef<V | null>(null);
  const [status, setStatus] = useState<OfficeViewerStatus>('idle');
  const [error, setError] = useState<Error | null>(null);
  const [reloadVersion, setReloadVersion] = useState(0);

  const renderOfficeViewer = useCallback(
    (props: Omit<HTMLAttributes<HTMLDivElement>, 'ref'> = {}) => <div {...props} ref={mountRef} />,
    [],
  );

  useEffect(() => {
    const container = mountRef.current;
    if (!container) {
      setStatus('idle');
      return;
    }

    const controller = new AbortController();
    setStatus('loading');
    setError(null);

    Promise.resolve(createViewer(container))
      .then(async (viewer) => {
        if (controller.signal.aborted) {
          viewer.destroy();
          return null;
        }
        try {
          await viewer.load(source);
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
        viewerRef.current = viewer;
        setStatus('ready');
      })
      .catch((reason: unknown) => {
        if (controller.signal.aborted) return;
        setError(reason instanceof Error ? reason : new Error(String(reason)));
        setStatus('error');
      });

    return () => {
      controller.abort();
      const viewer = viewerRef.current;
      viewerRef.current = null;
      viewer?.destroy();
    };
  }, [createViewer, reloadVersion, source]);

  const reload = useCallback(() => setReloadVersion((version) => version + 1), []);
  const getScale = useCallback(() => viewerRef.current?.getScale(), []);
  const setScale = useCallback((scale: number) => viewerRef.current?.setScale(scale), []);
  const zoomIn = useCallback(() => viewerRef.current?.zoomIn(), []);
  const zoomOut = useCallback(() => viewerRef.current?.zoomOut(), []);
  const fitWidth = useCallback(() => viewerRef.current?.fitWidth(), []);
  const fitPage = useCallback(() => viewerRef.current?.fitPage(), []);

  return {
    status,
    error,
    renderOfficeViewer,
    reload,
    getScale,
    setScale,
    zoomIn,
    zoomOut,
    fitWidth,
    fitPage,
  };
}

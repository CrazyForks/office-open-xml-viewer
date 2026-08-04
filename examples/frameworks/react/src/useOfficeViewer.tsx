import {
  useCallback,
  useEffect,
  useRef,
  useState,
  type HTMLAttributes,
  type ReactElement,
} from 'react';
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
  format: F;
  source: OfficeSource;
  /** Keep this object referentially stable (for example, with useMemo). */
  viewerOptions?: OfficeViewerOptionsByFormat[F];
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

export function useOfficeViewer<F extends OfficeFormat>({
  format,
  source,
  viewerOptions,
}: UseOfficeViewerOptions<F>): UseOfficeViewerResult {
  const mountRef = useRef<HTMLDivElement>(null);
  const viewerRef = useRef<OfficeViewerByFormat[F] | null>(null);
  const [status, setStatus] = useState<OfficeViewerStatus>('idle');
  const [error, setError] = useState<Error | null>(null);
  const [reloadVersion, setReloadVersion] = useState(0);

  const renderOfficeViewer = useCallback(
    (props: Omit<HTMLAttributes<HTMLDivElement>, 'ref'> = {}) => <div {...props} ref={mountRef} />,
    [],
  );

  useEffect(() => {
    const mountPoint = mountRef.current;
    if (!mountPoint) {
      setStatus('idle');
      return;
    }

    const canvasTarget = format === 'xlsx'
      ? null
      : mountPoint.appendChild(document.createElement('canvas'));
    const target = (canvasTarget ?? mountPoint) as OfficeViewerTargetByFormat[F];
    let active = true;
    let mountedViewer: OfficeViewerByFormat[F] | null = null;
    setStatus('loading');
    setError(null);

    void mountOfficeViewer({ format, target, source, options: viewerOptions })
      .then((viewer) => {
        if (!active) {
          viewer.destroy();
          canvasTarget?.remove();
          return;
        }
        mountedViewer = viewer;
        viewerRef.current = viewer;
        setStatus('ready');
      })
      .catch((reason: unknown) => {
        canvasTarget?.remove();
        if (!active) return;
        setError(reason instanceof Error ? reason : new Error(String(reason)));
        setStatus('error');
      });

    return () => {
      active = false;
      if (viewerRef.current === mountedViewer) viewerRef.current = null;
      mountedViewer?.destroy();
      canvasTarget?.remove();
    };
  }, [format, reloadVersion, source, viewerOptions]);

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

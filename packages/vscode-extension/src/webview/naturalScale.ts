/**
 * The VS Code preview's default document scale.
 *
 * `1` is the viewers' absolute natural scale: authored points/EMUs map to their
 * normal CSS-pixel size. Unlike fit-width, this keeps equal font sizes visually
 * equal across portrait and landscape documents and across editor pane widths.
 */
export const NATURAL_DOCUMENT_SCALE = 1;

interface LoadableScrollViewer {
  setScale(scale: number): void;
  load(source: string | ArrayBuffer): Promise<void>;
}

/**
 * Set the absolute scale before loading. DocxScrollViewer and PptxScrollViewer
 * intentionally latch a pre-load setScale call and apply it after establishing
 * their internal base geometry, overriding the default container-width fit.
 */
export async function loadAtNaturalScale(
  viewer: LoadableScrollViewer,
  source: string | ArrayBuffer,
): Promise<void> {
  viewer.setScale(NATURAL_DOCUMENT_SCALE);
  await viewer.load(source);
}

/**
 * Load without latching an absolute scale. The ScrollViewer therefore uses its
 * container-derived fit width, including its built-in left/right desk gutters.
 */
export async function loadAtContainerFit(
  viewer: Pick<LoadableScrollViewer, 'load'>,
  source: string | ArrayBuffer,
): Promise<void> {
  await viewer.load(source);
}

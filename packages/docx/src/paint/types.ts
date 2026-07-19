import type {
  PaintResourceDescriptorKind,
  LayoutRect,
  PaintResourceKind,
  Matrix2DData,
  DrawingLayout,
  TextBoxLayout,
} from '../layout/types.js';
import type { ResolvedPaintResource } from './resource-session.js';
import type { TextRunPaintInfo } from './text-run-info.js';

export interface PaintPageOptions {
  readonly scale: number;
  readonly dpr: number;
}

export interface PaintCanvas2D {
  globalAlpha: number;
  fillStyle: string | CanvasGradient | CanvasPattern;
  strokeStyle: string | CanvasGradient | CanvasPattern;
  lineWidth: number;
  font: string;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  direction: CanvasDirection;
  letterSpacing: string;
  fontKerning: CanvasFontKerning;
  fillRect(x: number, y: number, width: number, height: number): void;
  strokeRect(x: number, y: number, width: number, height: number): void;
  setLineDash(segments: number[]): void;
  fillText(text: string, x: number, y: number): void;
  translate(x: number, y: number): void;
  rotate(angle: number): void;
  scale(x: number, y: number): void;
  transform(a: number, b: number, c: number, d: number, e: number, f: number): void;
  drawImage(image: CanvasImageSource, ...coordinates: number[]): void;
  save(): void;
  restore(): void;
  beginPath(): void;
  rect(x: number, y: number, width: number, height: number): void;
  clip(): void;
  moveTo(x: number, y: number): void;
  lineTo(x: number, y: number): void;
  stroke(): void;
  fill(): void;
}

export type CanvasPaintResourceHandler<K extends PaintResourceDescriptorKind> = (
  resource: ResolvedPaintResource<K>,
  bounds: LayoutRect,
  ctx: PaintCanvas2D,
) => void;

export type CanvasPaintResourceHandlers = Readonly<{
  [K in PaintResourceDescriptorKind]: CanvasPaintResourceHandler<K>;
}>;

export interface CanvasPaintResourcePainter {
  paint(
    resourceKey: string,
    kind: PaintResourceKind,
    bounds: LayoutRect,
    ctx: PaintCanvas2D,
  ): void;
}

export type DeferredCanvasPaint = () => void;
export type DeferredCanvasPaintWrapper = (paint: DeferredCanvasPaint) => DeferredCanvasPaint;

export interface CanvasPaintContext {
  readonly ctx: PaintCanvas2D;
  readonly scale: number;
  readonly dpr: number;
  readonly defaultTextColor?: string;
  readonly showTrackChanges?: boolean;
  readonly resources: CanvasPaintResourcePainter;
  /**
   * Affine map from the current retained point-space to final logical CSS pixels
   * (device-pixel scaling is represented separately by `dpr`). It mirrors every
   * production Canvas transform that can affect point geometry.
   */
  readonly pointToCss?: Matrix2DData;
  readonly onTextRun?: (run: TextRunPaintInfo) => void;
  readonly layoutTranslationPt?: Readonly<{ xPt: number; yPt: number }>;
  readonly textBoxVerticalMode?: 'vert' | 'vert270' | 'eaVert' | 'mongolianVert';
  /** Internal A6 traversal phase. Discovery walks retained placement and clip
   * frames without emitting ordinary table or paragraph ink. */
  readonly bodyDrawingPass?: 'normal' | 'discover-behind';
  /** Re-enters every Canvas frame owned by the recursive painter at the point
   * where a drawing was encountered. */
  readonly deferredPaintWrapper?: DeferredCanvasPaintWrapper;
  readonly deferBehindDrawing?: (
    drawing: DrawingLayout,
    textBoxes: readonly TextBoxLayout[],
    paint: () => void,
  ) => boolean;
  readonly deferFrontDrawing?: (
    drawing: DrawingLayout,
    textBoxes: readonly TextBoxLayout[],
    paint: () => void,
  ) => boolean;
}

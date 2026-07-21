import type { DocxDocumentModel } from './types.js';
import type { ResolvedLocalFontMetric } from './layout/text.js';
import { snapshotLocalMetrics } from './layout/text.js';
import type { MathLayoutResource } from './layout/resources.js';
import { productionDocumentInput } from './layout/resources.js';
import type { BodyLayoutKernel } from './layout/body-layout-kernel.js';
import type { LayoutServices } from './layout/types.js';
import type {
  MeasurementTextContext,
  VerticalGlyphMeasurementService,
} from './layout/measurement-capabilities.js';
import {
  createProductionBodyLayoutRuntime,
  type ProductionBodyModelGateway,
} from './layout/production-body-layout.js';
import { createProductionLayoutServices } from './layout/production-services.js';
import { verticalRunInkExtraPx } from './vertical-text.js';
import { attachBodyLayoutKernel } from './layout/runtime-state.js';

function createConcreteBodyLayoutKernel(
  doc: DocxDocumentModel,
  measureContext: MeasurementTextContext | null,
  resolvedLocalFonts: Readonly<Record<string, ResolvedLocalFontMetric>>,
  model: ProductionBodyModelGateway,
): BodyLayoutKernel {
  return createProductionBodyLayoutRuntime(
    doc,
    measureContext,
    resolvedLocalFonts,
    model,
  ).kernel;
}

export function createLayoutServices(
  doc: DocxDocumentModel,
  options: {
    readonly localMetrics?: Readonly<Record<string, ResolvedLocalFontMetric>>;
    readonly useGoogleFonts?: boolean;
    readonly mathResources?: readonly MathLayoutResource[];
    readonly mathDrawables?: ReadonlyMap<string, CanvasImageSource>;
    readonly measureContext?: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
    readonly embeddedFaces?: readonly FontFace[];
    readonly googleFaces?: readonly FontFace[];
  } = {},
): LayoutServices {
  const productionInput = productionDocumentInput(doc);
  doc = productionInput.document;
  const canvasContext = options.measureContext ?? (typeof OffscreenCanvas !== 'undefined'
    ? new OffscreenCanvas(1, 1).getContext('2d')
    : typeof document !== 'undefined'
      ? document.createElement('canvas').getContext('2d')
      : null);
  const context: MeasurementTextContext | null = canvasContext === null
    ? null
    : Object.freeze({
        get font() { return canvasContext.font; },
        set font(value: string) { canvasContext.font = value; },
        get letterSpacing() { return canvasContext.letterSpacing; },
        set letterSpacing(value: string) { canvasContext.letterSpacing = value; },
        get fontKerning() { return canvasContext.fontKerning; },
        set fontKerning(value: CanvasFontKerning) { canvasContext.fontKerning = value; },
        measureText(text: string) { return canvasContext.measureText(text); },
      });
  const hasDomVerticalProbe = canvasContext !== null
    && typeof HTMLCanvasElement !== 'undefined'
    && canvasContext.canvas instanceof HTMLCanvasElement
    && typeof document !== 'undefined';
  const verticalGlyphMeasurement: VerticalGlyphMeasurementService = Object.freeze({
    fingerprint: canvasContext === null
      ? 'vertical-glyph-measurement:deterministic-v1'
      : hasDomVerticalProbe
        ? 'vertical-glyph-measurement:dom-vert-probe-v1'
        : 'vertical-glyph-measurement:no-dom-vert-probe-v1',
    measureRunInkExtra(text: string): number {
      if (canvasContext === null) {
        throw new Error('Vertical glyph measurement requires a concrete text context');
      }
      return verticalRunInkExtraPx(canvasContext, text);
    },
  });
  const localMetrics = snapshotLocalMetrics(options.localMetrics);
  const services = createProductionLayoutServices(doc, {
    ...options,
    localMetrics,
    measureContext: context,
    verticalGlyphMeasurement,
    fontFamilyCharsets: productionInput.fontFamilyCharsets,
    mathOccurrences: productionInput.mathOccurrences,
    acquisitionInputs: productionInput.bodyModelGateway.acquisitionInputs,
  });
  attachBodyLayoutKernel(
    services,
    createConcreteBodyLayoutKernel(
      doc,
      context,
      localMetrics,
      productionInput.bodyModelGateway,
    ),
  );
  return services;
}

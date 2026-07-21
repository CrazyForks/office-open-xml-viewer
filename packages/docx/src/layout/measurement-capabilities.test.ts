import { describe, expect, it } from 'vitest';
import type {
  MeasurementTextContext,
  VerticalGlyphMeasurementService,
} from './measurement-capabilities.js';

function assertMeasurementSurface(context: MeasurementTextContext): void {
  context.font = context.font;
  context.letterSpacing = context.letterSpacing;
  context.fontKerning = context.fontKerning;
  context.measureText('measure');

  if (false) {
    // Acquisition receives text metrics, never a recoverable canvas or painter.
    // @ts-expect-error paint capability is outside the measurement contract
    context.canvas;
    // @ts-expect-error paint capability is outside the measurement contract
    context.fillText('paint', 0, 0);
    // @ts-expect-error paint capability is outside the measurement contract
    context.drawImage({} as CanvasImageSource, 0, 0);
    // @ts-expect-error mutable graphics state is outside the measurement contract
    context.save();
  }
}

describe('layout measurement capabilities', () => {
  it('expose metrics and an injected vertical measurement service without paint', () => {
    const context = {
      font: '10px serif',
      letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      measureText: (text: string) => ({ width: text.length * 5 }) as TextMetrics,
    } satisfies MeasurementTextContext;
    const vertical = {
      fingerprint: 'vertical:test',
      measureRunInkExtra: (_text: string) => 0,
    } satisfies VerticalGlyphMeasurementService;

    assertMeasurementSurface(context);
    expect(context.measureText('ab').width).toBe(10);
    expect(vertical.measureRunInkExtra('ab')).toBe(0);
  });
});

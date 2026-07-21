/** Canvas capabilities that retained layout may use synchronously.
 *
 * Vertical OpenType feature probing is deliberately excluded: it needs the
 * backing DOM canvas and is injected separately as a document-scoped service.
 */
export interface MeasurementTextContext {
  font: string;
  letterSpacing: string;
  fontKerning: CanvasFontKerning;
  measureText(text: string): TextMetrics;
}

/** Document-scoped vertical glyph measurement bound to the same concrete
 * context as {@link MeasurementTextContext}, without exposing its canvas or
 * paint methods to acquisition. Calls are synchronous because the underlying
 * OpenType feature probe mutates and restores font state in one stack frame. */
export interface VerticalGlyphMeasurementService {
  readonly fingerprint: string;
  measureRunInkExtra(text: string): number;
}

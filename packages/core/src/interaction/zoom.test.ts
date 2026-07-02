import { describe, it, expect } from 'vitest';
import { zoomStepScale } from './zoom.js';

/**
 * Ctrl/⌘ + wheel (and trackpad pinch, which the browser reports as a ctrl-wheel)
 * zoom. The old handler ignored `deltaY` magnitude and added a fixed ±0.1 per
 * event, so a trackpad pinch — which fires a high-frequency stream of small
 * wheel events — zoomed far too fast. `zoomStepScale` makes the step
 * exponential in `deltaY`, so the *total* zoom over a gesture is
 * `exp(-k·Σ deltaY)` and depends only on the total scroll distance, not on how
 * many events the OS splits it into.
 */
describe('zoomStepScale (ctrl/pinch zoom)', () => {
  it('scrolling up / pinching out (deltaY < 0) zooms in', () => {
    expect(zoomStepScale(1, -10)).toBeGreaterThan(1);
  });

  it('scrolling down / pinching in (deltaY > 0) zooms out', () => {
    expect(zoomStepScale(1, 10)).toBeLessThan(1);
  });

  it('honors deltaY magnitude (a bigger delta zooms more)', () => {
    const small = zoomStepScale(1, -2) - 1;
    const big = zoomStepScale(1, -20) - 1;
    expect(big).toBeGreaterThan(small);
  });

  it('is resolution-independent: two small events ≈ one event of their sum', () => {
    const twoSteps = zoomStepScale(zoomStepScale(1, -5), -5);
    const oneStep = zoomStepScale(1, -10);
    expect(twoSteps).toBeCloseTo(oneStep, 10);
  });

  it('is symmetric: zooming in then out by the same delta returns to start', () => {
    expect(zoomStepScale(zoomStepScale(1, -8), 8)).toBeCloseTo(1, 10);
  });

  it('scales relative to the current zoom (multiplicative, not additive)', () => {
    // Same delta from 200% must move proportionally more than from 100%.
    const from1 = zoomStepScale(1, -10) - 1;
    const from2 = zoomStepScale(2, -10) - 2;
    expect(from2).toBeCloseTo(from1 * 2, 10);
  });
});

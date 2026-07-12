/**
 * Vertical (tbRl) colon `：` / semicolon `；` form probe — ECMA-376 §17.6.20 +
 * UAX#50 §5, issue #969 follow-up.
 *
 * Both `：`(FF1A) and `；`(FF1B) are vo=Tr. The Unicode vertical presentation
 * forms are FE13 (︓) and FE14 (︔), but MOST render fonts (Hiragino Mincho ProN,
 * which macOS substitutes for the sample's Yu Mincho) do NOT contain FE13/FE14,
 * and a Canvas cannot reach the font's `vert` OpenType feature. Unconditionally
 * substituting the FE code point therefore reaches the system fallback cascade,
 * which supplies a DIFFERENT font's glyph positioned wrong (measured: FE13/FE14
 * ink lands ~0.25em to the RIGHT of the column centre in both skia AND Chrome).
 *
 * Word ground truth (sample-47 PDF, macOS Quartz):
 *   • `：` renders as two dots SIDE BY SIDE (horizontal), centred on the column —
 *     which is the base `：` (two vertically-stacked dots) ROTATED 90°, and is
 *     FE13's intrinsic design.
 *   • `；` renders as dot-over-comma VERTICAL (upright), centred — which is the
 *     base `；` drawn UPRIGHT, and is FE14's intrinsic design (NOT a rotation).
 *
 * So the correct, font-robust rendering is a GEOMETRIC fallback that reproduces
 * each vertical form's design directly: colon → rotate, semicolon → upright.
 * This probe measures the rendered ink through the REAL `drawVerticalRun` and
 * asserts each mark is (a) centred on the column centreline (cross-axis) and
 * (b) has the right orientation (colon wider-than-tall, semicolon taller-than-wide).
 *
 * CI-safe: gated on skia (devDependency) + Hiragino (macOS dev host).
 */
import { describe, it, expect } from 'vitest';
import { existsSync } from 'node:fs';
import { loadSkiaForTests, importForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, FontLibrary } = (skia ?? {}) as Skia;
const MINCHO = '/System/Library/Fonts/ヒラギノ明朝 ProN.ttc';
const haveFont = existsSync(MINCHO);
const vtMod = await importForTests(
  () => import('../../docx/src/vertical-text.ts'),
  'packages/docx/src/vertical-text.ts',
);
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

function ink(data: Uint8ClampedArray, i: number): number {
  return 255 - (0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2]);
}

describe.skipIf(!skia || !vtMod || !haveFont)('docx vertical colon/semicolon forms (§17.6.20 / UAX#50)', () => {
  const fontPx = 48;
  const W = 300, H = 300, baseline = 150, logX = 40, cellW = fontPx;
  const centerline = W - baseline; // column centreline in physical x

  /** Render "話X話" via drawVerticalRun and return the pixel buffer. */
  function render(mid: string): Uint8ClampedArray {
    const { drawVerticalRun } = vtMod as { drawVerticalRun: (...a: Any[]) => void };
    for (const fam of ['Yu Mincho', 'MS Mincho', 'Hiragino Mincho ProN']) FontLibrary.use(fam, [MINCHO]);
    const canvas = new Canvas(W, H);
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
    ctx.fillStyle = '#000';
    ctx.font = `${fontPx}px "MS Mincho", serif`;
    ctx.save();
    ctx.translate(W, 0); ctx.rotate(Math.PI / 2);
    drawVerticalRun(ctx, `話${mid}話`, logX, baseline, fontPx, 0);
    ctx.restore();
    return ctx.getImageData(0, 0, W, H).data;
  }

  /** Ink bbox + cross-axis centroid of the middle cell (2nd glyph, physical y band). */
  function midCell(data: Uint8ClampedArray): { cx: number; bw: number; bh: number } {
    const py0 = logX + cellW, py1 = logX + 2 * cellW; // 2nd cell along the column
    let sx = 0, sw = 0, x0 = W, x1 = 0, y0 = H, y1 = 0;
    for (let py = py0; py < py1; py++) for (let px = 0; px < W; px++) {
      const w = ink(data, (py * W + px) * 4);
      if (w > 40) { sx += px * w; sw += w; if (px < x0) x0 = px; if (px > x1) x1 = px; if (py < y0) y0 = py; if (py > y1) y1 = py; }
    }
    return { cx: sw > 0 ? sx / sw : NaN, bw: x1 - x0, bh: y1 - y0 };
  }

  it('colon ： is centred on the column and renders horizontal (wider than tall)', () => {
    const m = midCell(render('：'));
    // Centred on the column centreline (cross-axis), within 0.1em.
    expect(Math.abs(m.cx - centerline)).toBeLessThanOrEqual(0.1 * fontPx);
    // Two dots SIDE BY SIDE ⇒ cross-axis extent (bw) > along-column extent (bh).
    expect(m.bw).toBeGreaterThan(m.bh);
  });

  it('semicolon ； is centred on the column and renders upright (taller than wide)', () => {
    const m = midCell(render('；'));
    expect(Math.abs(m.cx - centerline)).toBeLessThanOrEqual(0.1 * fontPx);
    // Dot-over-comma stacked VERTICALLY ⇒ along-column extent (bh) > cross (bw).
    expect(m.bh).toBeGreaterThan(m.bw);
  });
});

import { mathToMathML, recolorSvg } from '@silurus/ooxml-core';
import type {
  MathLayoutResource,
  MathOccurrence,
  MathRenderer,
} from '../layout/types.js';

function svgToImage(svg: string): Promise<HTMLImageElement> {
  const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  const image = new Image();
  return new Promise((resolve, reject) => {
    image.onload = () => resolve(image);
    image.onerror = reject;
    image.src = url;
  });
}

export async function prepareBrowserMathResources(
  occurrences: readonly MathOccurrence[],
  math: MathRenderer,
) {
  if (occurrences.length === 0) return { records: [], drawables: new Map() };
  await math.loadMathJax();
  const records: MathLayoutResource[] = [];
  const drawables = new Map<string, CanvasImageSource>();
  const seen = new Set<string>();
  for (const occurrence of occurrences) {
    if (seen.has(occurrence.resourceKey)) {
      throw new Error(`Duplicate math occurrence: ${occurrence.resourceKey}`);
    }
    seen.add(occurrence.resourceKey);
    try {
      const output = await math.mathMLToSvg(mathToMathML(occurrence.nodes, occurrence.display));
      const image = await svgToImage(recolorSvg(output.svg, '#000000'));
      records.push({
        resourceKey: occurrence.resourceKey,
        widthEm: output.widthEm,
        ascentEm: output.ascentEm,
        descentEm: output.descentEm,
        diagnostics: [],
      });
      drawables.set(occurrence.resourceKey, image);
    } catch {
      records.push({
        resourceKey: occurrence.resourceKey,
        widthEm: 0,
        ascentEm: 0,
        descentEm: 0,
        available: false,
        diagnostics: [{
          code: 'UNSUPPORTED_FEATURE',
          severity: 'warning',
          message: 'Math conversion failed; using the deterministic text fallback',
        }],
      });
    }
  }
  return { records, drawables };
}

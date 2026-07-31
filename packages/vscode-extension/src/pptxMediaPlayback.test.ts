import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const bootstrap = readFileSync(
  new URL('./webview/bootstrap.ts', import.meta.url),
  'utf8',
);

describe('VS Code PPTX media playback wiring', () => {
  it('enables interactive media on the continuous-scroll viewer', () => {
    const pptxInitializer = bootstrap.match(
      /async function initPptx[\s\S]*?new PptxScrollViewer\([\s\S]*?\n  \}\);/,
    )?.[0];

    expect(pptxInitializer).toBeDefined();
    expect(pptxInitializer).toContain('enableMediaPlayback: true');
    expect(pptxInitializer).toContain('mediaOverscan: 1');
  });
});

import { describe, expect, it } from 'vitest';
import { build } from 'esbuild';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));

describe('VS Code webview math engine bundle', () => {
  it('retains the MathJax STIX2 side-effect entry used by DOCX, XLSX, and PPTX', async () => {
    const result = await build({
      stdin: {
        contents: "import '@silurus/ooxml-core/mathjax-stix2';",
        resolveDir: HERE,
        loader: 'js',
      },
      bundle: true,
      write: false,
      format: 'iife',
      platform: 'browser',
      target: 'es2020',
      logLevel: 'silent',
    });

    expect(result.warnings).toEqual([]);
    expect(result.outputFiles[0]?.text).toContain('globalThis.__ooxmlStix2');
  });
});

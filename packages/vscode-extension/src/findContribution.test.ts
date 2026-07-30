import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

interface Manifest {
  contributes: {
    commands: Array<{ command: string }>;
    keybindings: Array<{ command: string; key: string; mac?: string; when?: string }>;
  };
}

const manifest = JSON.parse(
  readFileSync(new URL('../package.json', import.meta.url), 'utf8'),
) as Manifest;

describe('VS Code find contribution', () => {
  it('routes Ctrl/Cmd+F only while an OOXML custom editor is active', () => {
    expect(manifest.contributes.commands).toContainEqual(
      expect.objectContaining({ command: 'ooxmlViewer.find' }),
    );
    expect(manifest.contributes.keybindings).toContainEqual({
      command: 'ooxmlViewer.find',
      key: 'ctrl+f',
      mac: 'cmd+f',
      when:
        'activeCustomEditorId == ooxmlViewer.docxEditor || activeCustomEditorId == ooxmlViewer.xlsxEditor || activeCustomEditorId == ooxmlViewer.pptxEditor',
    });
  });
});

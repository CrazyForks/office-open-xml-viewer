import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const source = readFileSync(new URL('./pages/try.astro', import.meta.url), 'utf8');

describe('Try Yours parsing progress', () => {
  it('uses concise, functional copy for the file workflow', () => {
    expect(source).toContain('Drop in a file.<br />See how it renders.');
    expect(source).toContain('Choose another file');
    expect(source).not.toContain('Inspect freely');
    expect(source).not.toContain('Verify in source');
    expect(source).not.toContain('Replace file ↗');
  });

  it('uses a semantic file-picker button without decorative step numbers', () => {
    expect(source).toContain('<button class="dropzone" id="dropzone" type="button">');
    expect(source).toContain("dz.addEventListener('click', () => input.click())");
    expect(source).not.toContain('dz-index');
    expect(source).not.toContain('02 / Privacy');
  });

  it('shows an accessible progress circle in the preview while renderFile is pending', () => {
    expect(source).toContain('id="stage-progress" role="status" aria-live="polite" hidden');
    expect(source).toContain('class="try-progress-circle" aria-hidden="true"');
    expect(source).toMatch(/stageProgress\.hidden = false;[\s\S]*await renderFile\(stage, file\)/);
  });

  it('hides the progress UI on both the current render success and failure paths', () => {
    expect(source.match(/stageProgress\.hidden = true;/g)).toHaveLength(2);
  });
});

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const source = readFileSync(new URL('./pages/try.astro', import.meta.url), 'utf8');

describe('Try Yours parsing progress', () => {
  it('shows an accessible progress circle in the preview while renderFile is pending', () => {
    expect(source).toContain('id="stage-progress" role="status" aria-live="polite" hidden');
    expect(source).toContain('class="try-progress-circle" aria-hidden="true"');
    expect(source).toMatch(/stageProgress\.hidden = false;[\s\S]*await renderFile\(stage, file\)/);
  });

  it('hides the progress UI on both the current render success and failure paths', () => {
    expect(source.match(/stageProgress\.hidden = true;/g)).toHaveLength(2);
  });
});

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const component = readFileSync(new URL('./components/ApiReference.astro', import.meta.url), 'utf8');
const readme = readFileSync(new URL('../../README.md', import.meta.url), 'utf8');

describe('public error documentation', () => {
  it('shows the error taxonomy on every format API page', () => {
    expect(component).toContain('<h4>Error handling</h4>');
    expect(component).toContain('<code>OoxmlError</code>');
    expect(component).toContain('<code>OoxmlResourceLimitError</code>');
    expect(component).toContain('<code>OoxmlDecodedImageLimitError</code>');
    expect(component).toContain('<code>parser-crashed</code>');
    expect(component).toContain('Do not branch on message text.');
  });

  it('makes the callback-versus-Promise behavior explicit in the README', () => {
    expect(readme).toContain('## Error handling');
    expect(readme).toContain('With `onError`, that failure is delivered to the callback and');
    expect(readme).toContain('`viewer.load()` resolves');
    expect(readme).toContain('messages are diagnostic text, not a programmatic API');
  });
});

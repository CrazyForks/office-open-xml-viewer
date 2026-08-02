import { describe, expect, it } from 'vitest';
import { codeThemes } from './lib/code-theme';

describe('official-site syntax themes', () => {
  it('uses the selected One pair in light and dark mode', () => {
    expect(codeThemes).toEqual({
      light: 'one-light',
      dark: 'one-dark-pro',
    });
  });
});

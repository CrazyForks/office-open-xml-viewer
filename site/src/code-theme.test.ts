import { describe, expect, it } from 'vitest';
import { codeThemes } from './lib/code-theme';

describe('official-site syntax themes', () => {
  it('uses the selected Rosé Pine pair in light and dark mode', () => {
    expect(codeThemes).toEqual({
      light: 'rose-pine-dawn',
      dark: 'rose-pine-moon',
    });
  });
});

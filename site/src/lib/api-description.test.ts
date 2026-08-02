import { describe, expect, it } from 'vitest';
import { apiReference } from './api-reference.js';
import { descriptionTokens } from './api-description.js';

describe('API description formatting', () => {
  it('turns backtick-delimited text into inline-code tokens', () => {
    expect(descriptionTokens({ desc: 'Set `false`, then call `load()`.' })).toEqual([
      { text: 'Set ', code: false, emphasized: false },
      { text: 'false', code: true, emphasized: false },
      { text: ', then call ', code: false, emphasized: false },
      { text: 'load()', code: true, emphasized: false },
      { text: '.', code: false, emphasized: false },
    ]);
  });

  it('preserves semantic emphasis around inline code', () => {
    expect(descriptionTokens({
      desc: 'Before. `{ 0, 0 }` means not loaded. After.',
      emphasis: '`{ 0, 0 }` means not loaded.',
    })).toEqual([
      { text: 'Before. ', code: false, emphasized: false },
      { text: '{ 0, 0 }', code: true, emphasized: true },
      { text: ' means not loaded.', code: false, emphasized: true },
      { text: ' After.', code: false, emphasized: false },
    ]);
  });

  it('rejects malformed authoring instead of exposing unmatched backticks', () => {
    expect(() => descriptionTokens({ desc: 'Call `load().' })).toThrow('Unmatched backtick');
  });

  it('formats every published API description without leaving delimiter text', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        for (const item of [...(apiClass.options ?? []), ...apiClass.methods]) {
          const tokens = descriptionTokens(item);
          expect(tokens.some(({ text }) => text.includes('`')), `${apiClass.name}: ${item.desc}`).toBe(false);
        }
      }
    }
  });
});

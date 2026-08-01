import { describe, expect, it } from 'vitest';
import {
  cappedAdd,
  jsonStringBytes,
  measureStructuralJson,
  utf8Bytes,
} from './resource-measurement.js';

const serializedBytes = (value: unknown): number =>
  new TextEncoder().encode(JSON.stringify(value)).byteLength;

describe('shared structural JSON resource measurement', () => {
  it('matches JSON.stringify for controls, unicode, and lone surrogates', () => {
    for (const value of [
      '',
      'quote" slash\\ \b\t\n\f\r \u0000\u000b\u001f',
      'é😀',
      '\ud800',
      '\udc00',
      '\ud800x\udc00',
    ]) {
      expect(jsonStringBytes(value)).toBe(serializedBytes(value));
      expect(utf8Bytes(value)).toBe(new TextEncoder().encode(value).byteLength);
    }
  });

  it('measures arrays and ordinary objects exactly while excluding keys from string values', () => {
    const value = {
      'ignored-é😀': 'kept-é😀',
      controls: '\u0000\n',
      nested: [undefined, () => 1, Symbol('x'), Number.NaN, Infinity, -Infinity, -0, 'x'],
      omitted: undefined,
    };
    const measured = measureStructuralJson(value);
    expect(measured.jsonBytes).toBe(serializedBytes(value));
    expect(measured.stringValueUtf8Bytes).toBe(
      utf8Bytes('kept-é😀') + utf8Bytes('\u0000\n') + utf8Bytes('x'),
    );
  });

  it('uses exact inclusive caps, canonical +1, and safe overflow checks', () => {
    expect(cappedAdd(4, 6, 10)).toBe(10);
    expect(cappedAdd(5, 6, 10)).toBe(11);
    expect(cappedAdd(11, 0, 10)).toBe(11);
    expect(jsonStringBytes('x', 2)).toBe(3);
    expect(measureStructuralJson(['long', 'still traversed'], 4).jsonBytes).toBe(5);
    expect(() => cappedAdd(Number.MAX_SAFE_INTEGER + 1, 0, 10)).toThrow(/safe integer/);
    expect(() => cappedAdd(0, 0, -1)).toThrow(/safe integer/);
    expect(() => utf8Bytes('', -1)).toThrow(/safe integer/);
    expect(() => measureStructuralJson(undefined, -1)).toThrow(/safe integer/);
    expect(cappedAdd(Number.MAX_SAFE_INTEGER, 1, Number.MAX_SAFE_INTEGER))
      .toBe(Number.MAX_SAFE_INTEGER);
  });

  it('matches JSON.stringify rejection of bigint values', () => {
    expect(() => measureStructuralJson({ value: 1n })).toThrow(TypeError);
  });
});

import { describe, expect, it } from 'vitest';
import type { DocxDocumentModel, SectionProps } from '../types.js';
import {
  isAllRotatedVerticalTextDirection,
  isVerticalSection,
  isVerticalTextDirection,
  physicalLayoutSection,
  verticalLayoutDoc,
  verticalLayoutSection,
} from './section-orientation.js';

const section = (textDirection: string | null): SectionProps => ({
  pageWidth: 600,
  pageHeight: 800,
  marginTop: 10,
  marginRight: 20,
  marginBottom: 30,
  marginLeft: 40,
  headerDistance: 5,
  footerDistance: 6,
  textDirection,
  columns: null,
} as SectionProps);

const documentWith = (value: SectionProps): DocxDocumentModel => ({
  section: value,
  body: [],
  headers: { default: null, first: null, even: null },
  footers: { default: null, first: null, even: null },
} as unknown as DocxDocumentModel);

describe('section orientation acquisition', () => {
  it('recognizes the complete Transitional vertical direction set', () => {
    for (const direction of ['tbRl', 'tbRlV', 'tbLrV', 'btLr']) {
      expect(isVerticalTextDirection(direction)).toBe(true);
      expect(isVerticalSection(section(direction))).toBe(true);
    }
    for (const direction of ['lrTb', 'lrTbV', 'unknown', null, undefined]) {
      expect(isVerticalTextDirection(direction)).toBe(false);
    }
    expect(isAllRotatedVerticalTextDirection('btLr')).toBe(true);
    expect(isAllRotatedVerticalTextDirection('tbRl')).toBe(false);
  });

  it('round-trips physical and logical page geometry without losing section facts', () => {
    const physical = section('tbRl');
    const logical = verticalLayoutSection(physical);

    expect(logical).toMatchObject({
      pageWidth: 800,
      pageHeight: 600,
      marginTop: 20,
      marginRight: 30,
      marginBottom: 40,
      marginLeft: 10,
      headerDistance: 5,
      footerDistance: 6,
      textDirection: 'tbRl',
    });
    expect(physicalLayoutSection(logical)).toEqual(physical);
  });

  it('preserves horizontal document identity and replaces only a vertical body section', () => {
    const horizontal = documentWith(section(null));
    const vertical = documentWith(section('tbRl'));

    expect(verticalLayoutDoc(horizontal)).toBe(horizontal);
    const logical = verticalLayoutDoc(vertical);
    expect(logical).not.toBe(vertical);
    expect(logical.body).toBe(vertical.body);
    expect(logical.section).toEqual(verticalLayoutSection(vertical.section));
  });
});

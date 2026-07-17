import { describe, expect, it } from 'vitest';
import type { LayoutServices } from './types.js';
import type { BodyLayoutKernel } from './body-layout-kernel.js';
import {
  attachBodyLayoutKernel,
  bodyLayoutKernelOf,
  createFieldAcquisitionServicesView,
} from './runtime-state.js';

const services = (): LayoutServices => Object.freeze({
  text: { fingerprint: 'text' } as LayoutServices['text'],
  images: { fingerprint: 'images' } as LayoutServices['images'],
  math: { fingerprint: 'math' } as LayoutServices['math'],
});

describe('private body layout kernel ownership', () => {
  it('follows legitimate service composition without becoming a public field', () => {
    const first = services();
    const second = services();
    const kernel = { openBodyLayoutSession() { throw new Error('unused'); } } as BodyLayoutKernel;

    attachBodyLayoutKernel(first, kernel);
    const fieldView = createFieldAcquisitionServicesView(first, { totalPages: 2 });
    const composed = Object.freeze({ ...first, text: first.text });

    expect(bodyLayoutKernelOf(first)).toBe(kernel);
    expect(bodyLayoutKernelOf(fieldView)).toBe(kernel);
    expect(bodyLayoutKernelOf(composed)).toBe(kernel);
    expect(bodyLayoutKernelOf(second)).toBeUndefined();
    expect(() => attachBodyLayoutKernel(first, kernel)).toThrow(/already attached/);
    expect(() => attachBodyLayoutKernel(composed, {
      openBodyLayoutSession() { throw new Error('foreign'); },
    } as BodyLayoutKernel)).toThrow(/already attached|foreign/i);
    expect(() => createFieldAcquisitionServicesView(second, { totalPages: 2 }))
      .toThrow(/kernel is not attached/i);
    attachBodyLayoutKernel(second, {
      openBodyLayoutSession() { throw new Error('foreign'); },
    } as BodyLayoutKernel);
    expect(() => bodyLayoutKernelOf(Object.freeze({
      text: first.text, images: second.images, math: second.math,
    }))).toThrow(/foreign runtime owners/i);
    expect(() => bodyLayoutKernelOf(Object.freeze({
      text: { fingerprint: 'unowned:text' } as LayoutServices['text'],
      images: { fingerprint: 'unowned:images' } as LayoutServices['images'],
      math: first.math,
    }))).toThrow(/missing service lineage/i);
    expect(first).not.toHaveProperty('bodyLayoutKernel');
    expect(Object.keys(first)).toEqual(['text', 'images', 'math']);
  });
});

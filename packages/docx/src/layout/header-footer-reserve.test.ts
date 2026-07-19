import { describe, expect, it } from 'vitest';
import { headerFooterOverflowReservePt } from './header-footer-reserve.js';

describe('header/footer body reserve', () => {
  it('charges only overflow beyond the signed-margin allowance', () => {
    expect(headerFooterOverflowReservePt(30, 72, 36)).toBe(0);
    expect(headerFooterOverflowReservePt(48, 72, 36)).toBe(12);
    expect(headerFooterOverflowReservePt(120, -72, 36)).toBe(0);
  });
});

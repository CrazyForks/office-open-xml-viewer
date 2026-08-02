import { describe, expect, it } from 'vitest';
import { announcements, latestAnnouncements } from './announcements';

describe('announcements', () => {
  it('keeps stable, unique routes in newest-first order', () => {
    expect(new Set(announcements.map(({ slug }) => slug)).size).toBe(announcements.length);
    expect(announcements.every(({ slug }) => /^[a-z0-9-]+$/.test(slug))).toBe(true);
    expect(announcements.map(({ date }) => date)).toEqual(
      [...announcements].map(({ date }) => date).sort().reverse(),
    );
  });

  it('uses the leading announcements on the home page', () => {
    expect(latestAnnouncements).toEqual(announcements.slice(0, 3));
  });
});

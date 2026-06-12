import type { MediaElement, Presentation } from './types';

/** Find the MIME type recorded for a media/poster zip path by scanning the
 *  slide elements (shared by the main-thread proxy and the render worker). */
export function findMimeTypeForPath(pres: Presentation, mediaPath: string): string {
  for (const slide of pres.slides) {
    for (const el of slide.elements) {
      if (el.type !== 'media') continue;
      const m = el as MediaElement;
      if (m.mediaPath === mediaPath) return m.mimeType;
      if (m.posterPath === mediaPath) return m.posterMimeType;
    }
  }
  return '';
}

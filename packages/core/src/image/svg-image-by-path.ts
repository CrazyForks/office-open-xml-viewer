// Path-keyed companion to `svg-image.ts` for the lazy, byte-on-demand image
// pipeline. Where `getCachedSvgImage` keys its cache on a self-contained
// `data:image/svg+xml;base64,…` URL, this variant keys on the embedded zip path
// (e.g. "word/media/image1.svg") and pulls the bytes lazily via a caller-
// supplied `fetchImage(path, mimeType)` — mirroring the pptx audio/video
// extraction pattern. The fetched bytes are wrapped in an object URL that this
// module owns and revokes on eviction (unlike a `data:` URL, an object URL is a
// live handle that leaks if never released).
//
// SVG is decoded via an `<img>` element rather than `createImageBitmap`,
// because `createImageBitmap` cannot rasterize SVG in every browser. The Promise
// is cached so concurrent first-renders dedupe; an HTMLImageElement holds no GPU
// resource, so eviction only needs to drop the entry and revoke its object URL.
// Bounded LRU.

const svgByPathCache = new Map<string, Promise<HTMLImageElement>>();
const urlByPath = new Map<string, string>(); // object URL owned here
const MAX = 256;

/**
 * Decode the SVG at `svgImagePath` to an `HTMLImageElement`, cached by path.
 * The bytes are fetched lazily through `fetchImage(path, mimeType)`; the
 * resulting object URL is owned by this module and revoked when the entry is
 * evicted. The returned image is drawable with `ctx.drawImage` exactly like an
 * ImageBitmap. Rejects if the SVG fails to load — callers should fall back to a
 * raster representation on rejection.
 */
export async function getCachedSvgImageByPath(
  svgImagePath: string,
  fetchImage: (path: string, mimeType: string) => Promise<Blob>,
): Promise<HTMLImageElement> {
  const hit = svgByPathCache.get(svgImagePath);
  if (hit) {
    // Refresh LRU position.
    svgByPathCache.delete(svgImagePath);
    svgByPathCache.set(svgImagePath, hit);
    return hit;
  }
  const p = (async () => {
    const blob = await fetchImage(svgImagePath, 'image/svg+xml');
    const url = URL.createObjectURL(blob);
    urlByPath.set(svgImagePath, url);
    const img = new Image();
    await new Promise<void>((resolve, reject) => {
      img.onload = () => resolve();
      img.onerror = () => reject(new Error(`svg load failed: ${svgImagePath}`));
      img.src = url;
    });
    return img;
  })();
  svgByPathCache.set(svgImagePath, p);
  if (svgByPathCache.size > MAX) {
    const oldest = svgByPathCache.keys().next().value as string;
    svgByPathCache.delete(oldest);
    const u = urlByPath.get(oldest);
    if (u) {
      URL.revokeObjectURL(u);
      urlByPath.delete(oldest);
    }
  }
  return p;
}

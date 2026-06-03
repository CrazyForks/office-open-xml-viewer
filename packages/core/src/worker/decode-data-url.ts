/**
 * Decode a `data:` URL into its raw bytes, or `null` for any non-data URL.
 *
 * The format workers receive the WASM binary either as a real URL (fetched by
 * `wasm-bindgen`'s `init`) or, in some bundler/inline-worker setups, as a
 * base64 `data:` URL. `wasm-bindgen`'s `init` cannot fetch a `data:` URL, so
 * the worker decodes it to an ArrayBuffer first. Shared by all three package
 * workers to keep the decoding identical.
 */
export function decodeDataUrl(url: string): ArrayBuffer | null {
  if (!url.startsWith('data:')) return null;
  const comma = url.indexOf(',');
  if (comma === -1) return null;
  const binary = atob(url.slice(comma + 1));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

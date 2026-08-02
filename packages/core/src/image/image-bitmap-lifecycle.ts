/** Release browser-owned ImageBitmap storage when the runtime exposes close().
 * Node canvas shims may return a structurally drawable native object without a
 * close method; those backends release their native storage through GC. */
export function closeImageBitmapIfSupported(bitmap: ImageBitmap): void {
  const close = (bitmap as ImageBitmap & { close?: unknown }).close;
  if (typeof close === 'function') close.call(bitmap);
}

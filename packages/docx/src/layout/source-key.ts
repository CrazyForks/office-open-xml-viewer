import type { SourceRef } from './types.js';

export function sourceKey(source: SourceRef): string {
  return `${source.story}:${encodeURIComponent(source.storyInstance)}:${source.path.join('.')}`;
}

export function bodyOccurrenceKey(
  source: SourceRef,
  flowDomainId: string,
  fragmentStartKey: string,
): string {
  if (flowDomainId.length === 0 || fragmentStartKey.length === 0) {
    throw new RangeError('Body occurrence identity requires a flow domain and fragment start');
  }
  return [
    'body-occurrence',
    encodeURIComponent(sourceKey(source)),
    encodeURIComponent(flowDomainId),
    encodeURIComponent(fragmentStartKey),
  ].join('/');
}

export function imageResourceKey(source: SourceRef, partPath: string): string {
  return `image:${sourceKey(source)}:${encodeURIComponent(partPath)}`;
}

export function mathResourceKey(source: SourceRef, localName: string): string {
  return `math:${sourceKey(source)}:${encodeURIComponent(localName)}`;
}

export function anchorOccurrenceKey(source: SourceRef, parserLocalId: string): string {
  return `anchor:${sourceKey(source)}:${encodeURIComponent(parserLocalId)}`;
}

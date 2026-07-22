/**
 * ECMA-376 §17.15.1.18 `w:characterSpacingControl`.
 *
 * `compressPunctuation` removes the half-em side bearing reserved by a
 * full-width East Asian punctuation cell. The reduction is part of the shaped
 * advance before line fitting; retained cluster origins then carry the same
 * reduction to paint. Part 4 §14.8.3.25 explicitly requires applications to
 * apply this compression before deciding whether the final character fits.
 */

const OPENING_FULL_WIDTH_PUNCTUATION = new Set([
  '（', '［', '｛', '〔', '〈', '《', '「', '『', '【', '〘', '〖', '〝', '‘', '“', '｟', '«',
]);

const CLOSING_OR_MIDDLE_FULL_WIDTH_PUNCTUATION = new Set([
  '、', '。', '，', '．', '：', '；', '！', '？',
  '）', '］', '｝', '〕', '〉', '》', '」', '』', '】', '〙', '〗', '〟', '’', '”', '｠', '»', '・',
]);

function isCompressiblePunctuation(value: string): boolean {
  return OPENING_FULL_WIDTH_PUNCTUATION.has(value)
    || CLOSING_OR_MIDDLE_FULL_WIDTH_PUNCTUATION.has(value);
}

export function characterSpacingControlCompressionPt(
  text: string,
  emPt: number,
  control: string | undefined,
): number {
  if (control !== 'compressPunctuation' || !Number.isFinite(emPt) || emPt <= 0) return 0;
  let count = 0;
  for (const scalar of text) if (isCompressiblePunctuation(scalar)) count += 1;
  return count * emPt / 2;
}

export interface CharacterSpacingCluster {
  readonly range: Readonly<{ start: number; end: number }>;
  readonly offsetPt: number;
  readonly advancePt: number;
}

export type CharacterSpacingControlledCluster = CharacterSpacingCluster & Readonly<{
  /** Advance removed from this cluster's retained cell. */
  compressionPt?: number;
  /** Glyph-origin adjustment inside the compressed cell (opening punctuation). */
  paintOffsetPt?: number;
}>;

export function applyCharacterSpacingControlToClusters(
  text: string,
  clusters: readonly CharacterSpacingCluster[],
  emPt: number,
  control: string | undefined,
): readonly CharacterSpacingControlledCluster[] {
  if (control !== 'compressPunctuation') return clusters;
  let precedingCompressionPt = 0;
  return clusters.map((cluster) => {
    const clusterText = text.slice(cluster.range.start, cluster.range.end);
    const requestedCompressionPt = characterSpacingControlCompressionPt(
      clusterText,
      emPt,
      control,
    );
    const compressionPt = Math.min(cluster.advancePt, requestedCompressionPt);
    const openingCount = [...clusterText]
      .filter((scalar) => OPENING_FULL_WIDTH_PUNCTUATION.has(scalar)).length;
    const openingShiftPt = Math.min(compressionPt, openingCount * emPt / 2);
    const controlled: CharacterSpacingControlledCluster = {
      ...cluster,
      offsetPt: cluster.offsetPt - precedingCompressionPt,
      advancePt: cluster.advancePt - compressionPt,
      ...(compressionPt === 0 ? {} : { compressionPt }),
      ...(openingShiftPt === 0 ? {} : { paintOffsetPt: -openingShiftPt }),
    };
    precedingCompressionPt += compressionPt;
    return controlled;
  });
}

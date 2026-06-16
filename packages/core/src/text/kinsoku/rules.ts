/* ------------------------------------------------------------------ *
 * Japanese line-breaking (kinsoku shori / 禁則処理)
 *
 * ECMA-376 §17.15.1.58 `w:kinsoku` is a document-wide on/off toggle for
 * "East Asian typography line-breaking rules". Its default, when the
 * element is absent from settings.xml, is TRUE (the toggle is a
 * ST_OnOff whose absence Word treats as enabled for kinsoku). So a doc
 * with no <w:kinsoku> still gets Japanese line breaking — which is what
 * Word does and what users see.
 *
 * §17.15.1.59 `w:noLineBreaksAfter` / §17.15.1.60 `w:noLineBreaksBefore`
 * let a document override the character set used by the kinsoku engine
 * for a given language (`w:lang`):
 *   - noLineBreaksBefore (§.60): characters that "cannot begin a line"
 *     (行頭禁則 — line-start-forbidden).
 *   - noLineBreaksAfter  (§.59): characters that "cannot end a line"
 *     (行末禁則 — line-end-forbidden).
 * The spec states the `w:val` "specifies the set of characters" — it is
 * the COMPLETE set, so a present override REPLACES the application's
 * default set for that language (it does not extend it). When the
 * element is absent the application's own default set is used. We
 * implement replace-vs-default exactly per that wording.
 *
 * The default sets below are Word's documented Japanese kinsoku tables
 * (Tools ▸ Options ▸ Typography ▸ "Use default kinsoku rules"). They
 * coincide with JIS X 4051 §6.1 (行頭禁則文字 / 行末禁則文字). We encode
 * them as two flat string constants (data, not scattered conditionals);
 * membership is a Set lookup.
 *
 * Word applies kinsoku only to East-Asian wrapping (the per-character
 * break path). Pure-Latin word wrap is untouched: these sets contain no
 * ASCII letters/space, and the Latin path never consults them.
 * ------------------------------------------------------------------ */

/** §17.15.1.60 default 行頭禁則 — characters that may NOT begin a line.
 *  Closing brackets/quotes, mid/end punctuation, small kana, prolonged
 *  sound mark, iteration marks, and their halfwidth forms. */
const KINSOKU_DEFAULT_LINE_START_FORBIDDEN =
  // closing brackets / quotes (fullwidth)
  '”’）〕］｝〉》」』】〙〗〟｠»' +
  // mid / end punctuation (fullwidth)
  '、。，．・：；／？！‐ー゠–〜～' +
  // small kana
  'ぁぃぅぇぉっゃゅょゎゕゖ' +
  'ァィゥェォッャュョヮヵヶ' +
  'ㇰㇱㇲㇳㇴㇵㇶㇷㇸㇹㇺㇻㇼㇽㇾㇿ' +
  // iteration / sound marks
  '々〻ゝゞヽヾ゛゜' +
  // misc trailing symbols
  '％‰℃°′″' +
  // halfwidth forms (cannot start a line either)
  '｡｣､･ｰﾞﾟ' +
  '!),.:;?]}｠';

/** §17.15.1.59 default 行末禁則 — characters that may NOT end a line.
 *  Opening brackets / quotes and currency/lead symbols. */
const KINSOKU_DEFAULT_LINE_END_FORBIDDEN =
  // opening brackets / quotes (fullwidth)
  '“‘（〔［｛〈《「『【〘〖〝｟«' +
  // currency / lead symbols
  '＄￥＃￡￠' +
  // halfwidth opening forms
  '([{｟';

/** Resolved kinsoku configuration for a document.
 *  `enabled` reflects §17.15.1.58; the two sets are §.60 / §.59 (custom
 *  sets replace the defaults — see resolveKinsokuRules). */
export interface KinsokuRules {
  enabled: boolean;
  /** Code points forbidden at line START (行頭禁則). */
  lineStartForbidden: Set<number>;
  /** Code points forbidden at line END (行末禁則). */
  lineEndForbidden: Set<number>;
}

function codePointSet(text: string): Set<number> {
  const out = new Set<number>();
  for (const ch of text) out.add(ch.codePointAt(0)!);
  return out;
}

/** Build the active {@link KinsokuRules} from the document settings.
 *  - `enabled` defaults to TRUE when undefined (§17.15.1.58 default).
 *  - A non-undefined custom set REPLACES the default for that direction
 *    (§17.15.1.59 / §.60 "specifies the set of characters"). An empty
 *    string is a legitimate replacement that disables that direction.
 */
export function resolveKinsokuRules(settings?: {
  kinsoku?: boolean;
  noLineBreaksBefore?: string;
  noLineBreaksAfter?: string;
}): KinsokuRules {
  return {
    enabled: settings?.kinsoku !== false,
    lineStartForbidden: codePointSet(
      settings?.noLineBreaksBefore ?? KINSOKU_DEFAULT_LINE_START_FORBIDDEN,
    ),
    lineEndForbidden: codePointSet(
      settings?.noLineBreaksAfter ?? KINSOKU_DEFAULT_LINE_END_FORBIDDEN,
    ),
  };
}

/** The default Japanese kinsoku rules (no document overrides). */
export const DEFAULT_KINSOKU_RULES: KinsokuRules = resolveKinsokuRules();

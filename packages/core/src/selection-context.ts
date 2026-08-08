/** Resource bounds for browser-native text selection context snapshots. */
export interface TextSelectionContextOptions {
  /** Maximum selected UTF-16 code units returned. Default and hard maximum 65,536. */
  readonly maxTextCharacters?: number;
  /** Maximum intersected rendered-run locators returned. Default and hard maximum 1,024. */
  readonly maxRunLocators?: number;
}

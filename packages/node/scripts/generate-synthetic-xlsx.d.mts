export interface SyntheticXlsxOptions {
  rows?: number;
  columns?: number;
  targetWorksheetBytes?: number;
  /** Test-only: emit a well-formed ZIP whose worksheet XML has a malformed tail. */
  malformedWorksheetTail?: boolean;
}

export interface SyntheticXlsxResult {
  outputPath: string;
  rows: number;
  columns: number;
  cells: number;
  worksheetBytes: number;
  zipBytes: number;
}

export function generateSyntheticXlsx(
  outputPath: string,
  options?: SyntheticXlsxOptions,
): Promise<SyntheticXlsxResult>;

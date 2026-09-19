import { SommelierError } from "@lucamattiazzi/sommelier-core";

/** Parsed zero-based rectangular A1 range. */
export interface ParsedRange {
  readonly address: string;
  readonly startRow: number;
  readonly startColumn: number;
  readonly endRow: number;
  readonly endColumn: number;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly cellCount: number;
}

const cellPattern = /^\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})$/i;

function columnNumber(label: string): number {
  let result = 0;
  for (const character of label.toUpperCase()) result = result * 26 + character.charCodeAt(0) - 64;
  return result - 1;
}

function columnLabel(index: number): string {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

/** Parse and canonicalize a bounded rectangular A1 address. */
export function parseA1Range(input: string): ParsedRange {
  const parts = input.trim().split(":");
  if (parts.length > 2 || parts.length === 0) throw invalidRange(input);
  const start = cellPattern.exec(parts[0] ?? "");
  const end = cellPattern.exec(parts[1] ?? parts[0] ?? "");
  if (!start || !end) throw invalidRange(input);
  const startRow = Number(start[2]) - 1;
  const endRow = Number(end[2]) - 1;
  const startColumn = columnNumber(start[1] ?? "");
  const endColumn = columnNumber(end[1] ?? "");
  if (endRow >= 1_048_576 || endColumn >= 16_384 || endRow < startRow || endColumn < startColumn)
    throw invalidRange(input);
  const rowCount = endRow - startRow + 1;
  const columnCount = endColumn - startColumn + 1;
  return {
    address: `${columnLabel(startColumn)}${startRow + 1}:${columnLabel(endColumn)}${endRow + 1}`,
    startRow,
    startColumn,
    endRow,
    endColumn,
    rowCount,
    columnCount,
    cellCount: rowCount * columnCount,
  };
}

/** Return a canonical A1 cell address from zero-based coordinates. */
export function toA1Cell(row: number, column: number): string {
  return `${columnLabel(column)}${row + 1}`;
}

function invalidRange(input: string): SommelierError {
  return new SommelierError({
    code: "SOMMELIER_EXCEL_RANGE_INVALID",
    message: `Invalid bounded A1 range: ${input}.`,
    context: { range: input },
    suggestedAction:
      "Use an explicit rectangular address such as A1:F40; whole rows and columns are not accepted.",
  });
}

/** Validate that a two-dimensional matrix exactly matches a range. */
export function assertMatrixShape(
  matrix: readonly (readonly unknown[])[],
  range: ParsedRange,
  label: string,
): void {
  const valid =
    matrix.length === range.rowCount && matrix.every((row) => row.length === range.columnCount);
  if (!valid) {
    throw new SommelierError({
      code: "SOMMELIER_EXCEL_SHAPE_MISMATCH",
      message: `${label} must be a ${range.rowCount} by ${range.columnCount} matrix for ${range.address}.`,
      context: {
        range: range.address,
        expectedRows: range.rowCount,
        expectedColumns: range.columnCount,
      },
      suggestedAction: "Send one matrix element for every target cell.",
    });
  }
}

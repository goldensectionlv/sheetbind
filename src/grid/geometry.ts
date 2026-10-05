/** One-based coordinates, independent of file-format limits. */
export interface GridAddress {
  readonly row: number
  readonly column: number
}

/** An ordered, inclusive rectangle. A single-cell range is valid. */
export interface GridRange {
  readonly start: GridAddress
  readonly end: GridAddress
}

export interface GridOffset {
  readonly rows: number
  readonly columns: number
}

/** A1 column labels share the same conversion in formulas and file adapters. */
export function columnNumber(letters: string): number {
  return [...letters.toUpperCase()].reduce((number, char) => number * 26 + char.charCodeAt(0) - 64, 0)
}

export function columnName(column: number): string {
  let name = ''
  for (; column > 0; column = Math.floor((column - 1) / 26)) {
    name = String.fromCharCode(65 + (column - 1) % 26) + name
  }
  return name
}

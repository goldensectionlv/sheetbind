export const WORKBOOK_LIMITS = { rows: 1_048_576, columns: 16_384 } as const

/** One-based XLSX cell coordinates. */
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

/** Formulas, cell values and native metadata share A1 column conversion. */
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
export const XLSX_MAX_ROW = WORKBOOK_LIMITS.rows
export const XLSX_MAX_COLUMN = WORKBOOK_LIMITS.columns

export function assertXlsxAddress(address: GridAddress): void {
  if (!Number.isSafeInteger(address.row) || address.row < 1) {
    throw new RangeError('Address row must be a positive safe integer')
  }
  if (!Number.isSafeInteger(address.column) || address.column < 1) {
    throw new RangeError('Address column must be a positive safe integer')
  }
  if (address.row > XLSX_MAX_ROW || address.column > XLSX_MAX_COLUMN) {
    throw new RangeError('Address exceeds XLSX row or column limits')
  }
}

export function parseAddress(ref: string): GridAddress {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/.exec(ref)
  if (!match) {
    throw new RangeError(`Expected a sheet-local A1 address: ${ref}`)
  }
  const address = {
    row: Number(match[2]),
    column: columnNumber(match[1]),
  }
  assertXlsxAddress(address)
  return address
}

export function parseRange(ref: string): GridRange {
  const parts = ref.split(':')
  if (parts.length > 2) {
    throw new RangeError(`Expected a sheet-local A1 range: ${ref}`)
  }
  const range = { start: parseAddress(parts[0]), end: parseAddress(parts[1] ?? parts[0]) }
  if (range.start.row > range.end.row || range.start.column > range.end.column) {
    throw new RangeError('Range endpoints must be ordered on both axes')
  }
  return range
}

export function formatAddress(address: GridAddress): string {
  assertXlsxAddress(address)
  return `${columnName(address.column)}${address.row}`
}

export function formatRange(range: GridRange): string {
  const start = formatAddress(range.start)
  const end = formatAddress(range.end)
  return start === end ? start : `${start}:${end}`
}

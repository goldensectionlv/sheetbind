import type { GridAddress, GridRange } from '../grid/geometry'
import { columnName, columnNumber } from '../grid/geometry'
import { WORKBOOK_LIMITS } from '../grid/workbook'

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

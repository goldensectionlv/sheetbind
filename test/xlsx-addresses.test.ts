import { expect, it } from 'vitest'
import { assertXlsxAddress, formatAddress, formatRange, parseAddress, parseRange, XLSX_MAX_COLUMN, XLSX_MAX_ROW } from '../src/grid/geometry'

it('reads absolute A1 references and formats addresses through the XLSX boundary', () => {
  for (const [ref, row, column] of [['A1', 1, 1], ['Z9', 9, 26], ['AA10', 10, 27], ['XFD1048576', XLSX_MAX_ROW, XLSX_MAX_COLUMN]] as const) {
    expect(parseAddress(ref)).toEqual({ row, column })
    expect(formatAddress({ row, column })).toBe(ref)
  }
  expect(parseAddress('$AA$10')).toEqual({ row: 10, column: 27 })
})

it('rejects invalid numeric coordinates on either axis', () => {
  for (const value of [0, -1, 1.5, NaN, Infinity, Number.MAX_SAFE_INTEGER]) {
    expect(() => assertXlsxAddress({ row: value, column: 1 })).toThrow(RangeError)
    expect(() => assertXlsxAddress({ row: 1, column: value })).toThrow(RangeError)
  }
  expect(() => formatAddress({ row: XLSX_MAX_ROW + 1, column: 1 })).toThrow(RangeError)
  expect(() => formatAddress({ row: 1, column: XLSX_MAX_COLUMN + 1 })).toThrow(RangeError)
})

it.each(['A0', 'A1048577', 'XFE1', 'A1:B2', 'Sheet!A1'])('rejects an invalid sheet-local address %s', ref => {
  expect(() => parseAddress(ref)).toThrow(RangeError)
})

it('reads ordered ranges and single cells without accepting reversed or extra endpoints', () => {
  expect(parseRange('$A$1:$C$2')).toEqual({ start: { row: 1, column: 1 }, end: { row: 2, column: 3 } })
  expect(parseRange('B2')).toEqual({ start: { row: 2, column: 2 }, end: { row: 2, column: 2 } })
  expect(formatRange({ start: { row: 1, column: 1 }, end: { row: 2, column: 3 } })).toBe('A1:C2')
  expect(formatRange({ start: { row: 2, column: 2 }, end: { row: 2, column: 2 } })).toBe('B2')
  for (const ref of ['A2:A1', 'B1:A1', 'A1:B2:C3']) {
    expect(() => parseRange(ref)).toThrow(RangeError)
  }
})

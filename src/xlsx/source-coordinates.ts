import { axisPositions, mapAxis } from '../grid/workbook-axis'
import type { WorkbookSheetPlan } from '../grid/workbook-layout'
import { FormulaEdge } from '../grid/workbook-formula'
import type { FormulaRows } from '../grid/workbook-formula'
import { columnName, columnNumber } from '../grid/geometry'
import { formatAddress, formatRange, parseAddress, parseRange } from './addresses'
import type { GridAddress, GridRange } from '../grid/geometry'

/** Source XLSX coordinates to final XLSX coordinates, including tag removal and form control rows. */
export interface SourceCoordinates {
  readonly name: string
  readonly part?: string
  readonly row: (position: number, edge?: FormulaEdge) => number | undefined
  readonly rows: (position: number) => readonly number[]
  readonly column: (position: number, edge?: FormulaEdge) => number | undefined
  readonly columns: (position: number) => readonly number[]
  /** First surviving copy; undefined when the source cell was removed. No view fallback. */
  readonly point: (address: string) => string | undefined
  /** Every surviving copy, or an empty list for a removed cell. */
  readonly points: (address: string) => readonly string[]
  /** Inclusive expanded bounds; undefined when the range collapses. */
  readonly range: (value: string) => string | undefined
  /** Surviving addresses and ranges. The XML consumer owns space-separated serialization. */
  readonly references: (value: string) => readonly string[]
}

/** One source-to-output coordinate transform for cells and coordinate-bearing Excel metadata. */
export function sourceCoordinates(plan: WorkbookSheetPlan, carrier?: FormulaRows): SourceCoordinates {
  const { definition: source, sheet: output, axes, coordinates } = plan
  const markers = source.xlsx?.markers ?? []
  const cells = new Map<string, GridAddress[]>(source.xlsx?.cells.map(address => [address, []]) ?? [])
  const addresses = new Map<string, string[]>()
  const ranges = new Map<string, readonly GridRange[]>()
  for (const cell of output.cells) {
    if (cell.xlsx && cell.xlsx.part === source.xlsx?.part) {
      const positions = cells.get(cell.xlsx.address) ?? []
      positions.push(cell.at)
      cells.set(cell.xlsx.address, positions)
    }
  }
  const authored = (row: number, edge: FormulaEdge) => row - markers.filter(marker => edge === FormulaEdge.End ? marker <= row : marker < row).length
  function row(position: number, edge = FormulaEdge.Cell): number | undefined {
    if (edge === FormulaEdge.Cell && markers.includes(position)) {
      return undefined
    }
    const logical = authored(position, edge)
    const placed = mapAxis(axes.rows, logical, edge)
    return placed === undefined ? undefined : carrier ? carrier(placed, edge) : placed
  }
  const column = (position: number, edge = FormulaEdge.Cell) => mapAxis(axes.columns, position, edge)
  function rows(position: number): number[] {
    if (markers.includes(position)) {
      return []
    }
    const logical = authored(position, FormulaEdge.Cell)
    return axisPositions(axes.rows, logical).map(value => carrier ? carrier(value, FormulaEdge.Cell)! : value)
  }
  const columns = (position: number): number[] => axisPositions(axes.columns, position)
  const placedRow = (row: number, edge = FormulaEdge.Cell) => carrier ? carrier(row, edge)! : row
  function points(address: string): string[] {
    if (cells.has(address)) {
      let found = addresses.get(address)
      if (!found) {
        found = cells.get(address)!.map(formatAddress)
        addresses.set(address, found)
      }
      return found
    }
    const at = parseAddress(address)
    if (markers.includes(at.row)) {
      return []
    }
    return coordinates.points({ ...at, row: authored(at.row, FormulaEdge.Cell) })
      .map(at => formatAddress({ ...at, row: placedRow(at.row) }))
  }
  function mappedRanges(value: string): readonly GridRange[] {
    const found = ranges.get(value)
    if (found) {
      return found
    }
    const parsed = parseRange(value)
    const start = { ...parsed.start, row: authored(parsed.start.row, FormulaEdge.Start) }
    const end = { ...parsed.end, row: authored(parsed.end.row, FormulaEdge.End) }
    const mapped = start.row > end.row ? [] : coordinates.ranges({ start, end })
    ranges.set(value, mapped)
    return mapped
  }
  function range(value: string): string | undefined {
    const parts = mappedRanges(value)
    if (!parts.length) {
      return undefined
    }
    const start = { row: Infinity, column: Infinity }
    const end = { row: 0, column: 0 }
    for (const part of parts) {
      start.row = Math.min(start.row, part.start.row)
      start.column = Math.min(start.column, part.start.column)
      end.row = Math.max(end.row, part.end.row)
      end.column = Math.max(end.column, part.end.column)
    }
    return formatRange({ start: { ...start, row: placedRow(start.row, FormulaEdge.Start) }, end: { ...end, row: placedRow(end.row, FormulaEdge.End) } })
  }
  function references(value: string): string[] {
    return value.split(/\s+/).filter(Boolean).flatMap(ref => {
      if (ref.includes(':')) {
        return mappedRanges(ref).flatMap(placedReferences)
      }
      return points(ref.replace(/\$/g, ''))
    })
  }
  function placedReferences(part: GridRange): string[] {
    const result: string[] = []
    let start = part.start.row
    while (start <= part.end.row) {
      const placed = placedRow(start)
      let end = part.end.row
      if (carrier) {
        // Carrier rows only insert gaps. Find each contiguous run without scanning a whole-column range.
        let low = start
        let high = end
        while (low < high) {
          const middle = Math.ceil((low + high) / 2)
          if (placedRow(middle) - placed === middle - start) {
            low = middle
          }
          else {
            high = middle - 1
          }
        }
        end = low
      }
      result.push(formatRange({ start: { ...part.start, row: placed }, end: { ...part.end, row: placed + end - start } }))
      start = end + 1
    }
    return result
  }
  return { name: output.name, part: source.xlsx?.part, row, rows, column, columns, point: address => points(address)[0], points, range, references }
}

/** Excel formula syntax stays opaque; only local A1 references are relocated. */
export function sourceFormula(value: string, current: SourceCoordinates, sheets: ReadonlyMap<string, SourceCoordinates>): string {
  const reference = /(?<![\p{L}\p{N}_.!'])(?:(?:'((?:[^']|'')+)'|([\p{L}_][\p{L}\p{N}_.]*))!)?(\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)(?![\p{L}\p{N}_(])/giu
  return value.split(/("(?:[^"]|"")*"|\[[^\]]*\](?:[^\s!+*/^&<>=(),;]*![A-Z0-9$:]+)?)/gi).map((part, index) => {
    if (index % 2) {
      return part
    }
    return part.replace(reference, (original, quoted: string | undefined, name: string | undefined, range: string) => {
      const map = name || quoted ? sheets.get((quoted?.replace(/''/g, "'") ?? name!).toLowerCase()) : current
      if (!map) {
        return original
      }
      const prefix = name || quoted ? `'${map.name.replace(/'/g, "''")}'!` : ''
      try {
        const [start, end] = range.split(':')
        const cells = end && /^\$?[A-Z]+\$?\d+$/i.test(start) && /^\$?[A-Z]+\$?\d+$/i.test(end)
        const mapped = cells ? map.range(range.replace(/\$/g, '')) : undefined
        if (cells && !mapped) {
          return prefix + '#REF!'
        }
        const bounds = mapped ? parseRange(mapped) : undefined
        function address(value: string, edge: FormulaEdge) {
          const match = /^(\$?[A-Z]+)?(\$?\d+)?$/i.exec(value)!
          const point = edge === FormulaEdge.Cell && match[1] && match[2] ? map!.point(value.replace(/\$/g, '')) : undefined
          if (edge === FormulaEdge.Cell && match[1] && match[2] && !point) {
            return undefined
          }
          const placed = bounds ? edge === FormulaEdge.Start ? bounds.start : bounds.end : point ? parseAddress(point) : undefined
          const column = placed?.column ?? (match[1] ? map!.column(columnNumber(match[1].replace('$', '')), edge) : undefined)
          const row = placed?.row ?? (match[2] ? map!.row(Number(match[2].replace('$', '')), edge) : undefined)
          if (match[1] && column === undefined || match[2] && row === undefined) {
            return undefined
          }
          const text = (column === undefined ? '' : (match[1].startsWith('$') ? '$' : '') + columnName(column))
          + (row === undefined ? '' : (match[2].startsWith('$') ? '$' : '') + row)
          return { column, row, text }
        }
        const first = address(start, end ? FormulaEdge.Start : FormulaEdge.Cell)
        const last = end ? address(end, FormulaEdge.End) : undefined
        // Check both bounds before writing: Excel normalizes reversed ranges.
        if (!first || end && (!last || first.row !== undefined && last.row !== undefined && first.row > last.row
          || first.column !== undefined && last.column !== undefined && first.column > last.column)) {
          return prefix + '#REF!'
        }
        return prefix + first.text + (last ? ':' + last.text : '')
      }
      catch {
        return original
      }
    })
  }).join('')
}

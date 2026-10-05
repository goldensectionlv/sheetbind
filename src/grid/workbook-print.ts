import type { GridRange } from './geometry'

export interface WorkbookPrint {
  readonly area?: GridRange
  readonly repeatRows?: { readonly start: number, readonly end: number }
}

/** Print ranges follow the same placement as their worksheet contents. */
export function mapWorkbookPrint(print: WorkbookPrint | undefined, coordinates: {
  rowStart: (row: number) => number
  rowEnd?: (row: number) => number
  columnStart?: (column: number) => number
  columnEnd?: (column: number) => number
}): WorkbookPrint | undefined {
  if (!print) {
    return undefined
  }
  const { rowStart, rowEnd = rowStart, columnStart = column => column, columnEnd = columnStart } = coordinates
  const area = print.area ? { start: { row: rowStart(print.area.start.row), column: columnStart(print.area.start.column) }, end: { row: rowEnd(print.area.end.row), column: columnEnd(print.area.end.column) } } : undefined
  const repeatRows = print.repeatRows ? { start: rowStart(print.repeatRows.start), end: rowEnd(print.repeatRows.end) } : undefined
  return {
    ...(area && area.end.row >= area.start.row && area.end.column >= area.start.column ? { area } : {}),
    ...(repeatRows && repeatRows.end >= repeatRows.start ? { repeatRows } : {}),
  }
}

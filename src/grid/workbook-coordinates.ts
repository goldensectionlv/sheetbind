import type { GridAddress, GridRange } from './geometry'
import { mapAxis } from './workbook-axis'
import type { WorkbookAxes } from './workbook-axis'
import { FormulaEdge } from './workbook-formula'
import { WORKBOOK_LIMITS } from './workbook'
import type { WorkbookRegionView } from './workbook'

export type WorkbookIndexes = ReadonlyMap<string, number>
const root: WorkbookIndexes = new Map()

function contains(region: WorkbookRegionView, at: GridAddress): boolean {
  return at.row >= region.row && at.row < region.row + region.height
    && at.column >= (region.column ?? 1) && at.column < (region.column ?? 1) + (region.width ?? WORKBOOK_LIMITS.columns)
}

/** Split at band boundaries, without enumerating the cells in a native range. */
export function partitionWorkbookRange(regions: readonly WorkbookRegionView[], range: GridRange) {
  const rows = new Set([range.start.row, range.end.row + 1])
  const columns = new Set([range.start.column, range.end.column + 1])
  for (const region of regions) {
    for (const row of [region.row, region.row + region.height]) {
      if (row > range.start.row && row <= range.end.row) {
        rows.add(row)
      }
    }
    for (const column of [region.column ?? 1, (region.column ?? 1) + (region.width ?? WORKBOOK_LIMITS.columns)]) {
      if (column > range.start.column && column <= range.end.column) {
        columns.add(column)
      }
    }
  }
  const rowCuts = [...rows].sort((a, b) => a - b)
  const columnCuts = [...columns].sort((a, b) => a - b)
  return rowCuts.slice(0, -1).flatMap((row, r) => columnCuts.slice(0, -1).map((column, c) => {
    const start = { row, column }
    return { range: { start, end: { row: rowCuts[r + 1] - 1, column: columnCuts[c + 1] - 1 } },
      owner: regions.findLast(region => contains(region, start)) }
  }))
}

/** Cells and native objects use the same axis plans and owning iteration indexes. */
export function placeWorkbookPoint(axes: WorkbookAxes, at: GridAddress, indexes: WorkbookIndexes): GridAddress | undefined {
  const row = mapAxis(axes.rows, at.row, FormulaEdge.Cell, indexes)
  const column = mapAxis(axes.columns, at.column, FormulaEdge.Cell, indexes)
  return row === undefined || column === undefined ? undefined : { row, column }
}

export function placeWorkbookRegion(axes: WorkbookAxes, region: WorkbookRegionView, indexes: WorkbookIndexes) {
  const row = mapAxis(axes.rows, region.row, FormulaEdge.Start, indexes)!
  const column = mapAxis(axes.columns, region.column ?? 1, FormulaEdge.Start, indexes)!
  return { row, column, height: mapAxis(axes.rows, region.row + region.height - 1, FormulaEdge.End, indexes)! - row + 1,
    width: mapAxis(axes.columns, (region.column ?? 1) + (region.width ?? WORKBOOK_LIMITS.columns) - 1, FormulaEdge.End, indexes)! - column + 1 }
}

export type WorkbookCoordinates = ReturnType<typeof workbookCoordinates>
function mergeRanges(ranges: GridRange[]): GridRange[] {
  for (const axis of ['column', 'row'] as const) {
    const other = axis === 'row' ? 'column' : 'row'
    ranges.sort((a, b) => a.start[other] - b.start[other] || a.end[other] - b.end[other] || a.start[axis] - b.start[axis])
    const merged: GridRange[] = []
    for (const range of ranges) {
      const previous = merged.at(-1)
      if (previous && previous.start[other] === range.start[other] && previous.end[other] === range.end[other]
        && range.start[axis] <= previous.end[axis] + 1) {
        merged[merged.length - 1] = { start: previous.start, end: { ...previous.end, [axis]: Math.max(previous.end[axis], range.end[axis]) } }
      }
      else {
        merged.push(range)
      }
    }
    ranges = merged
  }
  return ranges
}

export function workbookCoordinates(regions: readonly WorkbookRegionView[], axes: WorkbookAxes, instances: ReadonlyMap<string, readonly WorkbookIndexes[]>) {
  const contexts = (owner?: WorkbookRegionView) => owner ? instances.get(owner.id) ?? [] : [root]
  return {
    points(at: GridAddress): GridAddress[] {
      return contexts(regions.findLast(region => contains(region, at))).flatMap(indexes => {
        const placed = placeWorkbookPoint(axes, at, indexes)
        return placed ? [placed] : []
      })
    },
    ranges(range: GridRange): GridRange[] {
      return mergeRanges(partitionWorkbookRange(regions, range).flatMap(part => contexts(part.owner).flatMap(indexes => {
        const start = placeWorkbookPoint(axes, part.range.start, indexes)
        const end = placeWorkbookPoint(axes, part.range.end, indexes)
        return start && end ? [{ start, end }] : []
      })))
    },
  }
}

import { WORKBOOK_LIMITS } from './geometry'
import { TemplateError } from '../core/template'
import type { DataReference, ValueExpression } from '../core/template'
import type { GridAddress, GridOffset, GridRange } from './geometry'
import type { FieldRules } from '../core/field-rules'
import type { WorkbookFormula } from './workbook-formula'

type WorkbookValue = ValueExpression | WorkbookFormula
export interface WorkbookCell {
  readonly id: string
  readonly at: GridAddress
  readonly size: GridOffset
  readonly value: WorkbookValue
  readonly rules?: FieldRules
  /** Original XLSX cell; the adapter preserves content outside the supported model. */
  readonly xlsx?: { readonly part: string, readonly address: string, readonly value: WorkbookValue }
}
export interface WorkbookRow { readonly index: number, readonly height?: number, readonly hidden?: boolean }
export interface WorkbookColumn { readonly index: number, readonly width?: number, readonly hidden?: boolean }
export enum WorkbookAxis { Rows = 'rows', Columns = 'columns' }
export interface WorkbookBody {
  readonly cells: readonly WorkbookCell[]
  /** Native content with no cell value still occupies its part of the grid. */
  readonly occupied?: readonly GridRange[]
  readonly rows?: readonly WorkbookRow[]
  readonly regions?: readonly WorkbookRegion[]
}
/** Coordinates are local to the body. An omitted width retains the whole-row contract. */
export interface WorkbookRegion extends WorkbookBody {
  readonly id: string
  /** Authored opening marker, retained for adapter diagnostics. */
  readonly xlsx?: { readonly address: string }
  readonly type: 'scope' | 'repeat'
  readonly source: DataReference
  readonly row: number
  readonly height: number
  readonly column?: number
  readonly width?: number
  readonly axis?: `${WorkbookAxis}`
}
export interface WorkbookSheet extends WorkbookBody {
  readonly id: string
  readonly name: string
  readonly state?: 'visible' | 'hidden' | 'veryHidden'
  readonly columns?: readonly WorkbookColumn[]
  readonly print?: WorkbookPrint
  readonly xlsx?: { readonly part: string, readonly markers: readonly number[], readonly cells: readonly string[] }
}
/** Internal bindings and placement compiled from a workbook. */
export interface WorkbookDefinition { readonly sheets: readonly WorkbookSheet[] }
export { WORKBOOK_LIMITS } from './geometry'

export function workbookIssue(code: string, nodeId: string, message: string, path = nodeId, phase: 'template' | 'data' = 'template'): never {
  throw new TemplateError([{ phase, code, nodeId, path, message }])
}
/** Flattened regions with sheet coordinates and nesting information. */
export type WorkbookRegionView = WorkbookRegion & { readonly parentId?: string, readonly depth: number }
export function workbookRegions(sheet: WorkbookBody): WorkbookRegionView[] {
  const visit = (body: WorkbookBody, offset: number, columnOffset: number, width: number, depth: number, parentId?: string): WorkbookRegionView[] => (body.regions ?? []).flatMap(region => {
    const row = offset + region.row
    const column = columnOffset + (region.column ?? 1)
    const extent = region.width ?? width
    return [{ ...region, row, ...(column !== 1 || region.column !== undefined ? { column } : {}), ...(extent !== WORKBOOK_LIMITS.columns ? { width: extent } : {}), parentId, depth }, ...visit(region, row - 1, column - 1, extent, depth + 1, region.id)]
  })
  return visit(sheet, 0, 0, WORKBOOK_LIMITS.columns, 0)
}
/** A cell projected to sheet coordinates with its owning region. */
type WorkbookCellView = WorkbookCell & { readonly regionId?: string }
export function workbookCells(sheet: WorkbookSheet): WorkbookCellView[] {
  return [...sheet.cells, ...workbookRegions(sheet).flatMap(region => region.cells.map(cell => ({ ...cell, regionId: region.id, at: { row: region.row + cell.at.row - 1, column: (region.column ?? 1) + cell.at.column - 1 } })))]
}
export function workbookRows(sheet: WorkbookBody): WorkbookRow[] {
  return [...sheet.rows ?? [], ...workbookRegions(sheet).flatMap(region => (region.rows ?? []).map(row => ({ ...row, index: region.row + row.index - 1 })))]
}
/** Discover declared dictionaries without resolving data or layout. */
export function collectWorkbookDictionarySources(template: WorkbookDefinition): string[] {
  return [...new Set(template.sheets.flatMap(sheet => workbookCells(sheet).flatMap(cell => cell.rules?.list
    ? [cell.rules.list]
    : cell.rules?.choice && 'dictionary' in cell.rules.choice.source ? [cell.rules.choice.source.dictionary] : [])))].sort()
}
function intersects(a: Pick<WorkbookCell, 'at' | 'size'>, b: Pick<WorkbookCell, 'at' | 'size'>): boolean {
  return a.at.row < b.at.row + b.size.rows && b.at.row < a.at.row + a.size.rows
    && a.at.column < b.at.column + b.size.columns && b.at.column < a.at.column + a.size.columns
}
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

/** Geometric band compatibility is a template invariant, including empty repeats. */
function assertWorkbookBands(sheet: WorkbookBody): void {
  for (const axis of [WorkbookAxis.Rows, WorkbookAxis.Columns]) {
    const bands = new Map<string, { start: number, length: number }[]>()
    function visit(body: WorkbookBody, offset: number, context: string): void {
      for (const region of body.regions ?? []) {
        const start = offset + (axis === WorkbookAxis.Rows ? region.row : region.column ?? 1)
        const length = axis === WorkbookAxis.Rows ? region.height : region.width ?? WORKBOOK_LIMITS.columns
        if (region.type !== 'repeat' || (region.axis ?? WorkbookAxis.Rows) !== axis) {
          visit(region, start - 1, context)
          continue
        }
        const peers = bands.get(context) ?? []
        if (peers.some(peer => start < peer.start + peer.length && peer.start < start + length && (start !== peer.start || length !== peer.length))) {
          workbookIssue('growth-band-overlap', region.id, `Overlapping ${axis} repeats must reserve the same band`)
        }
        peers.push({ start, length })
        bands.set(context, peers)
        visit(region, 0, `${context}/${start}:${length}`)
      }
    }
    visit(sheet, 0, '')
  }
}

/** Validate ownership and repeat geometry after the XLSX adapter has parsed the tags. */
export function validateWorkbookDefinition(template: WorkbookDefinition): WorkbookDefinition {
  for (const sheet of template.sheets) {
    validateBody(sheet, WORKBOOK_LIMITS.rows, WORKBOOK_LIMITS.columns)
    assertWorkbookBands(sheet)
  }
  return template
}

function validateBody(body: WorkbookBody, height: number, width: number, owner?: WorkbookRegion): void {
  const regions = body.regions ?? []
  const bounds = (region: WorkbookRegion) => ({ at: { row: region.row, column: region.column ?? 1 }, size: { rows: region.height, columns: region.width ?? width } })
  for (const [index, region] of regions.entries()) {
    if (region.height < 1) {
      workbookIssue('empty-region', region.id, 'A region needs at least one body row')
    }
    const column = region.column ?? 1
    const extent = region.width ?? width
    if (region.row + region.height - 1 > height || column + extent - 1 > width) {
      workbookIssue('region-boundary', region.id, 'Region exceeds its parent body')
    }
    if (regions.slice(0, index).some(other => intersects(bounds(region), bounds(other)))) {
      workbookIssue('region-overlap', region.id, 'Sibling regions cannot overlap; use a child region for nesting')
    }
    validateBody(region, region.height, extent, region)
  }
  const occupied = new Set<number>()
  for (const cell of body.cells) {
    if (cell.at.row + cell.size.rows - 1 > height || cell.at.column + cell.size.columns - 1 > width) {
      workbookIssue('region-boundary', cell.id, owner ? 'Cell or merge crosses its region body' : 'Cell exceeds the XLSX grid')
    }
    if (regions.some(region => intersects(cell, bounds(region)))) {
      workbookIssue('region-boundary', cell.id, 'Cell or merge intersects a region without belonging to its body')
    }
    for (let row = cell.at.row; row < cell.at.row + cell.size.rows; row++) {
      for (let column = cell.at.column; column < cell.at.column + cell.size.columns; column++) {
        const key = row * WORKBOOK_LIMITS.columns + column
        if (occupied.has(key)) {
          workbookIssue('cell-overlap', cell.id, 'Cells or merges overlap')
        }
        occupied.add(key)
      }
    }
  }
}

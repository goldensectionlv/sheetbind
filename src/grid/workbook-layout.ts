import { resolveWorkbookData } from './workbook-data'
import type { WorkbookData, WorkbookDataOptions } from './workbook-data'
import type { Origin, TemplateValue } from '../core/template'
import type { FieldRules } from '../core/field-rules'
import type { GridAddress, GridOffset } from './geometry'
import type { WorkbookChoice } from '../core/choices'
import { workbookCells, workbookRegions, workbookRows, workbookIssue, WORKBOOK_LIMITS } from './workbook'
import type { WorkbookCell, WorkbookSheet, WorkbookDefinition, WorkbookRow, WorkbookColumn } from './workbook'
import { mapWorkbookPrint } from './workbook'
import type { WorkbookPrint } from './workbook'
import type { WorkbookFormula } from './workbook-formula'
import { resolveWorkbookFormulas } from './workbook-formulas'
import { axisPositions, mapAxis, planWorkbookAxes } from './workbook-axis'
import type { AxisPlan, WorkbookAxes } from './workbook-axis'
import { FormulaEdge } from './workbook-formula'
import { placeWorkbookPoint, placeWorkbookRegion, workbookCoordinates } from './workbook-coordinates'
import type { WorkbookCoordinates, WorkbookIndexes } from './workbook-coordinates'

/** Public resolved cell; source-package details stay in the internal placement. */
export interface WorkbookCellInstance {
  readonly id: string
  readonly definitionId: string
  readonly at: GridAddress
  readonly size: GridOffset
  readonly rules?: FieldRules
  readonly origin: Origin
  readonly contextPath: string
  readonly choice?: WorkbookChoice
  readonly value: { readonly literal: TemplateValue } | WorkbookFormula
}
export interface WorkbookLayout {
  readonly sheets: readonly {
    readonly id: string
    readonly name: string
    readonly state?: 'visible' | 'hidden' | 'veryHidden'
    readonly rows?: readonly WorkbookRow[]
    readonly columns?: readonly WorkbookColumn[]
    readonly print?: WorkbookPrint
    readonly cells: readonly WorkbookCellInstance[]
  }[]
}

/** Internal placement retains the source references needed by the XLSX writer. */
export interface WorkbookPlacedCell extends WorkbookCellInstance { readonly xlsx?: WorkbookCell['xlsx'] }
export interface WorkbookPlacedSheet extends Pick<WorkbookSheet, 'id' | 'name' | 'state' | 'rows' | 'columns' | 'print'> { readonly cells: readonly WorkbookPlacedCell[] }

/** A sheet and its geometry travel together through every placement stage. */
export interface WorkbookSheetPlan {
  readonly definition: WorkbookSheet
  readonly data: WorkbookData
  readonly sheet: WorkbookPlacedSheet
  readonly axes: WorkbookAxes
  readonly coordinates: WorkbookCoordinates
  readonly extent: { readonly rows: number, readonly columns: number }
  readonly authored: ReadonlyMap<string, WorkbookCell>
}
export interface WorkbookPlan { readonly sheets: readonly WorkbookSheetPlan[] }

/** Accept a normalized definition. Placement never reparses or modifies the input. */
export function planWorkbook(config: WorkbookDefinition, data: unknown, options: WorkbookDataOptions = {}): WorkbookPlan {
  return resolveWorkbookFormulas(placeWorkbook(config, data, options))
}

/** Keep authored formulas until all placement transforms are known. */
export function placeWorkbook(config: WorkbookDefinition, data: unknown, options: WorkbookDataOptions = {}): WorkbookPlan {
  const execution = resolveWorkbookData(config, data, options)
  return { sheets: config.sheets.map((sheet, index) => placeWorkbookSheet(sheet, execution[index])) }
}

export function placeWorkbookSheet(sheet: WorkbookSheet, group: WorkbookData): WorkbookSheetPlan {
  const cells: WorkbookPlacedCell[] = []
  const axes = planWorkbookAxes(group)
  const extent = { rows: 0, columns: 0 }
  const authored = new Map(workbookCells(sheet).map(cell => [cell.id, cell]))
  const views = new Map(workbookRegions(sheet).map(region => [region.id, region]))
  const instances = new Map<string, WorkbookIndexes[]>()
  function placeCells(body: WorkbookData, indexes: WorkbookIndexes): WorkbookPlacedCell[] {
    const suffix = JSON.stringify(body.iterations) + ']'
    return body.cells.map(value => {
      const cell = value.definition
      const origin = { nodeId: cell.id, dataPath: value.dataPath, iterations: body.iterations }
      const position = authored.get(cell.id)!.at
      const at = placeWorkbookPoint(axes, position, indexes)
      if (!at) {
        workbookIssue('growth-crosses-cell', cell.id, 'A fixed cell occupies a removed band; put it inside the repeat or outside its band', value.dataPath, 'data')
      }
      const end = { row: cell.size.rows === 1 ? at.row : mapAxis(axes.rows, position.row + cell.size.rows - 1, FormulaEdge.End, indexes)!,
        column: cell.size.columns === 1 ? at.column : mapAxis(axes.columns, position.column + cell.size.columns - 1, FormulaEdge.End, indexes)! }
      if (at.row === undefined || at.column === undefined || end.row === undefined || end.column === undefined) {
        workbookIssue('growth-crosses-cell', cell.id, 'A fixed cell occupies a removed band; put it inside the repeat or outside its band', value.dataPath, 'data')
      }
      const rows = end.row - at.row + 1
      const columns = end.column - at.column + 1
      const size = rows === cell.size.rows && columns === cell.size.columns ? cell.size : { rows, columns }
      if (end.row > WORKBOOK_LIMITS.rows) {
        workbookIssue('row-limit', cell.id, 'Rendered content exceeds the XLSX row limit', value.dataPath, 'data')
      }
      if (end.column > WORKBOOK_LIMITS.columns) {
        workbookIssue('column-limit', cell.id, 'Rendered content exceeds the XLSX column limit', value.dataPath, 'data')
      }
      extent.rows = Math.max(extent.rows, end.row)
      extent.columns = Math.max(extent.columns, end.column)
      return { id: '[' + JSON.stringify(cell.id) + ',' + suffix, definitionId: cell.id, at, size,
        value: 'path' in cell.value ? { literal: value.value } : cell.value, rules: cell.rules, xlsx: cell.xlsx, origin, contextPath: body.path, choice: value.choice }
    })
  }
  function layout(body: WorkbookData, id?: string): void {
    const definition = body.definition
    const indexes = new Map(body.iterations.map(item => [item.nodeId, item.index]))
    if (id) {
      const copies = instances.get(id) ?? []
      copies.push(indexes)
      instances.set(id, copies)
    }
    const owner = id ? views.get(id) : undefined
    for (const range of definition.occupied ?? []) {
      const end = placeWorkbookPoint(axes, {
        row: range.end.row + (owner?.row ?? 1) - 1,
        column: range.end.column + (owner?.column ?? 1) - 1,
      }, indexes)!
      extent.rows = Math.max(extent.rows, end.row)
      extent.columns = Math.max(extent.columns, end.column)
    }
    for (const node of [...body.regions].sort((a, b) => a.definition.row - b.definition.row)) {
      const region = node.definition
      const placed = placeWorkbookRegion(axes, views.get(region.id)!, indexes)
      if (placed.row + placed.height - 1 > WORKBOOK_LIMITS.rows) {
        workbookIssue('row-limit', region.id, 'Repeated rows exceed the XLSX row limit', node.path, 'data')
      }
      if (region.width !== undefined && placed.column + placed.width - 1 > WORKBOOK_LIMITS.columns) {
        workbookIssue('column-limit', region.id, 'Repeated columns exceed the XLSX column limit', node.path, 'data')
      }
      node.instances.forEach(instance => layout(instance, region.id))
    }
    cells.push(...placeCells(body, indexes))
  }
  layout(group)
  const print = mapWorkbookPrint(sheet.print, {
    rowStart: row => mapAxis(axes.rows, row, FormulaEdge.Start)!, rowEnd: row => mapAxis(axes.rows, row, FormulaEdge.End)!,
    columnStart: column => mapAxis(axes.columns, column, FormulaEdge.Start)!, columnEnd: column => mapAxis(axes.columns, column, FormulaEdge.End)!,
  })
  if (print?.area && (print.area.end.row > WORKBOOK_LIMITS.rows || print.area.end.column > WORKBOOK_LIMITS.columns) || print?.repeatRows && print.repeatRows.end > WORKBOOK_LIMITS.rows) {
    workbookIssue('print-limit', sheet.id, 'Rendered print range exceeds the XLSX grid', `$template.${sheet.id}.print`, 'data')
  }
  const rows = placeSettings(workbookRows(sheet), axes.rows, WORKBOOK_LIMITS.rows)
  const columns = placeSettings(sheet.columns ?? [], axes.columns, WORKBOOK_LIMITS.columns)
  const occupied = new Set<number>()
  for (const cell of cells) {
    for (let row = cell.at.row; row < cell.at.row + cell.size.rows; row++) {
      for (let column = cell.at.column; column < cell.at.column + cell.size.columns; column++) {
        const key = row * WORKBOOK_LIMITS.columns + column
        if (occupied.has(key)) {
          workbookIssue('placed-overlap', cell.definitionId, 'Rendered regions overlap', cell.origin.dataPath, 'data')
        }
        occupied.add(key)
      }
    }
  }
  const placed = { id: sheet.id, name: sheet.name, state: sheet.state, ...(print ? { print } : {}), rows, columns, cells: cells.sort((a, b) => a.at.row - b.at.row || a.at.column - b.at.column) }
  return { definition: sheet, data: group, sheet: placed, axes, coordinates: workbookCoordinates(sheet, axes, instances), authored, extent }
}

function placeSettings<T extends { readonly index: number }>(settings: readonly T[], axis: AxisPlan, limit: number): T[] {
  const result = new Map<number, T>()
  for (const setting of settings) {
    for (const index of axisPositions(axis, setting.index)) {
      if (index > limit) {
        throw new RangeError('Rendered row/column settings exceed the XLSX grid')
      }
      const placed = { ...setting, index }
      const previous = result.get(index)
      if (previous && Object.entries(placed).some(([key, value]) => Reflect.get(previous, key) !== value)) {
        throw new RangeError('Conflicting settings on a shared row or column')
      }
      result.set(index, placed)
    }
  }
  return [...result.values()].sort((a, b) => a.index - b.index)
}

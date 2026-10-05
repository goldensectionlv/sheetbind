import { TemplateError } from '../core/template'
import type { DataReference, ValueExpression } from '../core/template'
import type { GridAddress, GridOffset, GridRange } from './geometry'
import type { FieldRules } from '../core/field-rules'
import type { WorkbookPrint } from './workbook-print'
import type { WorkbookFormula } from './workbook-formula'

export type WorkbookValue = ValueExpression | WorkbookFormula
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
// Coordinates are bounded only by the XLSX format represented by this workbook.
export const WORKBOOK_LIMITS = { rows: 1_048_576, columns: 16_384 } as const

export function workbookIssue(code: string, nodeId: string, message: string, path = nodeId, phase: 'template' | 'data' = 'template'): never {
  throw new TemplateError([{ phase, code, nodeId, path, message }])
}
export function regionBounds(region: WorkbookRegion) {
  return { at: { row: region.row, column: region.column ?? 1 }, size: { rows: region.height, columns: region.width ?? WORKBOOK_LIMITS.columns } }
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
export type WorkbookCellView = WorkbookCell & { readonly regionId?: string }
export function workbookCells(sheet: WorkbookSheet): WorkbookCellView[] {
  return [...sheet.cells, ...workbookRegions(sheet).flatMap(region => region.cells.map(cell => ({ ...cell, regionId: region.id, at: { row: region.row + cell.at.row - 1, column: (region.column ?? 1) + cell.at.column - 1 } })))]
}
export function workbookRows(sheet: WorkbookBody): WorkbookRow[] {
  return [...sheet.rows ?? [], ...workbookRegions(sheet).flatMap(region => (region.rows ?? []).map(row => ({ ...row, index: region.row + row.index - 1 })))]
}
/** Transform local cells while retaining every unchanged branch of the definition. */
export function mapWorkbookCells<T extends WorkbookBody>(body: T, transform: (cell: WorkbookCell) => WorkbookCell): T {
  const cells = body.cells.map(transform)
  const regions = body.regions?.map(region => mapWorkbookCells(region, transform))
  if (cells.every((cell, index) => cell === body.cells[index]) && (!regions || regions.every((region, index) => region === body.regions![index]))) {
    return body
  }
  return { ...body, cells, ...(regions ? { regions } : {}) }
}

/** Discover declared dictionaries without resolving data or layout. */
export function collectWorkbookDictionarySources(template: WorkbookDefinition): string[] {
  return [...new Set(template.sheets.flatMap(sheet => workbookCells(sheet).flatMap(cell => cell.rules?.list
    ? [cell.rules.list]
    : cell.rules?.choice && 'dictionary' in cell.rules.choice.source ? [cell.rules.choice.source.dictionary] : [])))].sort()
}
export function intersects(a: Pick<WorkbookCell, 'at' | 'size'>, b: Pick<WorkbookCell, 'at' | 'size'>): boolean {
  return a.at.row < b.at.row + b.size.rows && b.at.row < a.at.row + a.size.rows
    && a.at.column < b.at.column + b.size.columns && b.at.column < a.at.column + a.size.columns
}
/** Remove object scopes using explicit paths; repeats establish a new current item. */
export function expandWorkbookScopes(template: WorkbookDefinition): WorkbookDefinition {
  const expand = <T extends WorkbookBody>(body: T, prefix?: WorkbookRegion['source'], local = false): T => {
    const qualify = (reference: WorkbookRegion['source']) => reference.from === 'root' || !prefix
      ? reference
      : { ...reference, path: `${prefix.path}.${reference.path}`, from: prefix.from }
    const cells = body.cells.map(cell => ({ ...cell, value: 'path' in cell.value ? qualify(cell.value) : cell.value,
      ...(cell.rules?.choice && 'path' in cell.rules.choice.source ? { rules: { ...cell.rules, choice: { ...cell.rules.choice, source: qualify(cell.rules.choice.source) } } } : {}),
    }))
    const rows = [...body.rows ?? []]
    const occupied = [...body.occupied ?? []]
    const regions: WorkbookRegion[] = []
    for (const region of body.regions ?? []) {
      const source = qualify(region.source)
      if (region.type === 'repeat') {
        regions.push({ ...expand(region, undefined, true), source })
        continue
      }
      const expanded = expand(region, { ...source, from: source.from === 'root' || !local ? 'root' : 'current' }, local)
      cells.push(...expanded.cells.map(cell => ({ ...cell, at: { row: cell.at.row + region.row - 1, column: cell.at.column + (region.column ?? 1) - 1 } })))
      rows.push(...(expanded.rows ?? []).map(row => ({ ...row, index: row.index + region.row - 1 })))
      occupied.push(...(expanded.occupied ?? []).map(range => ({
        start: { row: range.start.row + region.row - 1, column: range.start.column + (region.column ?? 1) - 1 },
        end: { row: range.end.row + region.row - 1, column: range.end.column + (region.column ?? 1) - 1 },
      })))
      regions.push(...(expanded.regions ?? []).map(child => ({ ...child, row: child.row + region.row - 1,
        ...(region.column !== undefined || child.column !== undefined ? { column: (child.column ?? 1) + (region.column ?? 1) - 1 } : {}),
        ...(region.width !== undefined && child.width === undefined ? { width: region.width } : {}) })))
    }
    return { ...body, cells, ...(occupied.length ? { occupied } : {}), ...(body.rows || rows.length ? { rows: rows.sort((a, b) => a.index - b.index) } : {}), ...(body.regions ? { regions } : {}) }
  }
  return { ...template, sheets: template.sheets.map(sheet => expand(sheet)) }
}

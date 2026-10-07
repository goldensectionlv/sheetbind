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

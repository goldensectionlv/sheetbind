import { createHash } from 'node:crypto'
import type { Worksheet } from 'exceljs'
import { createWorkbookFormData, referencePath, WorkbookFormInputError } from '../form/workbook'
import type { DataPath } from '../form/records'
import { isWorkbookFormRow, workbookFormRowFields } from '../form/workbook-rows'
import type { WorkbookFormRows } from '../form/workbook-rows'
import type { WorkbookBody, WorkbookRegion, WorkbookDefinition } from '../grid/workbook'
import { workbookCells, WORKBOOK_LIMITS } from '../grid/workbook'
import { mapWorkbookPrint } from '../grid/workbook-print'
import type { WorkbookPlan, WorkbookRegionLayout, WorkbookPlacement } from '../grid/workbook-layout'
import { formatAddress } from './addresses'
import { xlsxTextIssues } from './report-text'
import { parseValidation } from '../core/validation'
import type { FieldRules } from '../core/field-rules'
import { isDataObject } from '../core/json'
import { FormulaEdge } from '../grid/workbook-formula'
import type { FormulaRows } from '../grid/workbook-formula'

export enum FormMarkerKind {
  Sheet = 'sheet', SheetEnd = '/sheet', Repeat = 'repeat', RepeatEnd = '/repeat', Item = 'item', ItemEnd = '/item',
}
export const FORM_MARKER_COLUMN = 257
export const FORM_MARKER_PREFIX = 'sheetbind.form/3:'
type FormMarkerToken =
  | readonly [FormMarkerKind.Sheet, string]
  | readonly [FormMarkerKind.SheetEnd]
  | readonly [FormMarkerKind.Repeat | FormMarkerKind.RepeatEnd | FormMarkerKind.Item | FormMarkerKind.ItemEnd, number]
export interface FormMarker { readonly row: number, readonly token: FormMarkerToken }
interface FormMarkers { readonly column: number, readonly rows: readonly FormMarker[] }
interface FormCarrierDefinition {
  readonly numbers: ReadonlyMap<string, number>
  readonly regions: ReadonlyMap<string, WorkbookRegion>
  readonly signature: string
  readonly sheets: ReadonlySet<string>
}

function signatureRules(rules: FieldRules = {}) {
  const choice = rules.choice
  return {
    validation: parseValidation(rules.validation ?? []).map(use => ({ rule: use.rule, args: use.args ?? [] })),
    ...(rules.list ? { list: rules.list } : {}),
    ...(choice ? { choice: { ...choice, return: choice.return ?? 'object', source: 'path' in choice.source ? { ...choice.source, from: choice.source.from ?? 'current' } : choice.source } } : {}),
  }
}

export function formCarrierDefinition(template: WorkbookDefinition): FormCarrierDefinition {
  const sheets = new Set(template.sheets.filter(sheet => sheet.regions?.length || workbookCells(sheet).some(cell => 'path' in cell.value)).map(sheet => sheet.name))
  const numbers = new Map<string, number>()
  const regions = new Map<string, WorkbookRegion>()
  let serial = 0
  const reference = (value: { path: string, from?: string }) => ({ path: value.path, from: value.from ?? 'current' })
  const body = (source: WorkbookBody): unknown => ({
    cells: source.cells.filter(cell => 'path' in cell.value).sort((a, b) => a.at.row - b.at.row || a.at.column - b.at.column).map(cell => ({
      at: cell.at, size: cell.size,
      value: 'path' in cell.value ? { ...reference(cell.value), optional: !!cell.value.optional } : cell.value, rules: signatureRules(cell.rules),
    })),
    regions: [...source.regions ?? []].sort((a, b) => a.row - b.row).map(region => {
      numbers.set(region.id, ++serial)
      regions.set(region.id, region)
      return { row: region.row, height: region.height, column: region.column ?? 1, width: region.width ?? WORKBOOK_LIMITS.columns,
        source: reference(region.source), body: body(region) }
    }),
  })
  const semantic = template.sheets.filter(sheet => sheets.has(sheet.name)).sort((a, b) => a.name < b.name ? -1 : a.name > b.name ? 1 : 0).map(sheet => ({ name: sheet.name, body: body(sheet) }))
  const canonical = (value: unknown): unknown => Array.isArray(value)
    ? value.map(canonical)
    : isDataObject(value) ? Object.fromEntries(Object.keys(value).sort().map(key => [key, canonical(value[key])])) : value
  const signature = createHash('sha256').update(JSON.stringify(canonical(semantic))).digest('hex')
  return { numbers, signature, regions, sheets }
}

/** Hidden rows delimit the current structure without retaining issued counts or record identities. */
export function placeFormMarkers(plan: WorkbookPlan, definition: FormCarrierDefinition) {
  const markers = new Map<string, FormMarkers>()
  const formulaRows = new Map<string, FormulaRows>()
  const layout: WorkbookPlacement = { sheets: plan.layout.sheets.map((sheet, index) => {
    if (!definition.sheets.has(sheet.name)) {
      return sheet
    }
    const events: FormMarker[] = [{ row: 1, token: [FormMarkerKind.Sheet, definition.signature] }]
    const visit = (regions: readonly WorkbookRegionLayout[]): void => {
      for (const region of regions) {
        const id = definition.numbers.get(region.definitionId)!
        events.push({ row: region.row, token: [FormMarkerKind.Repeat, id] })
        if (isWorkbookFormRow(definition.regions.get(region.definitionId)!)) {
          events.push({ row: region.row + region.height, token: [FormMarkerKind.RepeatEnd, id] })
          continue
        }
        for (const instance of region.instances) {
          events.push({ row: instance.row, token: [FormMarkerKind.Item, id] })
          visit(instance.regions)
          events.push({ row: instance.row + instance.height, token: [FormMarkerKind.ItemEnd, id] })
        }
        events.push({ row: region.row + region.height, token: [FormMarkerKind.RepeatEnd, id] })
      }
    }
    visit(plan.regions[index])
    let end = Math.max(2, plan.extents[index].rows + 1)
    let markerColumn = Math.max(FORM_MARKER_COLUMN, plan.extents[index].columns + 1)
    for (const row of sheet.rows ?? []) {
      end = Math.max(end, row.index + 1)
    }
    for (const event of events) {
      end = Math.max(end, event.row)
    }
    events.push({ row: end, token: [FormMarkerKind.SheetEnd] })
    // Stable order retains nesting when empty regions share a logical row.
    events.sort((a, b) => a.row - b.row)
    const placed = events.map((event, offset) => ({ ...event, row: event.row + offset }))
    if (placed.at(-1)!.row > WORKBOOK_LIMITS.rows) {
      throw new RangeError('Form control rows exceed the XLSX row limit')
    }
    for (const column of sheet.columns ?? []) {
      markerColumn = Math.max(markerColumn, column.index + 1)
    }
    if (markerColumn > WORKBOOK_LIMITS.columns) {
      throw new RangeError('No XLSX column remains for form control markers')
    }
    markers.set(sheet.name, { column: markerColumn, rows: placed })
    const placedRows = new Map<number, number>()
    const row = (logical: number) => {
      const cached = placedRows.get(logical)
      if (cached !== undefined) {
        return cached
      }
      let low = 0
      let high = events.length
      while (low < high) {
        const middle = (low + high) >>> 1
        if (events[middle].row <= logical) {
          low = middle + 1
        }
        else {
          high = middle
        }
      }
      placedRows.set(logical, logical + low)
      return logical + low
    }
    formulaRows.set(sheet.name.toLowerCase(), (logical, edge) => edge === FormulaEdge.End ? row(logical + 1) - 1 : row(logical))
    return { ...sheet, ...(sheet.print ? { print: mapWorkbookPrint(sheet.print, { rowStart: row }) } : {}), cells: sheet.cells.map(cell => {
      const start = row(cell.at.row)
      const rows = cell.size.rows === 1 ? 1 : row(cell.at.row + cell.size.rows - 1) - start + 1
      return { ...cell, at: { row: start, column: cell.at.column }, size: rows === cell.size.rows ? cell.size : { ...cell.size, rows } }
    }),
    rows: [...(sheet.rows ?? []).map(setting => ({ ...setting, index: row(setting.index) })), ...placed.map(marker => ({ index: marker.row, hidden: true }))].sort((a, b) => a.index - b.index),
    columns: [...sheet.columns ?? [], { index: markerColumn, hidden: true, width: 2 }],
    }
  }) }
  return { layout, markers, formulaRows }
}

export function writeFormMarkers(sheet: Worksheet, markers: FormMarkers): void {
  for (const marker of markers.rows) {
    const text = FORM_MARKER_PREFIX + JSON.stringify(marker.token)
    const issue = xlsxTextIssues(text)[0]
    if (issue) {
      throw new RangeError(issue.message)
    }
    sheet.getCell(marker.row, markers.column).value = text
  }
}

function formMarkerColumn(sheet: Worksheet): number {
  let column = FORM_MARKER_COLUMN
  sheet.getRow(1).eachCell(cell => {
    if (typeof cell.value === 'string' && cell.value.startsWith(FORM_MARKER_PREFIX)) {
      column = Number(cell.col)
    }
  })
  return column
}

function isFormMarkerToken(value: unknown): value is FormMarkerToken {
  if (!Array.isArray(value)) {
    return false
  }
  switch (value[0]) {
    case FormMarkerKind.Sheet:
      return value.length === 2 && typeof value[1] === 'string'
    case FormMarkerKind.SheetEnd:
      return value.length === 1
    case FormMarkerKind.Repeat:
    case FormMarkerKind.RepeatEnd:
    case FormMarkerKind.Item:
    case FormMarkerKind.ItemEnd:
      return value.length === 2 && Number.isSafeInteger(value[1]) && value[1] > 0
    default:
      return false
  }
}

export function readFormMarkers(sheet: Worksheet): FormMarkers {
  const column = formMarkerColumn(sheet)
  const result: FormMarker[] = []
  sheet.getColumn(column).eachCell(cell => {
    if (cell.value === null) {
      return
    }
    const fail = (): never => {
      throw new WorkbookFormInputError({ phase: 'structure', code: 'invalid-marker', path: '$workbook', sheetName: sheet.name, address: cell.address, message: 'form control markers must remain intact' })
    }
    const value = cell.value
    if (typeof value !== 'string' || !value.startsWith(FORM_MARKER_PREFIX)) {
      return fail()
    }
    let token: unknown
    try {
      token = JSON.parse(value.slice(FORM_MARKER_PREFIX.length))
    }
    catch {
      fail()
    }
    if (!isFormMarkerToken(token)) {
      return fail()
    }
    result.push({ row: Number(cell.row), token })
  })
  return { column, rows: result.sort((a, b) => a.row - b.row) }
}

/** Interpret only the declared template tree; marker text never supplies a data path or rule. */
export function formDataFromMarkers(template: WorkbookDefinition, definition: FormCarrierDefinition, markers: ReadonlyMap<string, FormMarkers>) {
  const shape = createWorkbookFormData()
  const rows: WorkbookFormRows[] = []
  for (const sheet of template.sheets) {
    if (!definition.sheets.has(sheet.name)) {
      continue
    }
    const { column, rows: found } = markers.get(sheet.name)!
    let cursor = 0
    const fail = (message: string): never => {
      throw new WorkbookFormInputError({ phase: 'structure', code: 'form-markers', path: '$workbook', sheetName: sheet.name,
        address: found[cursor] ? formatAddress({ row: found[cursor].row, column }) : undefined, message })
    }
    const expect = (...token: FormMarkerToken) => {
      const actual = found[cursor]?.token
      if (actual?.[0] !== token[0] || actual[1] !== token[1]) {
        fail('form boundaries must match the template structure')
      }
      cursor++
    }
    const visit = (body: WorkbookBody, context: DataPath): void => {
      shape.context(context)
      for (const cell of body.cells) {
        if ('path' in cell.value) {
          shape.field(referencePath(cell.value, context))
        }
      }
      for (const region of [...body.regions ?? []].sort((a, b) => a.row - b.row)) {
        const id = definition.numbers.get(region.id)!
        const start = found[cursor]?.row
        expect(FormMarkerKind.Repeat, id)
        let end = cursor
        let count = 0
        while (end < found.length && !(found[end].token[0] === FormMarkerKind.RepeatEnd && found[end].token[1] === id)) {
          if (found[end].token[0] === FormMarkerKind.Item && found[end].token[1] === id) {
            count++
          }
          end++
        }
        if (end === found.length) {
          fail('repeat closing boundary is missing')
        }
        const path = referencePath(region.source, context)
        const row = isWorkbookFormRow(region)
        if (row) {
          count = found[end].row - start! - 1
          rows.push({ path, fields: workbookFormRowFields(region) })
        }
        shape.collection(path, count)
        for (let index = 0; index < count; index++) {
          if (!row) {
            expect(FormMarkerKind.Item, id)
          }
          visit(region, [...path, index])
          if (!row) {
            expect(FormMarkerKind.ItemEnd, id)
          }
        }
        expect(FormMarkerKind.RepeatEnd, id)
      }
    }
    expect(FormMarkerKind.Sheet, definition.signature)
    visit(sheet, [])
    expect(FormMarkerKind.SheetEnd)
    if (cursor !== found.length) {
      fail('unexpected form control marker')
    }
  }
  return { data: shape.data, contexts: shape.contexts, rows }
}

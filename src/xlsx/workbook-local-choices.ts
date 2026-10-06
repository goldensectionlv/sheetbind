import type { Workbook, Worksheet } from 'exceljs'
import { isDataObject } from '../core/json'
import { hasValidation } from '../core/field-rules'
import type { WorkbookPlacement } from '../grid/workbook-layout'
import type { WorkbookFormField } from '../form/workbook-read'
import { writeData } from '../form/records'
import { WorkbookFormInputError } from '../form/workbook'
import { formatAddress, XLSX_MAX_COLUMN, XLSX_MAX_ROW } from './addresses'

interface LocalSource {
  readonly sheet: string
  readonly column: number
  readonly values: readonly (readonly Readonly<Record<string, unknown>>[])[]
}

/** A copied row carries its choice-source reference; no application record key is needed. */
export function writeLocalChoiceSources(book: Workbook, helper: Worksheet, sheets: WorkbookPlacement['sheets']): Record<string, LocalSource> {
  const result: Record<string, LocalSource> = {}
  const names = new Set(book.definedNames.model.map(entry => entry.name.toLowerCase()))
  let serial = 0
  function column(values: readonly string[]): string {
    const index = helper.columnCount + 1
    if (index > XLSX_MAX_COLUMN || values.length > XLSX_MAX_ROW) {
      throw new RangeError('Form choice sources exceed the XLSX worksheet limits')
    }
    helper.getCell(1, index).value = null
    values.forEach((value, row) => {
      helper.getCell(row + 1, index).value = value
    })
    let name: string
    do {
      name = `_sb_local_${++serial}`
    } while (names.has(name.toLowerCase()))
    names.add(name.toLowerCase())
    book.definedNames.add(`'${helper.name.replaceAll("'", "''")}'!${formatAddress({ row: 1, column: index })}:${formatAddress({ row: Math.max(1, values.length), column: index })}`, name)
    return name
  }
  for (const plan of sheets) {
    const sheet = book.getWorksheet(plan.name)!
    const fields = new Map<string, typeof plan.cells[number][]>()
    for (const cell of plan.cells) {
      const source = cell.rules?.choice?.source
      if (source && 'path' in source && source.from !== 'root') {
        const cells = fields.get(cell.definitionId) ?? []
        cells.push(cell)
        fields.set(cell.definitionId, cells)
      }
    }
    for (const [id, cells] of fields) {
      const index = sheet.columnCount + 1
      if (index > XLSX_MAX_COLUMN) {
        throw new RangeError('No XLSX column remains for local choice references')
      }
      sheet.getColumn(index).hidden = true
      const values: Readonly<Record<string, unknown>>[][] = []
      const lists: string[] = []
      const seen = new Map<string, number>()
      for (const cell of cells) {
        const items = cell.choice?.items ?? []
        const source = items.map(item => item.value)
        const token = JSON.stringify(source)
        let reference = seen.get(token)
        if (reference === undefined) {
          reference = values.length + 1
          seen.set(token, reference)
          values.push(source)
          // Blank lookup entries make Excel's IgnoreBlank accept unrelated values too.
          lists.push(!items.length && cell.rules?.choice?.emptySource === 'input'
            ? '#input'
            : column(items.map(item => item.text)))
        }
        sheet.getCell(cell.at.row, index).value = reference
      }
      const lookup = column(lists)
      for (const cell of cells) {
        const driver = formatAddress({ row: cell.at.row, column: index }).replace(/^([A-Z]+)/, '$$$1')
        const range = `INDEX(${lookup},${driver})`
        const target = sheet.getCell(formatAddress(cell.at))
        const strictEmpty = cell.rules?.choice?.emptySource !== 'input' && values.some(items => !items.length)
        target.dataValidation = { ...target.dataValidation, type: 'list', allowBlank: !hasValidation(cell.rules, 'required') && !strictEmpty,
          showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Choose a listed value', error: 'Select a value from the dropdown list.',
          formulae: [`INDIRECT(IF(${range}="#input",ADDRESS(ROW(),COLUMN()),${range}))`] }
      }
      result[id] = { sheet: plan.name, column: index, values }
    }
  }
  return result
}

/** Resolve only a declared field's source; uploaded metadata cannot supply template paths. */
export function readLocalChoiceContext(book: Workbook, field: WorkbookFormField, local: Readonly<Record<string, unknown>>): Readonly<Record<string, unknown>> {
  try {
    const source = local[field.id]
    if (!isDataObject(source) || source.sheet !== field.location.sheetName || !Number.isSafeInteger(source.column)
      || (source.column as number) < 1 || (source.column as number) > XLSX_MAX_COLUMN || !Array.isArray(source.values)) {
      throw new Error('invalid local source')
    }
    const sheet = book.getWorksheet(field.location.sheetName)!
    const row = Number(sheet.getCell(field.location.address).row)
    const index = sheet.getCell(row, source.column as number).value
    if (!Number.isSafeInteger(index) || (index as number) < 1 || (index as number) > source.values.length) {
      throw new Error('missing local reference')
    }
    const values = source.values[(index as number) - 1]
    if (!Array.isArray(values) || !values.every(isDataObject)) {
      throw new Error('invalid local choices')
    }
    const rule = field.rules!.choice!
    const context: Record<string, unknown> = {}
    if ('path' in rule.source) {
      writeData(context, rule.source.path.split('.'), values, true)
    }
    return context
  }
  catch {
    throw new WorkbookFormInputError({ phase: 'xlsx', code: 'choice-source', ...field.location,
      message: 'local choice source is missing or malformed; copy the whole row or record block' })
  }
}

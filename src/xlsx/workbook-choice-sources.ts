import type { Workbook } from 'exceljs'
import { parseDictionaries } from '../core/dictionaries'
import { jsonSnapshot, parseJsonObject } from '../core/json'
import type { WorkbookChoiceSources } from '../form/workbook-choice-sources'
import { WorkbookFormInputError } from '../form/workbook'
import { formatAddress, parseRange, XLSX_MAX_COLUMN, XLSX_MAX_ROW } from './addresses'
import type { WorkbookPlacement } from '../grid/workbook-layout'
import { writeLocalChoiceSources } from './workbook-local-choices'

const NAME = '_sb_object_sources'
const VERSION = 'sheetbind.choices/1'

/** Store the issued lists and choice payloads independently of the return mode. */
export function writeWorkbookChoiceSources(book: Workbook, sheetName: string, sources: WorkbookChoiceSources, sheets: WorkbookPlacement['sheets']): void {
  if (book.definedNames.model.some(entry => entry.name.toLowerCase() === NAME)) {
    throw new RangeError('Object-choice source name is already defined')
  }
  const sheet = book.getWorksheet(sheetName) ?? book.addWorksheet(sheetName, { state: 'veryHidden' })
  const local = writeLocalChoiceSources(book, sheet, sheets)
  const payload = jsonSnapshot({ ...sources, local })
  const json = JSON.stringify(payload).replace(/[^\x20-\x7e]/g, character => `\\u${character.charCodeAt(0).toString(16).padStart(4, '0')}`)
  const column = sheet.columnCount + 1
  const chunks = [VERSION, ...json.match(/.{1,30000}/g)!]
  if (column > XLSX_MAX_COLUMN || chunks.length > XLSX_MAX_ROW) {
    throw new RangeError('Form source payload exceeds the XLSX worksheet limits')
  }
  chunks.forEach((text, index) => {
    const cell = sheet.getCell(index + 1, column)
    cell.value = text
    cell.numFmt = '@'
  })
  book.definedNames.add(`'${sheetName}'!${formatAddress({ row: 1, column })}:${formatAddress({ row: chunks.length, column })}`, NAME)
}

export function readWorkbookChoiceSources(book: Workbook, sheetName: string): WorkbookChoiceSources & { readonly local: Readonly<Record<string, unknown>> } {
  try {
    const ranges = book.definedNames.getRanges(NAME).ranges
    if (ranges.length !== 1) {
      throw new Error('missing source range')
    }
    const prefix = `'${sheetName}'!`
    const range = ranges[0]
    if (!range.startsWith(prefix)) {
      throw new Error('source range is on another sheet')
    }
    const { start, end } = parseRange(range.slice(prefix.length).replace(/\$/g, ''))
    if (start.column !== end.column) {
      throw new Error('invalid source range')
    }
    const sheet = book.getWorksheet(sheetName)!
    if (sheet.getCell(start.row, start.column).value !== VERSION) {
      throw new Error('unsupported source version')
    }
    let text = ''
    for (let row = start.row + 1; row <= end.row; row++) {
      const value = sheet.getCell(row, start.column).value
      if (typeof value !== 'string' || value.length > 30000) {
        throw new Error('invalid source payload')
      }
      text += value
    }
    const payload = parseJsonObject(JSON.parse(text))
    if (Object.keys(payload).some(key => !['context', 'dictionaries', 'local'].includes(key))) {
      throw new Error('unknown source property')
    }
    return { context: parseJsonObject(payload.context), dictionaries: parseDictionaries(payload.dictionaries), local: parseJsonObject(payload.local ?? {}) }
  }
  catch {
    throw new WorkbookFormInputError({ phase: 'xlsx', code: 'choice-source', path: '$workbook', message: 'form dictionary source data is missing or malformed' })
  }
}

import type { Workbook } from 'exceljs'
import { parseDictionaries } from '../core/dictionaries'
import { jsonSnapshot, assertJson, isDataObject } from '../core/json'
import type { WorkbookChoiceSources } from '../form/workbook-choice-sources'
import { WorkbookFormInputError } from '../form/workbook'
import { formatAddress, parseRange, XLSX_MAX_COLUMN, XLSX_MAX_ROW } from './addresses'
import type { WorkbookResources } from './workbook-resources'

const NAME = '_sb_object_sources'
const VERSION = 'sheetbind.choices/2'

/** Store the issued lists and choice payloads independently of the return mode. */
export function writeWorkbookChoiceSources(book: Workbook, sheetName: string, sources: WorkbookChoiceSources, resources: WorkbookResources): void {
  const name = resources.allocate(NAME)
  const sheet = book.getWorksheet(sheetName) ?? book.addWorksheet(sheetName, { state: 'veryHidden' })
  const payload = jsonSnapshot(sources)
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
  book.definedNames.add(`'${sheetName}'!${formatAddress({ row: 1, column })}:${formatAddress({ row: chunks.length, column })}`, name)
}

export function readWorkbookChoiceSources(book: Workbook, sheetName: string): Omit<WorkbookChoiceSources, 'local'> & { readonly local: Readonly<Record<string, unknown>> } {
  try {
    const prefix = `'${sheetName}'!`
    const ranges = book.definedNames.model.filter(entry => /^_sb_object_sources(?:_\d+)?$/i.test(entry.name))
      .flatMap(entry => entry.ranges).filter(range => range.startsWith(prefix))
    if (ranges.length !== 1) {
      throw new Error('missing source range')
    }
    const range = ranges[0]
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
    const payload: unknown = JSON.parse(text)
    assertJson(payload)
    if (!isDataObject(payload) || !isDataObject(payload.context) || payload.local !== undefined && !isDataObject(payload.local)) {
      throw new Error('invalid source payload')
    }
    if (Object.keys(payload).some(key => !['context', 'dictionaries', 'local'].includes(key))) {
      throw new Error('unknown source property')
    }
    return { context: payload.context, dictionaries: parseDictionaries(payload.dictionaries), local: payload.local ?? {} }
  }
  catch {
    throw new WorkbookFormInputError({ phase: 'xlsx', code: 'choice-source', path: '$workbook', message: 'form dictionary source data is missing or malformed' })
  }
}

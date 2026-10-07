import type { Workbook } from 'exceljs'
import { validateDictionaryShape } from '../core/dictionaries'
import type { Dictionaries } from '../core/dictionaries'
import { createChoiceResolver, createWorkbookChoiceDisplay } from '../core/choices'
import type { ChoiceRule } from '../core/choices'
import { hasValidation } from '../core/field-rules'
import { assertInputData, assertJson, isDataObject } from '../core/json'
import type { TemplateValue } from '../core/template'
import { workbookCells } from '../grid/workbook'
import type { GridRange } from '../grid/geometry'
import type { WorkbookPlan } from '../grid/workbook-layout'
import { formatAddress, formatRange, parseRange, XLSX_MAX_COLUMN, XLSX_MAX_ROW } from '../grid/geometry'
import { WorkbookFormInputError } from './form-definition'
import { readData, writeData } from './form-records'
import { cellXml } from './workbook-cells'
import type { WorkbookCells } from './workbook-cells'
import type { WorkbookResources } from './workbook-resources'
import { protect } from './report-text'
import { encodeXml, setXmlAttributes } from './xml'

const SOURCE_NAME = '_sb_object_sources'
const SOURCE_VERSION = 'sheetbind.choices/2'

// Escaping underscores and uppercase too prevents name collisions in case-insensitive Excel.
function namePart(value: string): string {
  return value.replace(/[^a-z0-9]/g, character => `_u${character.charCodeAt(0).toString(16).padStart(4, '0')}_`)
}
function rangeName(prefix: string, field?: string): string {
  return `${prefix}__${field === undefined ? 'text' : `field_${namePart(field)}`}`
}

/** Stable field names for a named dictionary projection, independent of object key order. */
export function workbookChoiceRange(rule: ChoiceRule, field?: string): string {
  const name = rangeName(choicePrefix(rule), field)
  if (name.length > 255) {
    throw new RangeError('Choice field reference exceeds the Excel defined-name limit')
  }
  return name
}

function choicePrefix(rule: ChoiceRule): string {
  if (!('dictionary' in rule.source)) {
    throw new RangeError('A stable choice range requires a named dictionary source')
  }
  const parts = [rule.source.dictionary, rule.key, rule.label].map(namePart)
  const prefix = parts.map(part => `${part.length}_${part}`).join('_')
  return `_sb_ref_${prefix}`
}

function fieldValue(value: unknown): string | number | boolean | null {
  if (value === undefined || value === null) {
    return null
  }
  return typeof value === 'object' ? JSON.stringify(value) : value as string | number | boolean
}

export function workbookListSheetName(authoredNames: readonly string[]): string {
  const names = new Set(authoredNames.map(name => name.toLowerCase()))
  let name = '_sheetbind_lists'
  let suffix = 1
  while (names.has(name.toLowerCase())) {
    name = `_sheetbind_lists_${++suffix}`
  }
  return name
}

/** One hidden sheet contains named lists, referenced dictionary fields and issued form sources. */
export function workbookLists(plan: WorkbookPlan, dictionaries: Dictionaries, resources: WorkbookResources,
  cells: WorkbookCells, sources?: string) {
  const name = workbookListSheetName(plan.sheets.map(({ sheet }) => sheet.name))
  const rows: string[][] = []
  const names: string[] = []
  let column = 0
  function writeColumn(reference: string, values: readonly TemplateValue[], header?: string): void {
    const start = header === undefined ? 1 : 2
    const end = Math.max(start, start + values.length - 1)
    if (++column > XLSX_MAX_COLUMN || end > XLSX_MAX_ROW) {
      throw new RangeError('Workbook lists exceed the XLSX worksheet limits')
    }
    function write(row: number, value: TemplateValue) {
      const content = cells.content(value)
      const target = rows[row - 1] ??= []
      target.push(cellXml(formatAddress({ row, column }), { ...content, style: typeof value === 'string' ? cells.textStyle() : cells.style(0) }))
    }
    if (header !== undefined) {
      write(1, header)
    }
    values.forEach((value, index) => write(start + index, value))
    const absolute = (row: number) => formatAddress({ row, column }).replace(/^([A-Z]+)(\d+)$/, (_, col, row) => `$${col}$${row}`)
    const range = `'${name}'!${absolute(start)}:${absolute(end)}`
    names.push(`<definedName name="${encodeXml(reference)}">${encodeXml(protect(range))}</definedName>`)
  }
  const texts = new WeakMap<object, readonly string[]>()
  const references = new Map<string, string>()
  const byItems = new WeakMap<readonly string[], string>()
  const validations = new Map<string, string>()
  for (const { sheet } of plan.sheets) {
    const ranges: { range: GridRange, xml: string }[] = []
    const last = new Map<number, typeof ranges[number]>()
    for (const cell of sheet.cells) {
      let items: readonly string[]
      if (cell.choice) {
        items = texts.get(cell.choice.items) ?? cell.choice.items.map(item => item.text)
        texts.set(cell.choice.items, items)
      }
      else if (cell.rules?.list) {
        items = dictionaries[cell.rules.list] as readonly string[]
      }
      else {
        continue
      }
      let reference = byItems.get(items)
      if (!reference) {
        const key = JSON.stringify(items)
        reference = references.get(key)
        if (!reference) {
          reference = resources.allocate(`_sb_list_${references.size + 1}`)
          writeColumn(reference, items)
          references.set(key, reference)
        }
        byItems.set(items, reference)
      }
      const xml = `<dataValidation type="${items.length ? 'list' : 'custom'}" allowBlank="${hasValidation(cell.rules, 'required') ? 0 : 1}" showErrorMessage="1" errorStyle="stop" errorTitle="Choose a listed value" error="Select a value from the dropdown list."><formula1>${items.length ? reference : 'FALSE'}</formula1></dataValidation>`
      const previous = last.get(cell.at.column)
      if (previous && previous.range.end.row + 1 === cell.at.row && previous.xml === xml) {
        previous.range = { start: previous.range.start, end: cell.at }
      }
      else {
        const entry = { range: { start: cell.at, end: cell.at }, xml }
        ranges.push(entry)
        last.set(cell.at.column, entry)
      }
    }
    validations.set(sheet.name, ranges.map(({ range, xml }) => setXmlAttributes(xml, { sqref: formatRange(range) })).join(''))
  }
  const rules = plan.sheets.flatMap(({ definition }) => workbookCells(definition).flatMap(cell => cell.rules?.choice ? [cell.rules.choice] : []))
  // Formula references declare dictionary columns independently of the issued rows.
  const named = new Set<string>()
  const resolve = createChoiceResolver()
  const display = createWorkbookChoiceDisplay()
  for (const rule of rules) {
    if (!('dictionary' in rule.source)) {
      continue
    }
    const prefix = choicePrefix(rule)
    if (named.has(prefix)) {
      continue
    }
    named.add(prefix)
    const fieldPrefix = prefix.toLowerCase() + '__field_'
    const fields = [...resources.references].filter(name => name.startsWith(fieldPrefix))
      .map(name => name.slice(fieldPrefix.length).replace(/_u([0-9a-f]{4})_/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16))))
      .filter(field => resources.references.has(rangeName(prefix, field).toLowerCase())).sort()
    if (!fields.length && !resources.references.has(rangeName(prefix).toLowerCase())) {
      continue
    }
    const items = display({ key: null, items: resolve(rule, dictionaries[rule.source.dictionary]) }).items
    writeColumn(resources.allocate(rangeName(prefix), true), items.map(item => item.text), 'Selection')
    for (const field of fields) {
      writeColumn(resources.allocate(rangeName(prefix, field), true), items.map(item => fieldValue(item.value[field])), field)
    }
  }
  if (sources) {
    const json = sources.replace(/[^\x20-\x7e]/g, character => `\\u${character.charCodeAt(0).toString(16).padStart(4, '0')}`)
    const chunks = [SOURCE_VERSION, ...json.match(/.{1,30000}/g)!]
    writeColumn(resources.allocate(SOURCE_NAME), chunks)
  }
  const xml = column
    ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${formatAddress({ row: Math.max(1, rows.length), column })}"/><sheetData>${rows.map((cells, row) => `<row r="${row + 1}">${cells.join('')}</row>`).join('')}</sheetData></worksheet>`
    : undefined
  return { name, xml, names, validations }
}

export function readWorkbookChoiceSources(book: Workbook, sheetName: string) {
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
    if (sheet.getCell(start.row, start.column).value !== SOURCE_VERSION) {
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
    return { context: payload.context, dictionaries: validateDictionaryShape(payload.dictionaries), local: payload.local ?? {} }
  }
  catch {
    throw new WorkbookFormInputError({ phase: 'xlsx', code: 'choice-source', path: '$workbook', message: 'form dictionary source data is missing or malformed' })
  }
}
/** Capture issued sources as JSON before asynchronous package writing can outlive the input. */
export function serializeWorkbookChoiceSources(plan: WorkbookPlan, data: unknown, dictionaries: Dictionaries): string | undefined {
  const context: Record<string, unknown> = {}
  const selected: Record<string, Dictionaries[string]> = {}
  const fields = plan.sheets.flatMap(({ definition }) => workbookCells(definition)).filter(cell => cell.rules?.choice || cell.rules?.list)
  if (!fields.length) {
    return undefined
  }
  for (const cell of fields) {
    if (cell.rules?.list) {
      selected[cell.rules.list] = dictionaries[cell.rules.list]
    }
    const source = cell.rules?.choice?.source
    if (!source) {
      continue
    }
    if ('dictionary' in source) {
      selected[source.dictionary] = dictionaries[source.dictionary]
    }
    else if (source.from === 'root') {
      const path = source.path.split('.')
      writeData(context, path, readData(data, path) ?? [], true)
    }
  }
  const local: Record<string, Record<string, readonly Readonly<Record<string, unknown>>[]>> = {}
  for (const { sheet } of plan.sheets) {
    for (const cell of sheet.cells) {
      const source = cell.rules?.choice?.source
      if (source && 'path' in source && source.from !== 'root') {
        const fields = local[cell.definitionId] ??= {}
        fields[cell.origin.dataPath] = cell.choice?.items.map(item => item.value) ?? []
      }
    }
  }
  const sources = { context, dictionaries: selected, local }
  assertInputData(sources)
  return JSON.stringify(sources)
}

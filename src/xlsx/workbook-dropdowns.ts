import { hasValidation } from '../core/field-rules'
import type { DataValidation, Workbook, Worksheet } from 'exceljs'
import { equalJson } from '../core/json'
import type { GridAddress } from '../grid/geometry'
import type { FieldRules } from '../core/field-rules'
import { parseDictionaries } from '../core/dictionaries'
import type { Dictionaries } from '../core/dictionaries'
import { formatAddress, formatRange, parseAddress } from './addresses'
import { assertXlsxText } from './report-text'

export interface DropdownTarget {
  readonly sheet: Worksheet
  readonly address: string
  readonly rules: Pick<FieldRules, 'validation' | 'list'>
  readonly items?: readonly string[]
}
interface PreparedDropdown {
  readonly target: DropdownTarget
  readonly source: string
}

/** ExcelJS keeps validation independently of worksheet cell objects. */
function validations(sheet: Worksheet) {
  return Reflect.get(sheet, 'dataValidations') as { model: Record<string, DataValidation>, add(address: string, value: DataValidation): void }
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

/** Install lists only after placement. Template files contain named dependencies, not these values. */
function prepareDropdowns(book: Workbook, targets: readonly DropdownTarget[], values: Dictionaries) {
  const dictionaries = parseDictionaries(values)
  const sources = new Map<string, readonly string[]>()
  const fields: PreparedDropdown[] = []
  const sourceIds = new WeakMap<readonly string[], string>()
  function addSource(items: readonly string[], target: DropdownTarget): string {
    const source = sourceIds.get(items) ?? JSON.stringify(items)
    sourceIds.set(items, source)
    if (sources.has(source)) {
      return source
    }
    for (const item of items) {
      assertXlsxText(item, `Dropdown ${target.sheet.name}!${target.address}`)
    }
    sources.set(source, items)
    return source
  }
  for (const target of targets) {
    const values = target.items ?? dictionaries[target.rules.list!]
    if (!values || !target.items && !values.length || values.some(value => typeof value !== 'string')) {
      throw new RangeError(`Dictionary ${target.rules.list} must contain string values`)
    }
    const items = values as readonly string[]
    const source = addSource(items, target)
    fields.push({ target, source })
  }
  for (const target of targets) {
    if (target.sheet.workbook !== book) {
      throw new RangeError('Dropdown target belongs to another workbook')
    }
    const existing = validations(target.sheet).model[target.address]
    // ExcelJS supports prompt-only 'any' at runtime, but omits it from its type union.
    if (existing?.type && String(existing.type) !== 'any') {
      throw new RangeError(`Existing validation at ${target.sheet.name}!${target.address}`)
    }
  }
  return { sources, fields }
}

/** Preflight precedes all writes; an existing validation is never silently replaced. */
export function writeWorkbookDropdowns(book: Workbook, targets: readonly DropdownTarget[], values: Dictionaries = {}): Worksheet | undefined {
  if (!targets.length) {
    return
  }
  const { sources, fields } = prepareDropdowns(book, targets, values)
  const sheetName = workbookListSheetName(book.worksheets.map(sheet => sheet.name))
  const sheet = book.addWorksheet(sheetName, { state: 'veryHidden' })
  const names = new Set(book.definedNames.model.map(entry => entry.name.toLowerCase()))
  const references = new Map<string, string>()
  let serial = 0
  function nextName(prefix: string): string {
    let name: string
    do {
      name = `${prefix}${++serial}`
    } while (names.has(name.toLowerCase()))
    names.add(name.toLowerCase())
    return name
  }
  for (const [index, [source, items]] of [...sources].entries()) {
    const name = nextName('_sb_list_')
    references.set(source, name)
    const column = index + 1
    for (const [offset, item] of items.entries()) {
      const cell = sheet.getCell(offset + 1, column)
      cell.value = item
      cell.numFmt = '@'
    }
    const start = formatAddress({ row: 1, column })
    const end = formatAddress({ row: Math.max(1, items.length), column })
    book.definedNames.add(`'${sheetName}'!${start}:${end}`, name)
  }
  const ranges: { sheet: Worksheet, range: { start: GridAddress, end: GridAddress }, validation: DataValidation }[] = []
  const last = new Map<Worksheet, Map<number, typeof ranges[number]>>()
  const rules = new WeakMap<DropdownTarget['rules'], Map<string, DataValidation>>()
  for (const { target, source } of fields) {
    const reference = references.get(source)
    const empty = !sources.get(source)!.length
    const existing = validations(target.sheet).model[target.address]
    const expression = empty ? 'FALSE' : reference!
    const cache = rules.get(target.rules) ?? new Map<string, DataValidation>()
    let validation = existing ? undefined : cache.get(expression)
    if (!validation) {
      validation = { showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Choose a listed value', error: 'Select a value from the dropdown list.',
        ...structuredClone(existing ?? {}), type: empty ? 'custom' : 'list', allowBlank: !hasValidation(target.rules, 'required'), formulae: [expression] }
      if (!existing) {
        cache.set(expression, validation)
        rules.set(target.rules, cache)
      }
    }
    const at = parseAddress(target.address)
    const columns = last.get(target.sheet) ?? new Map<number, typeof ranges[number]>()
    const previous = columns.get(at.column)
    if (previous && previous.range.end.row + 1 === at.row && equalJson(previous.validation, validation)) {
      previous.range.end = at
    }
    else {
      const entry = { sheet: target.sheet, range: { start: at, end: { ...at } }, validation }
      ranges.push(entry)
      columns.set(at.column, entry)
      last.set(target.sheet, columns)
    }
  }
  for (const { target } of fields) {
    delete validations(target.sheet).model[target.address]
  }
  for (const { sheet, range, validation } of ranges) {
    validations(sheet).add(formatRange(range), validation)
  }
  return sheet
}

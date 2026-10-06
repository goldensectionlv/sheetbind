import type { Workbook, Worksheet } from 'exceljs'
import type { ChoiceRule } from '../core/choices'
import type { WorkbookChoiceOption } from '../grid/workbook-choice-display'
import { equalJson } from '../core/json'
import { formatAddress, XLSX_MAX_COLUMN } from './addresses'
import { assertXlsxText } from './report-text'
import type { WorkbookResources } from './workbook-resources'

export interface WorkbookChoiceTarget {
  readonly rule: ChoiceRule
  readonly items: readonly WorkbookChoiceOption[]
}

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
  const result = typeof value === 'object' ? JSON.stringify(value) : value as string | number | boolean
  if (typeof result === 'string') {
    assertXlsxText(result)
  }
  return result
}

/** Full objects remain in the source payload; this flat projection is for Excel formulas. */
export function writeWorkbookChoiceFields(book: Workbook, sheet: Worksheet, targets: readonly WorkbookChoiceTarget[], resources: WorkbookResources): void {
  const named = new Map<string, readonly WorkbookChoiceOption[]>()
  const sources: { headers: readonly string[], rows: readonly ReturnType<typeof fieldValue>[][], names: readonly string[] }[] = []
  let columns = sheet.columnCount
  for (const { rule, items } of targets) {
    let prefix: string
    if ('dictionary' in rule.source) {
      prefix = choicePrefix(rule)
      if (![...resources.references].some(name => name.startsWith(prefix.toLowerCase() + '__'))) {
        continue
      }
      const name = rangeName(prefix)
      const previous = named.get(name)
      if (previous) {
        if (!equalJson(previous, items)) {
          throw new RangeError('A named choice source has conflicting objects')
        }
        continue
      }
      named.set(name, items)
    }
    else {
      continue
    }
    const fields = [...new Set(items.flatMap(item => Object.keys(item.value)))].filter(field => resources.references.has(rangeName(prefix, field).toLowerCase())).sort()
    if (!fields.length && !resources.references.has(rangeName(prefix).toLowerCase())) {
      continue
    }
    columns += fields.length + 1
    if (columns > XLSX_MAX_COLUMN) {
      throw new RangeError('Object-choice field ranges exceed the XLSX column limit')
    }
    const references = [rangeName(prefix), ...fields.map(field => rangeName(prefix, field))]
    fields.forEach(fieldValue)
    for (const name of references) {
      resources.allocate(name, true)
    }
    const rows = items.map(item => [fieldValue(item.text), ...fields.map(field => fieldValue(item.value[field]))])
    sources.push({ headers: ['Selection', ...fields], rows, names: references })
  }
  let column = sheet.columnCount + 1
  for (const source of sources) {
    const { headers, rows } = source
    headers.forEach((header, offset) => {
      sheet.getCell(1, column + offset).value = header
    })
    for (const [index, values] of rows.entries()) {
      values.forEach((value, offset) => {
        const cell = sheet.getCell(index + 2, column + offset)
        cell.value = value
        if (typeof value === 'string') {
          cell.numFmt = '@'
        }
      })
    }
    source.names.forEach((name, offset) => {
      const start = formatAddress({ row: 2, column: column + offset })
      const end = formatAddress({ row: Math.max(2, rows.length + 1), column: column + offset })
      book.definedNames.add(`'${sheet.name}'!${start}:${end}`, name)
    })
    column += headers.length
  }
}

import type { Workbook, Worksheet } from 'exceljs'
import type { ChoiceRule } from '../core/choices'
import type { WorkbookChoiceOption } from '../grid/workbook-choice-display'
import { equalJson } from '../core/json'
import { formatAddress, XLSX_MAX_COLUMN } from './addresses'
import { xlsxTextIssues } from './report-text'

export interface WorkbookChoiceTarget {
  readonly rule: ChoiceRule
  readonly items: readonly WorkbookChoiceOption[]
}

// Escaping underscores and uppercase too prevents name collisions in case-insensitive Excel.
function namePart(value: string): string {
  return value.replace(/[^a-z0-9]/g, character => `_u${character.charCodeAt(0).toString(16).padStart(4, '0')}_`)
}
function rangeName(prefix: string, field?: string): string {
  const name = `${prefix}__${field === undefined ? 'text' : `field_${namePart(field)}`}`
  if (name.length > 255) {
    throw new RangeError('Choice field reference exceeds the Excel defined-name limit')
  }
  return name
}

/** Stable field names for a named dictionary projection, independent of object key order. */
export function workbookChoiceRange(rule: ChoiceRule, field?: string): string {
  if (!('dictionary' in rule.source)) {
    throw new RangeError('A stable choice range requires a named dictionary source')
  }
  const parts = [rule.source.dictionary, rule.key, rule.label].map(namePart)
  const prefix = parts.map(part => `${part.length}_${part}`).join('_')
  return rangeName(`_sb_ref_${prefix}`, field)
}

function fieldValue(value: unknown): string | number | boolean | null {
  if (value === undefined || value === null) {
    return null
  }
  if (typeof value === 'number' && Number(value.toPrecision(15)) !== value) {
    throw new RangeError('Numeric choice fields must fit Excel precision; use a string for longer identifiers')
  }
  const result = typeof value === 'object' ? JSON.stringify(value) : value as string | number | boolean
  if (typeof result === 'string') {
    const issue = xlsxTextIssues(result)[0]
    if (issue) {
      throw new RangeError(issue.message)
    }
  }
  return result
}

/** Full objects remain in the source payload; this flat projection is for Excel formulas. */
export function writeWorkbookChoiceFields(book: Workbook, sheet: Worksheet, targets: readonly WorkbookChoiceTarget[]): void {
  const named = new Map<string, readonly WorkbookChoiceOption[]>()
  const contextual = new WeakSet<readonly WorkbookChoiceOption[]>()
  const sources: { headers: readonly string[], rows: readonly ReturnType<typeof fieldValue>[][], names: readonly string[] }[] = []
  const names = new Set(book.definedNames.model.map(entry => entry.name.toLowerCase()))
  let columns = sheet.columnCount
  for (const { rule, items } of targets) {
    let prefix: string
    if ('dictionary' in rule.source) {
      const name = workbookChoiceRange(rule)
      const previous = named.get(name)
      if (previous) {
        if (!equalJson(previous, items)) {
          throw new RangeError('A named choice source has conflicting objects')
        }
        continue
      }
      named.set(name, items)
      prefix = name.slice(0, -6)
    }
    else {
      if (contextual.has(items)) {
        continue
      }
      prefix = `_sb_context_${sources.length + 1}`
      contextual.add(items)
    }
    const fields = [...new Set(items.flatMap(item => Object.keys(item.value)))].sort()
    columns += fields.length + 1
    if (columns > XLSX_MAX_COLUMN) {
      throw new RangeError('Object-choice field ranges exceed the XLSX column limit')
    }
    const references = [rangeName(prefix), ...fields.map(field => rangeName(prefix, field))]
    fields.forEach(fieldValue)
    for (const name of references) {
      if (names.has(name.toLowerCase())) {
        throw new RangeError(`Choice range already exists: ${name}`)
      }
      names.add(name.toLowerCase())
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

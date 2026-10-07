import type { ChoiceRule } from '../core/choices'
import { createChoiceResolver } from '../core/choices'
import type { Dictionaries } from '../core/dictionaries'
import { createWorkbookChoiceDisplay } from '../grid/workbook-choice-display'
import type { WorkbookResources } from './workbook-resources'
import type { WriteWorkbookListColumn } from './workbook-lists'

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

/** Formula references declare dictionary columns, even when no input rows or source items exist. */
export function writeWorkbookChoiceFields(rules: readonly ChoiceRule[], dictionaries: Dictionaries, resources: WorkbookResources, writeColumn: WriteWorkbookListColumn): void {
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
}

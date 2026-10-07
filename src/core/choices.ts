import { isDataObject } from './json'
import { readDataPath } from './template'
import type { DataReference } from './template'

export interface ChoiceRule {
  readonly source: { readonly dictionary: string } | DataReference
  readonly key: string
  readonly label: string
  readonly return?: 'object' | 'key'
  readonly emptySource?: 'input'
}
export function returnsObject(rule: ChoiceRule | undefined): boolean {
  return !!rule && rule.return !== 'key'
}
export function allowsChoiceInput(rule: ChoiceRule, items: readonly ChoiceOption[]): boolean {
  return rule.emptySource === 'input' && !items.length
}
export interface ChoiceOption { readonly key: string | number, readonly label: string, readonly value: Readonly<Record<string, unknown>> }

/** Options borrow validated source records; consumers copy only values they retain. */
function buildOptions(source: readonly unknown[], rule: ChoiceRule): ChoiceOption[] {
  const keys = new Set<string>()
  return source.map(item => {
    const key = readDataPath(item, rule.key)
    const label = readDataPath(item, rule.label)
    if (!isDataObject(item) || !['string', 'number'].includes(typeof key) || typeof key === 'string' && !key.trim() || typeof key === 'number' && !Number.isFinite(key)) {
      throw new SyntaxError('Choice keys must be nonempty strings or finite numbers')
    }
    if (typeof label !== 'string' || !label.trim()) {
      throw new SyntaxError('Choice labels must be nonempty strings')
    }
    const token = JSON.stringify(key)
    if (keys.has(token)) {
      throw new SyntaxError('Choice keys must be unique')
    }
    keys.add(token)
    return { key: key as string | number, label, value: item }
  })
}

/** A cache belongs to one execution; the same source can have independent projections. */
export function createChoiceResolver() {
  const sources = new WeakMap<object, Map<string, readonly ChoiceOption[]>>()
  return function resolveChoice(rule: ChoiceRule, source: unknown): readonly ChoiceOption[] {
    if (!Array.isArray(source)) {
      throw new SyntaxError('Choice source must be an array of objects')
    }
    const projections = sources.get(source) ?? new Map<string, readonly ChoiceOption[]>()
    const projection = JSON.stringify([rule.key, rule.label])
    if (!projections.has(projection)) {
      projections.set(projection, buildOptions(source, rule))
    }
    sources.set(source, projections)
    return projections.get(projection)!
  }
}

export function choiceKey(rule: ChoiceRule, value: unknown): unknown {
  return returnsObject(rule) ? readDataPath(value, rule.key) : value
}
interface WorkbookChoiceOption extends ChoiceOption { readonly text: string }
export interface WorkbookChoice { readonly text: string | null, readonly items: readonly WorkbookChoiceOption[] }

/** Workbook layouts and files use the same unambiguous choice labels. */
export function createChoiceLabels() {
  const sources = new WeakMap<readonly ChoiceOption[], readonly WorkbookChoiceOption[]>()
  return (options: readonly ChoiceOption[]): readonly WorkbookChoiceOption[] => {
    let items = sources.get(options)
    if (!items) {
      const counts = new Map<string, number>()
      const texts = new Set<string>()
      options.forEach(option => counts.set(option.label.toLowerCase(), (counts.get(option.label.toLowerCase()) ?? 0) + 1))
      items = options.map(option => {
        const text = counts.get(option.label.toLowerCase())! > 1 ? `${option.label} [${option.key}]` : option.label
        if (texts.has(text.toLowerCase())) {
          throw new SyntaxError('Choice display labels are ambiguous after adding keys')
        }
        texts.add(text.toLowerCase())
        return { ...option, text }
      })
      sources.set(options, items)
    }
    return items
  }
}

import { isDataObject } from './json'
import { readDataPath } from './template'
import type { DataReference, TemplateValue } from './template'

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

interface ChoiceResolution { readonly items: readonly WorkbookChoiceOption[], readonly problem?: string }

/** Resolve one projection. Rendering can omit an unusable list; reading must reject a damaged mapping. */
export function createChoiceResolver() {
  const sources = new WeakMap<object, Map<string, ChoiceResolution>>()
  return function resolveChoice(rule: ChoiceRule, source: unknown): ChoiceResolution {
    if (source == null) {
      return { items: [] }
    }
    if (!Array.isArray(source)) {
      return { items: [], problem: 'Choice source must be an array of objects' }
    }
    const projections = sources.get(source) ?? new Map<string, ChoiceResolution>()
    const projection = JSON.stringify([rule.key, rule.label])
    if (!projections.has(projection)) {
      try {
        projections.set(projection, { items: choiceLabels(buildOptions(source, rule)) })
      }
      catch (error) {
        if (!(error instanceof SyntaxError)) {
          throw error
        }
        projections.set(projection, { items: [], problem: error.message })
      }
    }
    sources.set(source, projections)
    return projections.get(projection)!
  }
}

export function choiceKey(rule: ChoiceRule, value: unknown): unknown {
  return returnsObject(rule) ? readDataPath(value, rule.key) : value
}
interface WorkbookChoiceOption extends ChoiceOption { readonly text: string }
export interface WorkbookChoice { readonly text: TemplateValue, readonly items: readonly WorkbookChoiceOption[] }

/** Workbook layouts and files use the same unambiguous choice labels. */
function choiceLabels(options: readonly ChoiceOption[]): readonly WorkbookChoiceOption[] {
  const counts = new Map<string, number>()
  const texts = new Set<string>()
  options.forEach(option => counts.set(option.label.toLowerCase(), (counts.get(option.label.toLowerCase()) ?? 0) + 1))
  return options.map(option => {
    const text = counts.get(option.label.toLowerCase())! > 1 ? `${option.label} [${option.key}]` : option.label
    if (texts.has(text.toLowerCase())) {
      throw new SyntaxError('Choice display labels are ambiguous after adding keys')
    }
    texts.add(text.toLowerCase())
    return { ...option, text }
  })
}

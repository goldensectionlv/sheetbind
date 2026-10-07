import { isDataObject, jsonSnapshot } from './json'
import type { JsonValue } from './json'
import { isDataPath, readDataPath } from './reference'
import type { DataReference } from './template'
import { isDictionaryName } from './dictionaries'

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
export interface ResolvedChoice { readonly key: string | number | null, readonly items: readonly ChoiceOption[] }

export function parseChoiceRule(value: unknown): ChoiceRule {
  if (!isDataObject(value) || Object.keys(value).some(key => !['source', 'key', 'label', 'return', 'emptySource'].includes(key)) || !isDataPath(value.key) || !isDataPath(value.label) || !isDataObject(value.source)) {
    throw new SyntaxError('Choice requires a source and safe key/label paths')
  }
  if (value.return !== undefined && value.return !== 'object' && value.return !== 'key') {
    throw new SyntaxError('Choice return must be object or key')
  }
  if (value.emptySource !== undefined && (value.emptySource !== 'input' || value.return !== 'key')) {
    throw new SyntaxError('Choice emptySource=input requires return=key')
  }
  const source = value.source
  if ('dictionary' in source) {
    if (Object.keys(source).length !== 1 || !isDictionaryName(source.dictionary)) {
      throw new SyntaxError('Choice dictionary must be a named source')
    }
  }
  else if (Object.keys(source).some(key => !['path', 'from'].includes(key)) || !isDataPath(source.path) || source.from !== undefined && source.from !== 'root' && source.from !== 'current') {
    throw new SyntaxError('Choice source requires a safe path and current/root context')
  }
  return structuredClone(value) as unknown as ChoiceRule
}

/** Empty sources are valid; a nonblank selection must still match an available key. */
function buildOptions(source: readonly unknown[], rule: ChoiceRule): ChoiceOption[] {
  const keys = new Set<string>()
  const options = (jsonSnapshot(source) as readonly JsonValue[]).map(item => {
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
  return options
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

export function selectedChoice(rule: ChoiceRule, items: readonly ChoiceOption[], value: unknown): ChoiceOption | undefined {
  const key = choiceKey(rule, value)
  return items.find(item => item.key === key)
}

export function choiceKey(rule: ChoiceRule, value: unknown): unknown {
  return returnsObject(rule) ? readDataPath(value, rule.key) : value
}

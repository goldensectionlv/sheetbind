import { jsonSnapshot, isDataObject } from './json'

/** Application-supplied values; templates retain only named dependencies. */
export type Dictionaries = Readonly<Record<string, readonly string[] | readonly Readonly<Record<string, unknown>>[]>>

export function warnMissingDictionaries(names: Iterable<string>, dictionaries: Dictionaries): void {
  for (const name of new Set(names)) {
    if (!Object.hasOwn(dictionaries, name) || !dictionaries[name].length) {
      console.warn(`sheetbind: dictionary ${name} is missing or empty; its dropdown, lookup and list validation are skipped`)
    }
  }
}

export function isDictionaryName(value: unknown): value is string {
  return typeof value === 'string' && /^[A-Za-z_][A-Za-z0-9_.-]*$/.test(value)
    && !['__proto__', 'constructor', 'prototype'].includes(value)
}

/** Own only declared sources; an unavailable optional list must not stop rendering. */
export function parseDictionaries(value: unknown, names: Iterable<string>): Dictionaries {
  const dictionaries: Record<string, Dictionaries[string]> = {}
  for (const name of new Set(names)) {
    const property = isDataObject(value) ? Object.getOwnPropertyDescriptor(value, name) : undefined
    if (!property || 'value' in property && property.value == null) {
      warnMissingDictionaries([name], {})
      continue
    }
    try {
      if (!('value' in property)) {
        throw new SyntaxError('Dictionary sources cannot be accessors')
      }
      const source = jsonSnapshot(property.value)
      const items = validateDictionaryShape({ [name]: source })[name]
      const unique = items.every(item => typeof item === 'string') ? [...new Set(items)] : items
      if (unique.length !== items.length) {
        console.warn(`sheetbind: dictionary ${name} contains duplicate strings; repeated entries are ignored`)
      }
      dictionaries[name] = unique
      warnMissingDictionaries([name], dictionaries)
    }
    catch (error) {
      if (!(error instanceof SyntaxError)) {
        throw error
      }
      console.warn(`sheetbind: dictionary ${name} is unavailable: ${error.message}; its dropdown, lookup and list validation are skipped`)
    }
  }
  return dictionaries
}

/** The caller has already validated the JSON content; this check does not copy it. */
export function validateDictionaryShape(value: unknown): Dictionaries {
  if (!isDataObject(value)) {
    throw new SyntaxError('Dictionaries must be an object of named lists')
  }
  const entries = Object.entries(value)
  for (const [name, items] of entries) {
    if (!isDictionaryName(name) || !Array.isArray(items)) {
      throw new SyntaxError(`Invalid dictionary: ${name}`)
    }
    if (items.every(item => typeof item === 'string')) {
      if (items.some(item => !item.length)) {
        throw new SyntaxError(`Dictionary ${name} must contain nonempty strings`)
      }
      continue
    }
    if (!items.every(isDataObject)) {
      throw new SyntaxError(`Dictionary ${name} must contain only strings or only objects`)
    }
  }
  return value as Dictionaries
}

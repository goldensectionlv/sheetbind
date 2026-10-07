import { jsonSnapshot, isDataObject } from './json'

/** Application-supplied values; templates retain only named dependencies. */
export type Dictionaries = Readonly<Record<string, readonly string[] | readonly Readonly<Record<string, unknown>>[]>>

export function warnMissingDictionaries(names: Iterable<string>, dictionaries: Dictionaries): void {
  for (const name of new Set(names)) {
    if (!Object.hasOwn(dictionaries, name)) {
      console.warn(`sheetbind: dictionary ${name} was not supplied; its dropdown, lookup and list validation are skipped`)
    }
  }
}

export function isDictionaryName(value: unknown): value is string {
  return typeof value === 'string' && /^[A-Za-z_][A-Za-z0-9_.-]*$/.test(value)
    && !['__proto__', 'constructor', 'prototype'].includes(value)
}

/** Own application dictionaries before asynchronous workbook writing. */
export function parseDictionaries(value: unknown): Dictionaries {
  return validateDictionaryShape(jsonSnapshot(value))
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
      if (new Set(items).size !== items.length) {
        throw new SyntaxError(`Dictionary ${name} contains duplicate values`)
      }
      continue
    }
    if (!items.every(isDataObject)) {
      throw new SyntaxError(`Dictionary ${name} must contain only strings or only objects`)
    }
  }
  return value as Dictionaries
}

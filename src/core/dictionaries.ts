import { assertJson, isDataObject } from './json'

/** Application-supplied values; templates retain only named dependencies. */
export type Dictionaries = Readonly<Record<string, readonly string[] | readonly Readonly<Record<string, unknown>>[]>>
export function isDictionaryName(value: unknown): value is string {
  return typeof value === 'string' && /^[A-Za-z_][A-Za-z0-9_.-]*$/.test(value)
    && !['__proto__', 'constructor', 'prototype'].includes(value)
}

/** Object dictionaries are interpreted by each field's explicit key/label mapping. */
export function parseDictionaries(value: unknown): Dictionaries {
  assertJson(value)
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
  return structuredClone(value) as Dictionaries
}

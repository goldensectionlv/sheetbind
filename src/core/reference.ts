import { isDataObject } from './json'

export interface DataReference {
  readonly path: string
  readonly from?: 'current' | 'root'
}

export function isDataPath(value: unknown): value is string {
  return typeof value === 'string' && /^[\p{ID_Start}_]\p{ID_Continue}*(?:\.[\p{ID_Start}_]\p{ID_Continue}*)*$/u.test(value)
    && !value.split('.').some(part => ['__proto__', 'prototype', 'constructor'].includes(part))
}

/** Dotted references traverse own properties of data objects, never collection indexes. */
export function readDataPath(value: unknown, path: string): unknown {
  return compileDataPath(path)(value)
}
export function compileDataPath(path: string): (value: unknown) => unknown {
  const parts = path.split('.')
  return value => {
    for (const part of parts) {
      value = isDataObject(value) && Object.hasOwn(value, part) ? value[part] : undefined
    }
    return value
  }
}
/** Human-readable references; stored definitions keep path and origin separate. */
export function parseDataReference(text: string): DataReference {
  const from = text.startsWith('$root.') ? 'root' : 'current'
  const path = text.startsWith('$root.') ? text.slice(6) : text.startsWith('.') ? text.slice(1) : text
  if (!isDataPath(path)) {
    throw new SyntaxError('Use a dotted path such as customer.name, or .name for the current item')
  }
  return { path, from }
}

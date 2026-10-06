import { isDataObject } from '../core/json'

export type DataPath = readonly (string | number)[]

export function readData(data: unknown, path: DataPath): unknown {
  let value = data
  for (const part of path) {
    value = (isDataObject(value) || Array.isArray(value)) && Object.hasOwn(value, part) ? Reflect.get(value, part) : undefined
  }
  return value
}

/** Paths come from the template; constructed containers belong to the current operation. */
export function writeData(data: Record<string, unknown>, path: DataPath, value: unknown, create = false): void {
  let container: Record<string, unknown> | unknown[] = data
  for (const part of path.slice(0, -1)) {
    if (create && !Object.hasOwn(container, part) && !Reflect.set(container, part, {})) {
      throw new TypeError('Field container must be writable')
    }
    const child: unknown = Reflect.get(container, part)
    if (!isDataObject(child) && !Array.isArray(child)) {
      throw new Error('Field container must be an object')
    }
    container = child
  }
  if (!Reflect.set(container, path[path.length - 1], value)) {
    throw new TypeError('Field container must be writable')
  }
}

/** An issuance owns changed containers; unchanged records remain read-only and shared. */
export function createDataDraft(input: Record<string, unknown>) {
  const data = { ...input }
  const owned = new WeakSet<object>([data])
  function replace(path: DataPath, value: unknown): void {
    let container: Record<string, unknown> | unknown[] = data
    for (const part of path.slice(0, -1)) {
      const child: unknown = Reflect.get(container, part) ?? {}
      if (!isDataObject(child) && !Array.isArray(child)) {
        throw new Error('Field container must be an object')
      }
      const copy = owned.has(child) ? child : Array.isArray(child) ? [...child] : { ...child }
      owned.add(copy)
      Reflect.set(container, part, copy)
      container = copy
    }
    Reflect.set(container, path[path.length - 1], value)
  }
  return { data, replace }
}

export function dataPath(path: DataPath): string {
  return '$data' + path.map(part => typeof part === 'number' ? `[${part}]` : `.${part}`).join('')
}

export type JsonValue = null | boolean | number | string | readonly JsonValue[] | { readonly [key: string]: JsonValue }

export function isDataObject(value: unknown): value is Record<string, unknown> {
  if (value === null || typeof value !== 'object') {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

/** Validate execution data without copying it or evaluating accessors. */
export function assertInputData(value: unknown): asserts value is Record<string, unknown> {
  if (!isDataObject(value)) {
    throw new SyntaxError('Input data must be a plain object')
  }
  assertData(value, true)
}

/** Snapshot only data being retained; undefined object properties represent absence. */
export function jsonSnapshot(value: unknown): JsonValue {
  assertData(value, true)
  function copy(value: unknown): JsonValue {
    if (Array.isArray(value)) {
      return value.map(copy)
    }
    if (isDataObject(value)) {
      return Object.fromEntries(Object.entries(value).filter(([, item]) => item !== undefined).map(([key, item]) => [key, copy(item)]))
    }
    return value as JsonValue
  }
  return copy(value)
}

/** JSON object property order is not part of a value's identity. */
export function equalJson(left: unknown, right: unknown): boolean {
  if (left === right) {
    return true
  }
  if (Array.isArray(left) && Array.isArray(right)) {
    return left.length === right.length && left.every((value, index) => equalJson(value, right[index]))
  }
  if (!isDataObject(left) || !isDataObject(right)) {
    return false
  }
  const keys = Object.keys(left)
  return keys.length === Object.keys(right).length && keys.every(key => Object.hasOwn(right, key) && equalJson(left[key], right[key]))
}

/** Stored definitions and payloads contain strict JSON, unlike runtime input objects. */
export function assertJson(value: unknown): void {
  assertData(value, false)
}

/** Validate data without evaluating accessors or accepting undefined array items. */
function assertData(value: unknown, allowUndefinedProperties: boolean): void {
  const ancestors = new Set<object>()
  function visit(item: unknown): void {
    if (item === null || typeof item === 'string' || typeof item === 'boolean' || typeof item === 'number' && Number.isFinite(item)) {
      return
    }
    if (!Array.isArray(item) && !isDataObject(item) || ancestors.has(item as object)) {
      throw new SyntaxError('Only finite, non-circular JSON values can be saved')
    }
    const object = item as object
    const keys = Reflect.ownKeys(object).filter(key => !Array.isArray(object) || key !== 'length')
    if (Array.isArray(object) && (keys.length !== object.length || keys.some((key, index) => key !== String(index)))) {
      throw new SyntaxError('Sparse arrays and extra array properties cannot be saved as JSON')
    }
    ancestors.add(object)
    for (const key of keys) {
      const descriptor = Object.getOwnPropertyDescriptor(object, key)!
      if (typeof key !== 'string' || !descriptor.enumerable || !('value' in descriptor)) {
        throw new SyntaxError('JSON cannot contain symbols, accessors or hidden properties')
      }
      if (allowUndefinedProperties && !Array.isArray(object) && descriptor.value === undefined) {
        continue
      }
      visit(descriptor.value)
    }
    ancestors.delete(object)
  }
  visit(value)
}

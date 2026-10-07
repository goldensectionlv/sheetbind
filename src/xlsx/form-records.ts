import { isDataObject, assertInputData } from '../core/json'
import { isBlank } from '../core/validation'
import type { DataReference } from '../core/template'
import type { WorkbookBody, WorkbookDefinition, WorkbookRegion } from '../grid/workbook'

export type DataPath = readonly (string | number)[]

export function referencePath(reference: DataReference, context: DataPath): DataPath {
  return [...(reference.from === 'root' ? [] : context), ...reference.path.split('.')]
}

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
export interface WorkbookFormRows { readonly path: DataPath, readonly fields: readonly DataPath[] }

export function isWorkbookFormRow(region: WorkbookRegion): boolean {
  return region.height === 1 && !region.regions?.length
}

export function workbookFormRowFields(region: WorkbookRegion): DataPath[] {
  return region.cells.flatMap(cell => 'path' in cell.value && cell.value.from !== 'root' ? [cell.value.path.split('.')] : [])
}

/** Empty line collections need one styled input row; populated collections keep their actual length. */
export function issueWorkbookFormData(template: WorkbookDefinition, input: unknown) {
  assertInputData(input)
  const { data, replace } = createDataDraft(input)
  const blanks = new Map<string, { path: DataPath, value: Record<string, unknown> }>()
  const visit = (body: WorkbookBody, context: DataPath): void => {
    for (const region of body.regions ?? []) {
      const path = referencePath(region.source, context)
      const items = readData(data, path) ?? []
      if (!Array.isArray(items)) {
        continue
      }
      const fields = workbookFormRowFields(region)
      if (!items.length && isWorkbookFormRow(region) && fields.length) {
        const name = dataPath(path)
        const blank = blanks.get(name) ?? { path, value: {} }
        for (const field of fields) {
          writeData(blank.value, field, null, true)
        }
        blanks.set(name, blank)
      }
      items.forEach((_item, index) => visit(region, [...path, index]))
    }
  }
  template.sheets.forEach(sheet => visit(sheet, []))
  for (const blank of blanks.values()) {
    replace(blank.path, [blank.value])
  }
  return data
}

/** Compact privately owned decoded data; zero, false and partially filled records remain. */
export function readWorkbookRows(data: Record<string, unknown>, rows: readonly WorkbookFormRows[], invalid: readonly DataPath[] = []) {
  const entered = new Set(invalid.flatMap(path => path.flatMap((part, index) => typeof part === 'number' ? [dataPath(path.slice(0, index + 1))] : [])))
  const collections = new Map<string, { path: DataPath, fields: DataPath[] }>()
  for (const row of rows) {
    const name = dataPath(row.path)
    const collection = collections.get(name) ?? { path: row.path, fields: [] }
    collection.fields.push(...row.fields)
    collections.set(name, collection)
  }
  const indices = new Map<string, Map<number, number>>()
  for (const [name, { path, fields }] of collections) {
    const items = readData(data, path) as unknown[]
    const kept = new Map<number, number>()
    const values = items.filter((item, index) => {
      if (!entered.has(dataPath([...path, index])) && fields.length && fields.every(field => isBlank(readData(item, field)))) {
        return false
      }
      kept.set(index, kept.size)
      return true
    })
    indices.set(name, kept)
    writeData(data, path, values)
  }
  return { data, remapPath }

  function remapPath(source: DataPath): DataPath | undefined {
    const target: (string | number)[] = []
    for (const [index, part] of source.entries()) {
      const map = typeof part === 'number' ? indices.get(dataPath(source.slice(0, index))) : undefined
      if (map && !map.has(part as number)) {
        return undefined
      }
      target.push(map ? map.get(part as number)! : part)
    }
    return target
  }
}

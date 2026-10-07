import { assertInputData } from '../core/json'
import { isBlank } from '../core/validation'
import type { WorkbookBody, WorkbookDefinition, WorkbookRegion } from '../grid/workbook'
import { createDataDraft, dataPath, readData, writeData } from './records'
import type { DataPath } from './records'
import { referencePath } from './workbook'

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

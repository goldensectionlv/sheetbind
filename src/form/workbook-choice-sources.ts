import type { Dictionaries } from '../core/dictionaries'
import type { WorkbookDefinition } from '../grid/workbook'
import { workbookCells } from '../grid/workbook'
import { readData, writeData } from './records'

export interface WorkbookChoiceSources {
  readonly dictionaries: Dictionaries
  readonly context: Readonly<Record<string, unknown>>
}

/** Keep the issued dictionaries and root sources needed to read any return mode. */
export function workbookChoiceSources(template: WorkbookDefinition, data: unknown, dictionaries: Dictionaries): WorkbookChoiceSources | undefined {
  const context: Record<string, unknown> = {}
  const selected: Record<string, Dictionaries[string]> = {}
  const fields = template.sheets.flatMap(workbookCells).filter(cell => cell.rules?.choice || cell.rules?.list)
  if (!fields.length) {
    return undefined
  }
  for (const cell of fields) {
    if (cell.rules?.list) {
      selected[cell.rules.list] = dictionaries[cell.rules.list]
    }
    const source = cell.rules?.choice?.source
    if (!source) {
      continue
    }
    if ('dictionary' in source) {
      selected[source.dictionary] = dictionaries[source.dictionary]
    }
    else if (source.from === 'root') {
      const path = source.path.split('.')
      writeData(context, path, structuredClone(readData(data, path) ?? []), true)
    }
  }
  return { context, dictionaries: selected }
}

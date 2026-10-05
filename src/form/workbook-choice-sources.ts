import { returnsObject } from '../core/choices'
import type { Dictionaries } from '../core/dictionaries'
import type { WorkbookDefinition } from '../grid/workbook'
import { workbookCells } from '../grid/workbook'
import { readData, writeData } from './records'

export interface WorkbookChoiceSources {
  readonly dictionaries: Dictionaries
  readonly context: Readonly<Record<string, unknown>>
}

/** Store only the shared sources needed to recover selected objects from a form. */
export function workbookChoiceSources(template: WorkbookDefinition, data: unknown, dictionaries: Dictionaries): WorkbookChoiceSources | undefined {
  const context: Record<string, unknown> = {}
  const selected: Record<string, Dictionaries[string]> = {}
  const fields = template.sheets.flatMap(workbookCells).filter(cell => returnsObject(cell.rules?.choice))
  if (!fields.length) {
    return undefined
  }
  for (const cell of fields) {
    const source = cell.rules!.choice!.source
    if ('dictionary' in source) {
      selected[source.dictionary] = dictionaries[source.dictionary]
    }
    else {
      const path = source.path.split('.')
      writeData(context, path, structuredClone(readData(data, path)), true)
    }
  }
  return { context, dictionaries: selected }
}

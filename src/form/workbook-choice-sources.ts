import type { Dictionaries } from '../core/dictionaries'
import { workbookCells } from '../grid/workbook'
import { readData, writeData } from './records'
import type { WorkbookPlan } from '../grid/workbook-layout'

export interface WorkbookChoiceSources {
  readonly dictionaries: Dictionaries
  readonly context: Readonly<Record<string, unknown>>
  readonly local: Readonly<Record<string, Readonly<Record<string, readonly Readonly<Record<string, unknown>>[]>>>>
}

/** Keep the issued dictionaries and root sources needed to read any return mode. */
export function workbookChoiceSources(plan: WorkbookPlan, data: unknown, dictionaries: Dictionaries): WorkbookChoiceSources | undefined {
  const context: Record<string, unknown> = {}
  const selected: Record<string, Dictionaries[string]> = {}
  const fields = plan.sheets.flatMap(({ definition }) => workbookCells(definition)).filter(cell => cell.rules?.choice || cell.rules?.list)
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
      writeData(context, path, readData(data, path) ?? [], true)
    }
  }
  const local: Record<string, Record<string, readonly Readonly<Record<string, unknown>>[]>> = {}
  for (const { sheet } of plan.sheets) {
    for (const cell of sheet.cells) {
      const source = cell.rules?.choice?.source
      if (source && 'path' in source && source.from !== 'root') {
        const fields = local[cell.definitionId] ??= {}
        fields[cell.origin.dataPath] = cell.choice?.items.map(item => item.value) ?? []
      }
    }
  }
  return { context, dictionaries: selected, local }
}

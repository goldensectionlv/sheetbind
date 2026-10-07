import ExcelJS from 'exceljs'
import type { Dictionaries } from '../core/dictionaries'
import type { TemplateValue } from '../core/template'
import type { WorkbookCell, WorkbookSheet } from '../grid/workbook'
import { workbookCells } from '../grid/workbook'
import type { WorkbookPlan } from '../grid/workbook-layout'
import type { WorkbookFormula } from '../grid/workbook-formula'
import type { WorkbookChoice } from '../grid/workbook-choice-display'
import { formatAddress, formatRange } from './addresses'
import { writeWorkbookPrint } from './workbook-print'
import { writeWorkbookDropdowns } from './workbook-dropdowns'
import type { DropdownTarget } from './workbook-dropdowns'
import { writeWorkbookChoiceFields } from './workbook-choice-fields'
import { assertXlsxText } from './report-text'
import { createWorkbookResources } from './workbook-resources'
import type { WorkbookResources } from './workbook-resources'

/** Ready-to-write cells; bindings, repetition and carrier shifts are already resolved. */
export interface WorkbookOutputCell extends Pick<WorkbookCell, 'at' | 'size' | 'rules' | 'xlsx'> {
  readonly value: { readonly literal: TemplateValue } | WorkbookFormula
  readonly choice?: WorkbookChoice
  readonly text?: boolean
}
export interface WorkbookOutputSheet extends Omit<WorkbookSheet, 'cells' | 'regions'> {
  readonly cells: readonly WorkbookOutputCell[]
}

export function createWorkbookOutput(plan: WorkbookPlan, dictionaries: Dictionaries, resources: WorkbookResources = createWorkbookResources()) {
  const book = new ExcelJS.Workbook()
  book.calcProperties.fullCalcOnLoad = true
  const dropdowns: DropdownTarget[] = []
  const choiceTexts = new WeakMap<object, readonly string[]>()
  for (const { sheet: output } of plan.sheets) {
    const sheet = book.addWorksheet(output.name, { state: output.state ?? 'visible' })
    writeWorkbookPrint(sheet, output.print)
    for (const row of output.rows ?? []) {
      const target = sheet.getRow(row.index)
      if (row.height !== undefined) {
        target.height = row.height
      }
      target.hidden = !!row.hidden
      target.getCell(1)
    }
    for (const column of output.columns ?? []) {
      const target = sheet.getColumn(column.index)
      if (column.width !== undefined) {
        target.width = column.width
      }
      target.hidden = !!column.hidden
    }
    for (const definition of output.cells) {
      const at = definition.at
      const value = definition.value
      if ('literal' in value) {
        if (typeof value.literal === 'string') {
          assertXlsxText(value.literal)
        }
      }
      const merged = definition.size.rows > 1 || definition.size.columns > 1
      const address = merged || definition.choice || definition.rules?.list ? formatAddress(at) : ''
      if (definition.choice) {
        const choice = definition.choice
        if (choice.text !== null) {
          assertXlsxText(choice.text)
        }
        const items = choiceTexts.get(choice.items) ?? choice.items.map(item => item.text)
        choiceTexts.set(choice.items, items)
        dropdowns.push({ sheet, address, rules: definition.rules ?? {}, items })
      }
      if (merged) {
        sheet.mergeCells(formatRange({ start: at, end: { row: at.row + definition.size.rows - 1, column: at.column + definition.size.columns - 1 } }))
      }
      if (definition.rules?.list) {
        dropdowns.push({ sheet, address, rules: definition.rules, items: dictionaries[definition.rules.list] as readonly string[] })
      }
    }
  }
  const listSheet = writeWorkbookDropdowns(book, dropdowns, resources)
  const rules = plan.sheets.flatMap(({ definition }) => workbookCells(definition).flatMap(cell => cell.rules?.choice ? [cell.rules.choice] : []))
  writeWorkbookChoiceFields(book, rules, dictionaries, resources, listSheet)
  return book
}

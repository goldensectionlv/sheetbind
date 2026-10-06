import ExcelJS from 'exceljs'
import { returnsObject } from '../core/choices'
import type { Dictionaries } from '../core/dictionaries'
import type { TemplateValue } from '../core/template'
import type { WorkbookCell, WorkbookSheet } from '../grid/workbook'
import type { WorkbookFormula } from '../grid/workbook-formula'
import type { WorkbookChoice } from '../grid/workbook-choice-display'
import { formatAddress, formatRange } from './addresses'
import { writeWorkbookPrint } from './workbook-print'
import { workbookListSheetName, writeWorkbookDropdowns } from './workbook-dropdowns'
import type { DropdownTarget } from './workbook-dropdowns'
import { writeWorkbookChoiceFields } from './workbook-choice-fields'
import type { WorkbookChoiceTarget } from './workbook-choice-fields'
import { assertXlsxText } from './report-text'

/** Ready-to-write cells; bindings, repetition and carrier shifts are already resolved. */
export interface WorkbookOutputCell extends Pick<WorkbookCell, 'at' | 'size' | 'rules' | 'xlsx'> {
  readonly value: { readonly literal: TemplateValue } | WorkbookFormula
  readonly choice?: WorkbookChoice
  readonly text?: boolean
}
export interface WorkbookOutputSheet extends Omit<WorkbookSheet, 'cells' | 'regions'> {
  readonly cells: readonly WorkbookOutputCell[]
}

export function createWorkbookOutput(sheets: readonly WorkbookOutputSheet[], dictionaries?: Dictionaries, form = false) {
  const book = new ExcelJS.Workbook()
  book.calcProperties.fullCalcOnLoad = true
  const dropdowns: DropdownTarget[] = []
  const choices: WorkbookChoiceTarget[] = []
  const choiceSources = new WeakMap<WorkbookChoiceTarget['rule'], WeakSet<WorkbookChoiceTarget['items']>>()
  const choiceTexts = new WeakMap<object, readonly string[]>()
  for (const plan of sheets) {
    const sheet = book.addWorksheet(plan.name, { state: plan.state ?? 'visible' })
    writeWorkbookPrint(sheet, plan.print)
    for (const row of plan.rows ?? []) {
      const target = sheet.getRow(row.index)
      if (row.height !== undefined) {
        target.height = row.height
      }
      target.hidden = !!row.hidden
      target.getCell(1)
    }
    for (const column of plan.columns ?? []) {
      const target = sheet.getColumn(column.index)
      if (column.width !== undefined) {
        target.width = column.width
      }
      target.hidden = !!column.hidden
    }
    for (const definition of plan.cells) {
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
        if (returnsObject(definition.rules?.choice)) {
          const rule = definition.rules!.choice!
          const sources = choiceSources.get(rule) ?? new WeakSet<WorkbookChoiceTarget['items']>()
          if (!sources.has(choice.items)) {
            choices.push({ rule, items: choice.items })
            sources.add(choice.items)
            choiceSources.set(rule, sources)
          }
        }
        if (choice.text !== null) {
          assertXlsxText(choice.text)
        }
        const items = choiceTexts.get(choice.items) ?? choice.items.map(item => item.text)
        choiceTexts.set(choice.items, items)
        const source = definition.rules!.choice!.source
        const localFormChoice = form && 'path' in source && source.from !== 'root'
        // Form-local lists are written together with their copied source references.
        if (!localFormChoice && (items.length || definition.rules?.choice?.emptySource !== 'input')) {
          dropdowns.push({ sheet, address, rules: definition.rules ?? {}, items })
        }
      }
      if (merged) {
        sheet.mergeCells(formatRange({ start: at, end: { row: at.row + definition.size.rows - 1, column: at.column + definition.size.columns - 1 } }))
      }
      if (definition.rules?.list) {
        dropdowns.push({ sheet, address, rules: definition.rules })
      }
    }
  }
  const listSheet = writeWorkbookDropdowns(book, dropdowns, dictionaries)
    ?? (choices.length ? book.addWorksheet(workbookListSheetName(book.worksheets.map(sheet => sheet.name)), { state: 'veryHidden' }) : undefined)
  if (listSheet) {
    writeWorkbookChoiceFields(book, listSheet, choices)
  }
  return book
}

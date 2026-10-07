import type { Dictionaries } from '../core/dictionaries'
import { hasValidation } from '../core/field-rules'
import type { TemplateValue } from '../core/template'
import type { WorkbookChoiceSources } from '../form/workbook-choice-sources'
import { workbookCells } from '../grid/workbook'
import type { GridRange } from '../grid/geometry'
import type { WorkbookPlan } from '../grid/workbook-layout'
import { formatAddress, formatRange, XLSX_MAX_COLUMN, XLSX_MAX_ROW } from './addresses'
import { cellXml } from './workbook-cells'
import type { WorkbookCells } from './workbook-cells'
import { writeWorkbookChoiceFields } from './workbook-choice-fields'
import { writeWorkbookChoiceSources } from './workbook-choice-sources'
import type { WorkbookResources } from './workbook-resources'
import type { sourceStyles } from './source-styles'
import { protect } from './report-text'
import { encodeXml, setXmlAttributes } from './xml'

export type WriteWorkbookListColumn = (name: string, values: readonly TemplateValue[], header?: string) => void

export function workbookListSheetName(authoredNames: readonly string[]): string {
  const names = new Set(authoredNames.map(name => name.toLowerCase()))
  let name = '_sheetbind_lists'
  let suffix = 1
  while (names.has(name.toLowerCase())) {
    name = `_sheetbind_lists_${++suffix}`
  }
  return name
}

/** One hidden sheet contains named lists, referenced dictionary fields and issued form sources. */
export function workbookLists(plan: WorkbookPlan, dictionaries: Dictionaries, resources: WorkbookResources,
  cells: WorkbookCells, styles: Awaited<ReturnType<typeof sourceStyles>>, sources?: WorkbookChoiceSources) {
  const name = workbookListSheetName(plan.sheets.map(({ sheet }) => sheet.name))
  const rows: string[][] = []
  const names: string[] = []
  let column = 0
  const writeColumn: WriteWorkbookListColumn = (reference, values, header) => {
    const start = header === undefined ? 1 : 2
    const end = Math.max(start, start + values.length - 1)
    if (++column > XLSX_MAX_COLUMN || end > XLSX_MAX_ROW) {
      throw new RangeError('Workbook lists exceed the XLSX worksheet limits')
    }
    function write(row: number, value: TemplateValue) {
      const content = cells.content(value)
      const target = rows[row - 1] ??= []
      target.push(cellXml(formatAddress({ row, column }), { ...content, style: typeof value === 'string' ? styles.textStyle() : styles.style(0) }))
    }
    if (header !== undefined) {
      write(1, header)
    }
    values.forEach((value, index) => write(start + index, value))
    const absolute = (row: number) => formatAddress({ row, column }).replace(/^([A-Z]+)(\d+)$/, (_, col, row) => `$${col}$${row}`)
    const range = `'${name}'!${absolute(start)}:${absolute(end)}`
    names.push(`<definedName name="${encodeXml(reference)}">${encodeXml(protect(range))}</definedName>`)
  }
  const texts = new WeakMap<object, readonly string[]>()
  const references = new Map<string, string>()
  const byItems = new WeakMap<readonly string[], string>()
  const validations = new Map<string, string>()
  for (const { sheet } of plan.sheets) {
    const ranges: { range: GridRange, xml: string }[] = []
    const last = new Map<number, typeof ranges[number]>()
    for (const cell of sheet.cells) {
      let items: readonly string[]
      if (cell.choice) {
        items = texts.get(cell.choice.items) ?? cell.choice.items.map(item => item.text)
        texts.set(cell.choice.items, items)
      }
      else if (cell.rules?.list) {
        items = dictionaries[cell.rules.list] as readonly string[]
      }
      else {
        continue
      }
      let reference = byItems.get(items)
      if (!reference) {
        const key = JSON.stringify(items)
        reference = references.get(key)
        if (!reference) {
          reference = resources.allocate(`_sb_list_${references.size + 1}`)
          writeColumn(reference, items)
          references.set(key, reference)
        }
        byItems.set(items, reference)
      }
      const xml = `<dataValidation type="${items.length ? 'list' : 'custom'}" allowBlank="${hasValidation(cell.rules, 'required') ? 0 : 1}" showErrorMessage="1" errorStyle="stop" errorTitle="Choose a listed value" error="Select a value from the dropdown list."><formula1>${items.length ? reference : 'FALSE'}</formula1></dataValidation>`
      const previous = last.get(cell.at.column)
      if (previous && previous.range.end.row + 1 === cell.at.row && previous.xml === xml) {
        previous.range = { start: previous.range.start, end: cell.at }
      }
      else {
        const entry = { range: { start: cell.at, end: cell.at }, xml }
        ranges.push(entry)
        last.set(cell.at.column, entry)
      }
    }
    validations.set(sheet.name, ranges.map(({ range, xml }) => setXmlAttributes(xml, { sqref: formatRange(range) })).join(''))
  }
  const rules = plan.sheets.flatMap(({ definition }) => workbookCells(definition).flatMap(cell => cell.rules?.choice ? [cell.rules.choice] : []))
  writeWorkbookChoiceFields(rules, dictionaries, resources, writeColumn)
  if (sources) {
    writeWorkbookChoiceSources(sources, resources, writeColumn)
  }
  const xml = column
    ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${formatAddress({ row: Math.max(1, rows.length), column })}"/><sheetData>${rows.map((cells, row) => `<row r="${row + 1}">${cells.join('')}</row>`).join('')}</sheetData></worksheet>`
    : undefined
  return { name, xml, names, validations }
}

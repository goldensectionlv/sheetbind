import type JSZip from 'jszip'
import { columnName } from '../grid/geometry'
import { formatRange, parseAddress, parseRange } from './addresses'
import type { WorkbookOutputCell, WorkbookOutputSheet } from './workbook-output'
import { protect } from './report-text'
import { appendXmlChildren, encodeXml, setXmlAttributes, xmlAttributes, xmlElements } from './xml'

export type WorkbookCells = Awaited<ReturnType<typeof prepareWorkbookCells>>

/** Serialize resolved values into the final workbook string table. */
export async function prepareWorkbookCells(zip: JSZip, inheritedStrings: readonly string[] = []) {
  const prior = await zip.file('xl/sharedStrings.xml')?.async('string')
  const strings = [...inheritedStrings, ...xmlElements(prior ?? '', 'si')]
  const values = new Map<string, number>()
  function text(value: string): number {
    let index = values.get(value)
    if (index !== undefined) {
      return index
    }
    index = strings.length
    values.set(value, index)
    strings.push(`<si><t xml:space="preserve">${encodeXml(protect(value))}</t></si>`)
    return index
  }
  function content(cell: WorkbookOutputCell): { style: number, type?: string, body: string } {
    const style = 0
    if ('formula' in cell.value) {
      return { style, body: `<f>${encodeXml(cell.value.formula)}</f>` }
    }
    const value = cell.choice ? cell.choice.text : cell.value.literal
    if (typeof value === 'string') {
      return { style, type: 's', body: `<v>${text(value)}</v>` }
    }
    if (typeof value === 'boolean') {
      return { style, type: 'b', body: `<v>${Number(value)}</v>` }
    }
    return { style, body: value === null ? '' : `<v>${value}</v>` }
  }
  return { content, strings: strings as readonly string[], stringOffset: inheritedStrings.length }
}

/** Cells share one string table for the final package; no intermediate ZIP round-trip. */
export async function writeWorkbookStrings(zip: JSZip, strings: readonly string[]): Promise<void> {
  if (!strings.length) {
    return
  }
  const present = zip.file('xl/sharedStrings.xml') !== null
  zip.file('xl/sharedStrings.xml', new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="${strings.length}">${strings.join('')}</sst>`))
  if (present) {
    return
  }
  const relations = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file('xl/_rels/workbook.xml.rels', appendXmlChildren(relations, 'Relationships', ['<Relationship Id="sheetbindStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>']))
  zip.file('[Content_Types].xml', appendXmlChildren(types, 'Types', ['<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>']))
}

export function cellXml(address: string, content: ReturnType<WorkbookCells['content']>): string {
  return `<c r="${address}"${content.style ? ` s="${content.style}"` : ''}${content.type ? ` t="${content.type}"` : ''}>${content.body}</c>`
}

/** Emit each row once, retaining carrier cells, merged edges and unowned source cells. */
export function worksheetCells(xml: string, sheet: WorkbookOutputSheet, render: (cell: WorkbookOutputCell, address: string) => string,
  source?: { rows: ReadonlyMap<number, string>, cells: ReadonlyMap<number, ReadonlyMap<number, string>> }) {
  const generated = new Map(xmlElements(xmlElements(xml, 'sheetData')[0] ?? '', 'row').map(row => [Number(xmlAttributes(row.slice(0, row.indexOf('>') + 1)).r), row]))
  const definitions = new Map<number, WorkbookOutputCell[]>()
  const columns = new Map<number, string>()
  const dimension = xmlElements(xml, 'dimension')[0]
  const initial = dimension ? parseRange(xmlAttributes(dimension).ref) : { start: { row: 1, column: 1 }, end: { row: 1, column: 1 } }
  const bounds = { start: { ...initial.start }, end: { ...initial.end } }
  let currentRow = 0
  let currentCells: WorkbookOutputCell[] = []
  for (const cell of sheet.cells) {
    if (currentRow !== cell.at.row) {
      currentRow = cell.at.row
      currentCells = definitions.get(currentRow) ?? []
      definitions.set(currentRow, currentCells)
    }
    currentCells.push(cell)
    bounds.start.row = Math.min(bounds.start.row, cell.at.row)
    bounds.start.column = Math.min(bounds.start.column, cell.at.column)
    bounds.end.row = Math.max(bounds.end.row, cell.at.row + cell.size.rows - 1)
    bounds.end.column = Math.max(bounds.end.column, cell.at.column + cell.size.columns - 1)
  }
  const rows = [...new Set([...generated.keys(), ...definitions.keys(), ...source?.rows.keys() ?? [], ...source?.cells.keys() ?? []])].sort((a, b) => a - b).map(index => {
    const row = generated.get(index) ?? ''
    const original = source?.rows.get(index)
    let head = original ?? (row ? row.slice(0, row.indexOf('>') + 1) : `<row r="${index}"/>`)
    if (row && original) {
      const attributes = xmlAttributes(row.slice(0, row.indexOf('>') + 1))
      head = setXmlAttributes(head, { ht: attributes.ht, customHeight: attributes.customHeight, hidden: attributes.hidden })
    }
    head = setXmlAttributes(head, { spans: undefined }).replace(/\/>$/, '>')
    const cells = new Map(xmlElements(row, 'c').map(cell => [parseAddress(xmlAttributes(cell.slice(0, cell.indexOf('>') + 1)).r).column, cell]))
    for (const [column, cell] of source?.cells.get(index) ?? []) {
      cells.set(column, cell)
    }
    const extra = [...cells].sort(([a], [b]) => a - b)
    const values: string[] = []
    let cursor = 0
    for (const cell of (definitions.get(index) ?? []).sort((a, b) => a.at.column - b.at.column)) {
      while (cursor < extra.length && extra[cursor][0] < cell.at.column) {
        values.push(extra[cursor++][1])
      }
      if (extra[cursor]?.[0] === cell.at.column) {
        cursor++
      }
      let column = columns.get(cell.at.column)
      if (!column) {
        column = columnName(cell.at.column)
        columns.set(cell.at.column, column)
      }
      values.push(render(cell, column + index))
    }
    while (cursor < extra.length) {
      values.push(extra[cursor++][1])
    }
    return head + values.join('') + '</row>'
  })
  return { data: `<sheetData>${rows.join('')}</sheetData>`, dimension: `<dimension ref="${formatRange(bounds)}"/>` }
}

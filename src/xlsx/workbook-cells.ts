import type JSZip from 'jszip'
import { columnName } from '../grid/geometry'
import type { TemplateValue } from '../core/template'
import type { WorkbookFormula } from '../grid/workbook-formula'
import { formatRange } from '../grid/geometry'
import type { WorkbookPlacedCell, WorkbookPlacedSheet } from '../grid/workbook-layout'
import { assertXlsxText, protect } from './report-text'
import { workbookRelationshipId } from './workbook-resources'
import { appendXmlChildren, encodeXml, setXmlAttributes, setXmlElement, xmlAttributes, xmlElements } from './xml'

export type WorkbookCells = Awaited<ReturnType<typeof prepareWorkbookCells>>

const defaultStyles = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
  + '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
  + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
  + '<borders count="1"><border/></borders>'
  + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
  + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
  + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

/** One writer owns shared strings and derived styles for all output cells. */
export async function prepareWorkbookCells(zip: JSZip) {
  const source = await zip.file('xl/styles.xml')?.async('string')
  let xml = source ?? defaultStyles
  if (!xmlElements(xmlElements(xml, 'cellXfs')[0] ?? '', 'xf').length) {
    // Excel needs base resources when the implicit default style becomes explicit.
    xml = xml.replace(/(<styleSheet\b[^>]*?)\/>/, '$1></styleSheet>')
    let next = /<(?:dxfs|tableStyles|colors|extLst)\b[^>]*>|<\/styleSheet>/.exec(xml)![0]
    for (const [name, child] of [['cellStyles', 'cellStyle'], ['cellXfs', 'xf'], ['cellStyleXfs', 'xf'], ['borders', 'border'], ['fills', 'fill'], ['fonts', 'font']]) {
      let table = xmlElements(xml, name)[0]
      if (!table || !xmlElements(table, child).length) {
        const fallback = xmlElements(defaultStyles, name)[0]
        xml = table ? xml.replace(table, () => fallback) : xml.replace(next, () => fallback + next)
        table = fallback
      }
      next = table
    }
  }
  const originals = xmlElements(xmlElements(xml, 'cellXfs')[0] ?? '', 'xf')
  const all = [...originals]
  const textStyles = new Map<number, number>()
  let normal: number | undefined
  function textStyle(original = 0): number {
    const cached = textStyles.get(original)
    if (cached !== undefined) {
      return cached
    }
    const id = all.length
    all.push(setXmlAttributes(originals[original] ?? originals[0], { numFmtId: 49, applyNumberFormat: 1 }))
    textStyles.set(original, id)
    return id
  }
  function style(original: number, text = false): number {
    if (text && Number(xmlAttributes(originals[original] ?? originals[0]).numFmtId ?? 0) === 0) {
      return textStyle(original)
    }
    // A nonzero ID prevents a moved General cell from inheriting its new row/column format.
    if (!original && normal === undefined) {
      normal = all.length
      all.push(originals[0])
    }
    return original || normal!
  }
  const prior = await zip.file('xl/sharedStrings.xml')?.async('string')
  const strings = xmlElements(prior ?? '', 'si')
  const values = new Map<string, number>()
  function text(value: string): number {
    let index = values.get(value)
    if (index !== undefined) {
      return index
    }
    assertXlsxText(value)
    index = strings.length
    values.set(value, index)
    strings.push(`<si><t xml:space="preserve">${encodeXml(protect(value))}</t></si>`)
    return index
  }
  function content(value: TemplateValue | WorkbookFormula): { style: number, type?: string, body: string } {
    const style = 0
    if (value !== null && typeof value === 'object') {
      return { style, body: `<f>${encodeXml(value.formula)}</f>` }
    }
    if (typeof value === 'string') {
      return { style, type: 's', body: `<v>${text(value)}</v>` }
    }
    if (typeof value === 'boolean') {
      return { style, type: 'b', body: `<v>${Number(value)}</v>` }
    }
    return { style, body: value === null ? '' : `<v>${value}</v>` }
  }
  return { content, style, textStyle,
    async save() {
      if (all.length !== originals.length || xml !== source) {
        await writeCellPart(zip, 'styles', setXmlElement(xml, 'cellXfs', `<cellXfs count="${all.length}">${all.join('')}</cellXfs>`))
      }
      if (strings.length) {
        await writeCellPart(zip, 'sharedStrings', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="${strings.length}">${strings.join('')}</sst>`)
      }
    } }
}

/** Style and string tables are optional in the source package. */
async function writeCellPart(zip: JSZip, part: 'styles' | 'sharedStrings', xml: string): Promise<void> {
  const path = `xl/${part}.xml`
  const present = zip.file(path) !== null
  zip.file(path, new TextEncoder().encode(xml))
  if (present) {
    return
  }
  const relations = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
  const types = await zip.file('[Content_Types].xml')!.async('string')
  const id = workbookRelationshipId(relations, part === 'styles' ? 'sheetbindStyles' : 'sheetbindStrings')
  zip.file('xl/_rels/workbook.xml.rels', appendXmlChildren(relations, 'Relationships', [`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${part}" Target="${part}.xml"/>`]))
  zip.file('[Content_Types].xml', appendXmlChildren(types, 'Types', [`<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.${part}+xml"/>`]))
}

export function cellXml(address: string, content: ReturnType<WorkbookCells['content']>): string {
  return `<c r="${address}"${content.style ? ` s="${content.style}"` : ''}${content.type ? ` t="${content.type}"` : ''}>${content.body}</c>`
}

/** Emit each row once, retaining carrier cells, merged edges and unowned source cells. */
export function worksheetCells(sheet: WorkbookPlacedSheet, render: (cell: WorkbookPlacedCell, address: string) => string,
  source: { rows: ReadonlyMap<number, string>, cells: ReadonlyMap<number, ReadonlyMap<number, string>> }) {
  const settings = new Map((sheet.rows ?? []).map(row => [row.index, row]))
  const definitions = new Map<number, WorkbookPlacedCell[]>()
  const columns = new Map<number, string>()
  const bounds = { start: { row: 1, column: 1 }, end: { row: 1, column: 1 } }
  let currentRow = 0
  let currentCells: WorkbookPlacedCell[] = []
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
  const rows = [...new Set([...settings.keys(), ...definitions.keys(), ...source.rows.keys(), ...source.cells.keys()])].sort((a, b) => a - b).map(index => {
    let head = source.rows.get(index) ?? `<row r="${index}"/>`
    const setting = settings.get(index)
    if (setting) {
      head = setXmlAttributes(head, { ht: setting.height, customHeight: setting.height === undefined ? undefined : 1, hidden: setting.hidden ? 1 : undefined })
    }
    head = setXmlAttributes(head, { spans: undefined }).replace(/\/>$/, '>')
    const extra = [...source.cells.get(index) ?? []].sort(([a], [b]) => a - b)
    bounds.end.row = Math.max(bounds.end.row, index)
    bounds.end.column = Math.max(bounds.end.column, extra.at(-1)?.[0] ?? 1)
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

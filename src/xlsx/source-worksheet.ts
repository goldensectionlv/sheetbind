import { FormulaEdge } from '../grid/workbook-formula'
import type { GridRange } from '../grid/geometry'
import type { WorkbookPlacedCell, WorkbookPlacedSheet } from '../grid/workbook-layout'
import { hasValidation } from '../core/field-rules'
import { cellXml, worksheetCells } from './workbook-cells'
import type { WorkbookCells } from './workbook-cells'
import type { SourceCoordinates } from './source-coordinates'
import { sourceFormula } from './source-coordinates'
import { relocateFilterRanges } from './source-metadata'
import type { sourceStyles } from './source-styles'
import { formatAddress, formatRange, parseAddress, parseRange } from '../grid/geometry'
import { FORM_MARKER_PREFIX } from './workbook-form-markers'
import type { FormMarkers } from './workbook-form-markers'
import { decodeXml, encodeXml, setXmlAttributes, setXmlElement, xmlAttributes, xmlBody, xmlElements } from './xml'

/** Keep worksheet features in place and replace only the rows/cells the layout owns. */
export function sourceWorksheet(xml: string, options: {
  readonly output: WorkbookPlacedSheet
  readonly map: SourceCoordinates
  readonly maps: ReadonlyMap<string, SourceCoordinates>
  readonly styles: Awaited<ReturnType<typeof sourceStyles>>
  readonly writer: WorkbookCells
  readonly validations: string
  readonly markers?: FormMarkers
  readonly form?: boolean
}) {
  const { output, map, maps, styles, writer } = options
  const sourceRows = new Map<number, string>()
  const extra = new Map<number, Map<number, string>>()
  const originals = new Map<string, ReturnType<typeof sourceCell>>()
  const owned = new Set<string>()
  for (const cell of output.cells) {
    if (cell.xlsx && cell.xlsx.part === map.part) {
      owned.add(cell.xlsx.address)
    }
  }
  for (const row of xmlElements(xmlElements(xml, 'sheetData')[0] ?? '', 'row')) {
    const index = Number(xmlAttributes(row.split('>')[0]).r)
    for (const target of map.rows(index)) {
      sourceRows.set(target, setXmlAttributes(row.split('>')[0].replace(/\/$/, '') + '/>', { r: target, spans: undefined }))
    }
    for (const cell of xmlElements(row, 'c')) {
      const original = sourceCell(cell, map, maps, styles, options.form)
      const address = original.attributes.r
      originals.set(address, original)
      if (owned.has(address)) {
        continue
      }
      for (const target of map.points(address)) {
        const at = parseAddress(target)
        const cells = extra.get(at.row) ?? new Map<number, string>()
        cells.set(at.column, setXmlAttributes(original.xml, { r: target }))
        extra.set(at.row, cells)
      }
    }
  }
  if (options.markers) {
    const column = options.markers.column
    for (const { row, token } of options.markers.rows) {
      const cells = extra.get(row) ?? new Map<number, string>()
      cells.set(column, cellXml(formatAddress({ row, column }), writer.content(FORM_MARKER_PREFIX + JSON.stringify(token))))
      extra.set(row, cells)
    }
  }
  const rows = worksheetCells(output, (definition, address) => {
    const content = writer.content('formula' in definition.value ? definition.value : definition.choice ? definition.choice.text : definition.value.literal)
    const original = definition.xlsx && definition.xlsx.part === map.part ? originals.get(definition.xlsx.address) : undefined
    return original ? original.render(address, content.type, definition, content.body) : cellXml(address, { ...content, style: styles.style(0) })
  }, { rows: sourceRows, cells: extra })
  const merges = output.cells.filter(cell => cell.size.rows > 1 || cell.size.columns > 1).map(cell =>
    `<mergeCell ref="${formatRange({ start: cell.at, end: { row: cell.at.row + cell.size.rows - 1, column: cell.at.column + cell.size.columns - 1 } })}"/>`)
  xml = mapWorksheetReferences(xml, map, maps)
  xml = setXmlElement(xml, 'sheetData', '<sheetData/>')
  xml = setXmlElement(xml, 'cols', mergeColumns(xml, output, map))
  xml = setXmlElement(xml, 'mergeCells', merges.length ? `<mergeCells count="${merges.length}">${merges.join('')}</mergeCells>` : '')
  xml = setXmlElement(xml, 'dimension', rows.dimension)
  xml = setXmlElement(xml, 'dataValidations', mergeValidations(xml, options.validations, map.name))
  return orderWorksheet(xml).replace('<sheetData/>', () => rows.data)
}

function sourceCell(xml: string, map: SourceCoordinates, maps: ReadonlyMap<string, SourceCoordinates>, styles: Awaited<ReturnType<typeof sourceStyles>>, form = false) {
  xml = mapCellFormula(xml, map, maps)
  const head = xml.slice(0, xml.indexOf('>') + 1)
  const attributes = xmlAttributes(head)
  const body = xmlBody(xml)
  const variants = new Map<string, string>()
  function render(address: string, type: string | undefined, definition: WorkbookPlacedCell, content: string): string {
    const before = definition.xlsx!.value
    const unchanged = 'literal' in definition.value && 'literal' in before && before.literal === definition.value.literal
    if (unchanged) {
      type = attributes.t
      content = body
    }
    const text = form && 'path' in before && (hasValidation(definition.rules, 'string') || !!definition.choice)
    const key = `${text}:${type ?? ''}`
    let prefix = variants.get(key)
    if (!prefix) {
      const merged = styles.style(Number(attributes.s ?? 0), text)
      prefix = setXmlAttributes(head, { r: undefined, s: merged, t: type }).replace(/\/?>$/, '')
      variants.set(key, prefix)
    }
    return `${prefix} r="${address}">${content}</c>`
  }
  return { xml, attributes, render }
}

function mapCellFormula(cell: string, map: SourceCoordinates, maps: ReadonlyMap<string, SourceCoordinates>): string {
  return cell.replace(/<f\b([^>]*)>([\s\S]*?)<\/f>/g, (_, attributes: string, value: string) => {
    const ref = xmlAttributes(attributes).ref
    const head = ref ? setXmlAttributes(`<f${attributes}>`, { ref: map.range(ref) ?? '' }) : `<f${attributes}>`
    return head + encodeXml(sourceFormula(decodeXml(value), map, maps)) + '</f>'
  })
}

/** A view anchor follows the first copy, or the grid position left by a removed cell. */
function viewCell(address: string, map: SourceCoordinates): string {
  const cell = parseAddress(address)
  return map.point(address) ?? formatAddress({
    row: map.row(cell.row, FormulaEdge.Start)!,
    column: map.column(cell.column, FormulaEdge.Start)!,
  })
}

function mapSelection(tag: string, map: SourceCoordinates): string {
  const attr = xmlAttributes(tag)
  if (!attr.sqref && !attr.activeCell) {
    return tag
  }
  const ranges = map.references(attr.sqref ?? attr.activeCell)
  const active = attr.activeCell ? map.point(attr.activeCell) : undefined
  const cell = active ? parseAddress(active) : undefined
  const index = cell
    ? ranges.findIndex(reference => {
      const range = parseRange(reference)
      return cell.row >= range.start.row && cell.row <= range.end.row
        && cell.column >= range.start.column && cell.column <= range.end.column
    })
    : -1
  const activeCell = index >= 0 ? active! : ranges[0]?.split(':')[0] ?? viewCell(attr.activeCell ?? 'A1', map)
  return setXmlAttributes(tag, {
    activeCell,
    activeCellId: index > 0 ? index : undefined,
    sqref: ranges.join(' ') || activeCell,
  })
}

function mapWorksheetReferences(xml: string, map: SourceCoordinates, maps: ReadonlyMap<string, SourceCoordinates>): string {
  const data = xmlElements(xml, 'sheetData')[0] ?? ''
  xml = xml.replace(data, '<sheetData/>')
  xml = relocateFilterRanges(xml, map)
  xml = xml.replace(/<hyperlink\b[^>]*\/?>(?:[\s\S]*?<\/hyperlink>)?/g, element => {
    const ref = xmlAttributes(element.split('>')[0]).ref
    return ref ? map.references(ref).map(ref => setXmlAttributes(element, { ref })).join('') : element
  })
  xml = xml.replace(/<[\w:]+\b[^>]*>/g, tag => {
    if (/^<selection\b/.test(tag)) {
      return mapSelection(tag, map)
    }
    const attr = xmlAttributes(tag)
    const values: Record<string, string | number> = {}
    if (attr.sqref) {
      values.sqref = map.references(attr.sqref).join(' ')
    }
    if (attr.topLeftCell) {
      values.topLeftCell = viewCell(attr.topLeftCell, map)
    }
    if (/^<pane\b/.test(tag)) {
      if (attr.state?.startsWith('frozen') && attr.ySplit) {
        values.ySplit = map.row(Number(attr.ySplit) + 1, FormulaEdge.Start)! - 1
      }
      if (attr.state?.startsWith('frozen') && attr.xSplit) {
        values.xSplit = map.column(Number(attr.xSplit) + 1, FormulaEdge.Start)! - 1
      }
    }
    return setXmlAttributes(tag, values)
  })
  xml = xml.replace(/<(formula1|formula2|formula)\b([^>]*)>([\s\S]*?)<\/\1>/g, (_, tag: string, attributes: string, value: string) => `<${tag}${attributes}>${encodeXml(sourceFormula(decodeXml(value), map, maps))}</${tag}>`)
  xml = xml.replace(/<conditionalFormatting\b[^>]*>[\s\S]*?<\/conditionalFormatting>/g, node => xmlAttributes(node.split('>')[0]).sqref ? node : '')
  xml = xml.replace(/<hyperlinks\b[^>]*>\s*<\/hyperlinks>/g, '')
  return xml.replace('<sheetData/>', data)
}
function mergeColumns(xml: string, output: WorkbookPlacedSheet, map: SourceCoordinates): string {
  const columns = new Map<number, string>()
  for (const node of xmlElements(xmlElements(xml, 'cols')[0] ?? '', 'col')) {
    const attr = xmlAttributes(node)
    for (let index = Number(attr.min); index <= Number(attr.max); index++) {
      for (const target of map.columns(index)) {
        columns.set(target, setXmlAttributes(node, { min: target, max: target }))
      }
    }
  }
  for (const { index, width, hidden } of output.columns ?? []) {
    columns.set(index, setXmlAttributes(columns.get(index) ?? '<col/>', { min: index, max: index, width, hidden: hidden ? 1 : undefined, customWidth: width === undefined ? undefined : 1 }))
  }
  return columns.size ? `<cols>${[...columns].sort(([a], [b]) => a - b).map(([, node]) => node).join('')}</cols>` : ''
}
function validationOverlap(a: GridRange, b: GridRange): GridRange | undefined {
  const start = { row: Math.max(a.start.row, b.start.row), column: Math.max(a.start.column, b.start.column) }
  const end = { row: Math.min(a.end.row, b.end.row), column: Math.min(a.end.column, b.end.column) }
  return start.row <= end.row && start.column <= end.column ? { start, end } : undefined
}
function withoutValidation(range: GridRange, cut: GridRange): GridRange[] {
  const overlap = validationOverlap(range, cut)
  if (!overlap) {
    return [range]
  }
  const { start, end } = overlap
  return [
    { start: range.start, end: { row: start.row - 1, column: range.end.column } },
    { start: { row: end.row + 1, column: range.start.column }, end: range.end },
    { start: { row: start.row, column: range.start.column }, end: { row: end.row, column: start.column - 1 } },
    { start: { row: start.row, column: end.column + 1 }, end: { row: end.row, column: range.end.column } },
  ].filter(part => part.start.row <= part.end.row && part.start.column <= part.end.column)
}

/** Merge generated dropdowns with native rules at their final coordinates. */
function mergeValidations(xml: string, fresh: string, sheetName: string): string {
  const read = (xml: string) => xmlElements(xml, 'dataValidation').map(node => {
    const attributes = xmlAttributes(node.split('>')[0])
    return { node, attributes, ranges: (attributes.sqref ?? '').split(/\s+/).filter(Boolean).map(parseRange) }
  })
  const source = read(xml)
  const generated = read(fresh)
  const claimed = generated.flatMap(entry => entry.ranges)
  const nodes: string[] = []
  for (const entry of source) {
    let remaining = entry.ranges
    for (const cut of claimed) {
      for (const range of remaining) {
        const overlap = validationOverlap(range, cut)
        if (overlap && entry.attributes.type && !['none', 'any'].includes(entry.attributes.type)) {
          throw new RangeError(`Existing validation at ${sheetName}!${formatRange(overlap)}`)
        }
      }
      remaining = remaining.flatMap(range => withoutValidation(range, cut))
    }
    if (remaining.length) {
      nodes.push(setXmlAttributes(entry.node, { sqref: remaining.map(formatRange).join(' ') }))
    }
  }
  for (const entry of generated) {
    let remaining = entry.ranges
    for (const original of source) {
      const inherited: GridRange[] = []
      for (const range of original.ranges) {
        for (const target of remaining) {
          const overlap = validationOverlap(range, target)
          if (overlap) {
            inherited.push(overlap)
          }
        }
        remaining = remaining.flatMap(target => withoutValidation(target, range))
      }
      if (inherited.length) {
        nodes.push(setXmlAttributes(entry.node, { ...original.attributes, type: entry.attributes.type, operator: undefined,
          allowBlank: entry.attributes.allowBlank, sqref: inherited.map(formatRange).join(' ') }))
      }
    }
    if (remaining.length) {
      nodes.push(setXmlAttributes(entry.node, { sqref: remaining.map(formatRange).join(' ') }))
    }
  }
  const container = xmlElements(xml, 'dataValidations')[0] ?? '<dataValidations>'
  const head = setXmlAttributes(container.split('>')[0].replace(/\/$/, '') + '>', { count: nodes.length })
  return nodes.length ? `${head}${nodes.join('')}</dataValidations>` : ''
}
function orderWorksheet(xml: string): string {
  const order = 'sheetPr dimension sheetViews sheetFormatPr cols sheetData sheetCalcPr sheetProtection protectedRanges scenarios autoFilter sortState dataConsolidate customSheetViews mergeCells phoneticPr conditionalFormatting dataValidations hyperlinks printOptions pageMargins pageSetup headerFooter rowBreaks colBreaks customProperties cellWatches ignoredErrors smartTags drawing legacyDrawing legacyDrawingHF drawingHF picture oleObjects controls webPublishItems tableParts extLst'.split(' ')
  const parts: string[] = []
  for (const tag of order) {
    for (const part of xmlElements(xml, tag)) {
      parts.push(part)
      xml = xml.replace(part, '')
    }
  }
  return xml.replace(/<\/worksheet>\s*$/, () => parts.join('') + '</worksheet>')
}

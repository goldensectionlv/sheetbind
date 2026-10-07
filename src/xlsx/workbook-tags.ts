import ExcelJS from 'exceljs'
import type { Worksheet } from 'exceljs'
import { WorkbookAxis } from '../grid/workbook'
import type { WorkbookRegion } from '../grid/workbook'
import { FormulaEdge } from '../grid/workbook-formula'
import type { ValueExpression } from '../core/template'
import type { FieldRules } from '../core/field-rules'
import { parseDataReference } from '../core/reference'
import { splitRuleText } from '../core/rule-syntax'
import { parseFieldTag } from './field-tag'
import { parseRange } from './addresses'
import { xlsxTextIssues } from './report-text'
import { TaggedXlsxError } from './tagged-template'

interface RegionNode extends Pick<WorkbookRegion, 'type' | 'source' | 'axis'> {
  id: string
  address: string
  row: number
  endRow: number
  column: number
  width: number
}
interface TaggedRegion extends RegionNode { readonly parentId?: string }
interface TaggedField { readonly expression: ValueExpression, readonly rules?: FieldRules }

function tagText(cell: ExcelJS.Cell): string | undefined {
  if (typeof cell.value === 'string') {
    return cell.value
  }
  if (cell.type === ExcelJS.ValueType.RichText || cell.type === ExcelJS.ValueType.Hyperlink) {
    return cell.text
  }
}

function parseShape(text: string) {
  const parts = splitRuleText(text, '|')
  const shape: { axis?: `${WorkbookAxis}` } = {}
  const binding = parts.filter((part, index) => {
    if (!index) {
      return true
    }
    const [key, value, extra] = part.split('=').map(value => value.trim())
    if (key !== 'axis') {
      return true
    }
    if (extra !== undefined || Object.hasOwn(shape, key)) {
      throw new SyntaxError('Duplicate or malformed region shape')
    }
    if (value !== WorkbookAxis.Rows && value !== WorkbookAxis.Columns) {
      throw new SyntaxError('Region axis must be rows or columns')
    }
    shape.axis = value
    return false
  })
  if (binding.length !== 1) {
    throw new SyntaxError('Unknown region option; use the closing marker to set the rectangle')
  }
  const head = binding[0]
  const scope = head.startsWith('with ')
  const source = parseDataReference(scope ? head.slice(5) : head)
  return { source, scope, close: scope ? 'with' : head, ...shape }
}

/** Rectangular ownership is reconstructed before parsing fields. Marker rows are a carrier only. */
export function compileWorkbookSheet(sheet: Worksheet) {
  const regions: RegionNode[] = []
  const fields = new Map<string, TaggedField>()
  const markers = new Set<number>()
  const pending: { close: string, region: RegionNode }[] = []
  function fail(address: string, code: string, message: string): never {
    throw new TaggedXlsxError([{ phase: 'template', code, message, sheetName: sheet.name, address, path: address }])
  }
  sheet.eachRow(row => {
    const cells: ExcelJS.Cell[] = []
    row.eachCell(cell => {
      if (cell.type !== ExcelJS.ValueType.Merge) {
        cells.push(cell)
      }
    })
    for (const cell of cells) {
      if (xlsxTextIssues(cell.text).length) {
        fail(cell.address, 'unsupported-text', 'Unsupported text in workbook')
      }
    }
    const structural = cells.filter(cell => /^\{[#/]/.test(tagText(cell) ?? ''))
    if (!structural.length) {
      return
    }
    const marker = structural[0]
    const match = /^\{([#/])([\s\S]+)\}$/.exec(tagText(marker)!)
    if (!match) {
      fail(marker.address, 'invalid-tag', 'Invalid structural tag')
    }
    if (structural.length !== 1 || cells.some(cell => cell !== marker && cell.value !== null && cell.value !== '')) {
      fail(marker.address, 'marker-row', 'A structural tag must be alone on its marker row')
    }
    markers.add(row.number)
    const column = Number(marker.col)
    if (match![1] === '/') {
      const index = pending.findLastIndex(open => open.close === match![2] && open.region.column <= column)
      if (index === -1) {
        if (pending.some(open => open.close === match![2])) {
          fail(marker.address, 'region-boundary', 'Closing tag must be at or to the right of its opening tag')
        }
        fail(marker.address, 'unmatched-close', 'Closing tag does not match an opening tag')
      }
      const [{ region }] = pending.splice(index, 1)
      Object.assign(region, { endRow: row.number, width: column - region.column + 1 })
    }
    else {
      const parsed = (() => {
        try {
          return parseShape(match![2])
        }
        catch (error) {
          return fail(marker.address, 'invalid-reference', (error as Error).message)
        }
      })()
      const { source, scope, axis } = parsed
      const region: RegionNode = { id: marker.address, address: marker.address, row: row.number, type: scope ? 'scope' : 'repeat', source, endRow: 0, column, width: 0,
        ...(axis ? { axis } : {}) }
      regions.push(region)
      pending.push({ close: parsed.close, region })
    }
  })
  if (pending.length) {
    fail(pending.at(-1)!.region.address, 'unmatched-open', 'Block has no closing tag')
  }
  const compiled: TaggedRegion[] = []
  for (const region of regions) {
    const parent = compiled.findLast(other => other.row < region.row && other.endRow > region.endRow && region.column >= other.column
      && region.column + region.width <= other.column + other.width)
    compiled.push({ ...region, ...(parent ? { parentId: parent.id } : {}) })
  }
  for (const reference of sheet.model.merges) {
    const range = parseRange(reference)
    if (markers.has(range.start.row) || regions.some(region => [
      { row: region.row, column: region.column }, { row: region.endRow, column: region.column + region.width - 1 },
    ].some(point => point.column >= range.start.column && point.column <= range.end.column && point.row >= range.start.row && point.row <= range.end.row))) {
      fail(reference, 'merge-crosses-block', 'Merge anchors cannot occupy marker rows or cover structural tags')
    }
  }
  sheet.eachRow(row => {
    if (markers.has(row.number)) {
      return
    }
    row.eachCell(cell => {
      if (cell.type === ExcelJS.ValueType.Merge) {
        return
      }
      const text = tagText(cell)
      if (!text?.startsWith('{')) {
        return
      }
      const literal = /^\{\{([^{}]*)\}\}$/.exec(text)
      try {
        const field = literal ? { value: { literal: `{${literal[1]}}` } } : parseFieldTag(text)
        fields.set(cell.address, { expression: field.value, ...('rules' in field && field.rules ? { rules: field.rules } : {}) })
      }
      catch (error) {
        fail(cell.address, 'invalid-reference', (error as Error).message)
      }
    })
  })
  const markerRows = [...markers]
  function authoredRow(row: number, edge: FormulaEdge): number | undefined {
    if (edge === FormulaEdge.Cell && markers.has(row)) {
      return undefined
    }
    return row - markerRows.filter(marker => edge === FormulaEdge.End ? marker <= row : marker < row).length
  }
  return { fields, regions: compiled, markers, authoredRow }
}

import type { GridAddress } from './geometry'
import { columnName, columnNumber } from './geometry'
import { WORKBOOK_LIMITS } from './workbook'

export interface WorkbookFormula { readonly formula: string }
export interface FormulaAddress extends GridAddress { readonly absoluteRow: boolean, readonly absoluteColumn: boolean }
export interface FormulaReference { readonly sheet?: string, readonly start: FormulaAddress, readonly end?: FormulaAddress }
export enum FormulaEdge { Cell = 'cell', Start = 'start', End = 'end' }
export type FormulaRows = (row: number, edge: FormulaEdge) => number | undefined

const maxRow = WORKBOOK_LIMITS.rows
const maxColumn = WORKBOOK_LIMITS.columns
const identifier = /^[\p{L}_\\][\p{L}\p{N}_.\\]*/u
const endpoint = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)(?![\p{L}\p{N}_.\\])/u
function invalid(message: string): never {
  throw new SyntaxError(`Formula: ${message}`)
}
function address(value: FormulaAddress): string {
  return `${value.absoluteColumn ? '$' : ''}${columnName(value.column)}${value.absoluteRow ? '$' : ''}${value.row}`
}
function readAddress(source: string): { value: FormulaAddress, length: number } | undefined {
  const match = endpoint.exec(source)
  if (!match) {
    return undefined
  }
  const value = { row: Number(match[4]), column: columnNumber(match[2]), absoluteRow: !!match[3], absoluteColumn: !!match[1] }
  if (value.row > maxRow || value.column > maxColumn) {
    return undefined
  }
  return { value, length: match[0].length }
}
function referenceText(ref: FormulaReference): string {
  const valid = (value: FormulaAddress) => Number.isInteger(value.row) && value.row > 0 && value.row <= maxRow && Number.isInteger(value.column) && value.column > 0 && value.column <= maxColumn
  if (!valid(ref.start) || ref.end && (!valid(ref.end) || ref.end.row < ref.start.row || ref.end.column < ref.start.column)) {
    return '#REF!'
  }
  const prefix = ref.sheet === undefined ? '' : `'${ref.sheet.replace(/'/g, "''")}'!`
  return prefix + address(ref.start) + (ref.end ? `:${address(ref.end)}` : '')
}

/** Reference tokenizer, not a formula evaluator. Strings/functions/names are opaque. */
function rewriteWorkbookFormula(source: string, map: (ref: FormulaReference) => FormulaReference | undefined): string {
  return compileWorkbookFormula(source)(map)
}

/** Tokenize an authored expression once; each instance supplies only its coordinate map. */
export function compileWorkbookFormula(source: string): (map: (ref: FormulaReference) => FormulaReference | undefined) => string {
  if (typeof source !== 'string' || !source.trim() || source.length > 8192) {
    return invalid('use an expression of 1–8192 characters')
  }
  const parts: (string | FormulaReference)[] = []
  let result = ''
  let cursor = 0
  while (cursor < source.length) {
    const rest = source.slice(cursor)
    const char = source[cursor]
    if (char === '"') {
      const quoted = /^"(?:[^"]|"")*"/.exec(rest)
      if (!quoted) {
        return invalid('unclosed string')
      }
      result += quoted[0]
      cursor += quoted[0].length
      continue
    }
    if (char === '#') {
      const error = /^#(?:REF!|DIV\/0!|VALUE!|N\/A|NAME\?|NUM!|NULL!)/i.exec(rest)
      if (!error) {
        return invalid('spill references are outside this formula scope')
      }
      result += error[0].toUpperCase()
      cursor += error[0].length
      continue
    }
    let sheet: string | undefined
    let offset = 0
    const quotedSheet = /^'((?:[^']|'')+)'!/.exec(rest)
    const name = identifier.exec(rest)?.[0]
    if (quotedSheet) {
      sheet = quotedSheet[1].replace(/''/g, "'")
      offset = quotedSheet[0].length
    }
    else if (name && rest[name.length] === '!') {
      sheet = name
      offset = name.length + 1
    }
    const first = readAddress(rest.slice(offset))
    if (first && (offset || !/^\s*\(/.test(rest.slice(first.length)))) {
      let length = offset + first.length
      const tail = /^\s*:\s*/.exec(rest.slice(length))
      const last = tail ? readAddress(rest.slice(length + tail[0].length)) : undefined
      if (tail && !last) {
        return invalid('use an A1 cell range on one sheet')
      }
      if (last && (last.value.row < first.value.row || last.value.column < first.value.column)) {
        return invalid('write ranges from top left to bottom right')
      }
      if (last) {
        length += tail![0].length + last.length
      }
      parts.push(result, { ...(sheet === undefined ? {} : { sheet }), start: first.value, ...(last ? { end: last.value } : {}) })
      result = ''
      cursor += length
      continue
    }
    if (offset || char === "'" || char === '$' || /[[\]!:;@]/.test(char)) {
      return invalid('use A1 references; external, structured, whole-axis and 3D references are unsupported')
    }
    if (name) {
      result += name
      cursor += name.length
      continue
    }
    const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(rest)?.[0]
    if (number) {
      result += number
      cursor += number.length
      continue
    }
    // Array constants and their locale-dependent separators are a separate contract.
    if (!/[\s+\-*/^&%=<>() ,]/.test(char)) {
      return invalid(`unsupported token ${char}`)
    }
    result += char
    cursor++
  }
  parts.push(result)
  return map => parts.map(part => {
    if (typeof part === 'string') {
      return part
    }
    const ref = map(part)
    return ref ? referenceText(ref) : '#REF!'
  }).join('')
}

export function parseWorkbookFormula(value: unknown): string {
  if (typeof value !== 'string') {
    return invalid('expected a string')
  }
  const source = value.trim().replace(/^=/, '').trim()
  // eslint-disable-next-line no-control-regex
  if (/[\u0000-\u001f\u007f]|_x[0-9a-f]{4}_/i.test(source)) {
    return invalid('unsupported text')
  }
  return rewriteWorkbookFormula(source, ref => ref)
}

/** Carrier rows change addresses across the entire workbook, including other sheets. */
export function mapWorkbookFormulaRows(formula: string, sheet: string, rows: ReadonlyMap<string, FormulaRows>): string {
  return rewriteWorkbookFormula(formula, ref => {
    const map = rows.get((ref.sheet ?? sheet).toLowerCase())
    if (!map) {
      return invalid(`unknown worksheet ${ref.sheet ?? sheet}`)
    }
    const start = map(ref.start.row, ref.end ? FormulaEdge.Start : FormulaEdge.Cell)
    const end = ref.end ? map(ref.end.row, FormulaEdge.End) : undefined
    return start === undefined || ref.end && end === undefined ? undefined : { ...ref, start: { ...ref.start, row: start }, ...(ref.end ? { end: { ...ref.end, row: end! } } : {}) }
  })
}

/** Excel copy semantics; structural moves use a separate coordinate map. */
export function copyWorkbookFormula(formula: string, rows: number, columns: number): string {
  const move = (value: FormulaAddress) => ({ ...value, row: value.row + (value.absoluteRow ? 0 : rows), column: value.column + (value.absoluteColumn ? 0 : columns) })
  return rewriteWorkbookFormula(formula, ref => ({ ...ref, start: move(ref.start), ...(ref.end ? { end: move(ref.end) } : {}) }))
}

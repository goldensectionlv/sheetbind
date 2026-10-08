import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { assertInputData, jsonSnapshot } from '../core/json'
import { validateWorkbookDefinition, collectWorkbookDictionarySources, workbookRegions, WORKBOOK_LIMITS, mapWorkbookPrint, workbookCells } from '../grid/workbook'
import type { GridRange } from '../grid/geometry'
import { partitionWorkbookRange } from '../grid/workbook-coordinates'
import { planWorkbook } from '../grid/workbook-layout'
import type { WorkbookLayout } from '../grid/workbook-layout'
import type { WorkbookCell, WorkbookRegion, WorkbookRow, WorkbookSheet, WorkbookDefinition } from '../grid/workbook'
import { formatAddress, parseRange } from '../grid/geometry'
import { compileWorkbookSheet } from './workbook-tags'
import { writeWorkbookPackage } from './workbook-source'
import { parseDictionaries } from '../core/dictionaries'
import type { Dictionaries } from '../core/dictionaries'
import { readWorkbookPrint } from './workbook-print'
import { FormulaEdge, mapWorkbookFormulaRows, copyWorkbookFormula, parseWorkbookFormula } from '../grid/workbook-formula'
import type { FormulaRows } from '../grid/workbook-formula'
import { readSourceContent } from './source-metadata'
import { TaggedXlsxError, withTemplateLocations } from './tagged-template'
import { loadWorkbook } from './workbook-input'
import { readWorkbookResources, workbookParts } from './workbook-resources'
import type { WorkbookResourceSource } from './workbook-resources'
import type { Cell } from 'exceljs'
import { prepareFormatting } from '../core/formatters'
import type { TaggedXlsxIssue } from './tagged-template'

function unsupported(what: string): never {
  throw new RangeError(`${what} is outside the supported template model`)
}
/** Import a tagged workbook into the execution document. */
export async function importWorkbookXlsx(bytes: Uint8Array): Promise<WorkbookTemplate> {
  let book: ExcelJS.Workbook
  let zip: JSZip
  let parts: Awaited<ReturnType<typeof workbookParts>>
  try {
    zip = await JSZip.loadAsync(bytes)
    parts = await workbookParts(zip)
    book = await loadWorkbook(bytes, zip)
  }
  catch (error) {
    throw new TaggedXlsxError([{ phase: 'template', code: 'invalid-workbook', path: '$workbook', message: 'cannot read a supported unencrypted XLSX template' }], { cause: error })
  }
  const printSettings = await readWorkbookPrint(zip, book)
  const compiledSheets = new Map(book.worksheets.map(sheet => [sheet.name, compileWorkbookSheet(sheet)]))
  const nativeContent = new Map(await Promise.all([...parts].map(async ([name, part]) => [name, await readSourceContent(zip, part.part)] as const)))
  const formulaRows = new Map<string, FormulaRows>([...compiledSheets].map(([name, compiled]) => [name.toLowerCase(), compiled.authoredRow]))
  let id = 0
  function createId(kind: string): string {
    return kind + '-' + ++id
  }
  function importSheet(sheet: ExcelJS.Worksheet): WorkbookSheet {
    // Compile the merged workbook so structural tags cannot occupy merged cells.
    const compiled = compiledSheets.get(sheet.name)!
    const { markers, fields } = compiled
    const part = parts.get(sheet.name)!.part
    function authoredRow(row: number): number {
      return compiled.authoredRow(row, FormulaEdge.Start)!
    }
    if (sheet.rowCount - markers.size > WORKBOOK_LIMITS.rows) {
      unsupported('Authored rows exceed the XLSX format')
    }
    const { entries, regions } = importRegions(compiled, createId)
    const merges = new Map<string, ReturnType<typeof parseRange>>()
    const covered = new Set<string>()
    for (const ref of sheet.model.merges) {
      const range = parseRange(ref)
      const address = formatAddress(range.start)
      merges.set(address, range)
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let column = range.start.column; column <= range.end.column; column++) {
          const cell = formatAddress({ row, column })
          if (cell !== address) {
            covered.add(cell)
          }
        }
      }
    }
    const cells: WorkbookCell[] = []
    const rows: WorkbookRow[] = []
    const sourceCells: string[] = []
    function readCell(cell: ExcelJS.Cell): void {
      if (covered.has(cell.address)) {
        return
      }
      const range = merges.get(cell.address)
      if (cell.value === null && !range && !Object.keys(cell.style).length) {
        return
      }
      const field = fields.get(cell.address)
      let formula: string | undefined
      if (cell.type === ExcelJS.ValueType.Formula && Reflect.get(cell.value as object, 'shareType') !== 'array') {
        try {
          formula = mapWorkbookFormulaRows(readWorkbookFormula(cell), sheet.name, formulaRows)
        }
        catch { /* Unsupported expressions remain in the source cell. */ }
      }
      const scalar = typeof cell.value === 'string' || typeof cell.value === 'number' || typeof cell.value === 'boolean'
        ? cell.value
        : cell.type === ExcelJS.ValueType.RichText || cell.type === ExcelJS.ValueType.Hyperlink ? cell.text : null
      const literal = typeof scalar === 'string' && /_x[0-9a-f]{4}_|[{}]/i.test(scalar) ? null : scalar
      const row = Number(cell.row)
      const column = Number(cell.col)
      const index = authoredRow(row)
      const owner = entries.findLast(({ node }) => row > node.row && row < node.endRow && column >= node.column && column < node.column + node.width)
      const target = owner?.region.cells ?? cells
      const value = field?.expression ?? (formula === undefined ? { literal } : { formula })
      target.push({ id: createId('cell'), at: { row: owner ? index - owner.row + 1 : index, column: column - (owner ? owner.node.column - 1 : 0) },
        size: { rows: range ? compiled.authoredRow(range.end.row, FormulaEdge.End)! - index + 1 : 1, columns: range ? range.end.column - range.start.column + 1 : 1 },
        value, xlsx: { address: cell.address, value: field && 'literal' in field.expression ? { literal: scalar } : value },
        ...(field?.rules ? { rules: field.rules } : {}) })
      sourceCells.push(cell.address)
    }
    function readRow(row: ExcelJS.Row): void {
      if (markers.has(row.number)) {
        return
      }
      const index = authoredRow(row.number)
      const owner = entries.findLast(({ node }) => node.width === WORKBOOK_LIMITS.columns && row.number > node.row && row.number < node.endRow)
      if (row.height !== undefined || row.hidden) {
        (owner?.region.rows ?? rows).push({ index: owner ? index - owner.row + 1 : index, ...(row.height !== undefined ? { height: row.height } : {}), ...(row.hidden ? { hidden: true } : {}) })
      }
      row.eachCell({ includeEmpty: true }, readCell)
    }
    sheet.eachRow({ includeEmpty: true }, readRow)
    const occupied: GridRange[] = []
    const regionViews = workbookRegions({ cells: [], regions })
    const regionBodies = new Map(entries.map(entry => [entry.region.id, entry.region]))
    for (const range of nativeContent.get(sheet.name) ?? []) {
      const start = { ...range.start, row: compiled.authoredRow(range.start.row, FormulaEdge.Start)! }
      const end = { ...range.end, row: compiled.authoredRow(range.end.row, FormulaEdge.End)! }
      if (start.row > end.row) {
        continue
      }
      for (const piece of partitionWorkbookRange(regionViews, { start, end })) {
        const owner = piece.owner
        const target = owner ? regionBodies.get(owner.id)!.occupied : occupied
        const offset = { row: (owner?.row ?? 1) - 1, column: (owner?.column ?? 1) - 1 }
        target.push({ start: { row: piece.range.start.row - offset.row, column: piece.range.start.column - offset.column },
          end: { row: piece.range.end.row - offset.row, column: piece.range.end.column - offset.column } })
      }
    }
    const columns = (sheet.columns ?? []).flatMap((column, index) => {
      return column.width !== undefined || column.hidden ? [{ index: index + 1, ...(column.width !== undefined ? { width: column.width } : {}), ...(column.hidden ? { hidden: true } : {}) }] : []
    })
    const print = mapWorkbookPrint(printSettings.get(sheet.name), { rowStart: authoredRow, rowEnd: row => compiled.authoredRow(row, FormulaEdge.End)! })
    return { id: createId('sheet'), name: sheet.name, state: sheet.state, xlsx: { part, markers: [...markers], cells: sourceCells }, ...(print ? { print } : {}),
      rows, columns, cells, ...(occupied.length ? { occupied } : {}), ...(regions.length ? { regions } : {}) }
  }
  return new WorkbookTemplate(validateWorkbookDefinition({ sheets: book.worksheets.map(importSheet) }), bytes, await readWorkbookResources(zip))
}

type ImportedRegion = Omit<WorkbookRegion, 'cells' | 'rows' | 'regions' | 'occupied'> & { cells: WorkbookCell[], rows: WorkbookRow[], occupied: GridRange[], regions?: ImportedRegion[] }
function importRegions(compiled: ReturnType<typeof compileWorkbookSheet>, createId: (kind: string) => string) {
  const entries = new Map<string, { node: typeof compiled.regions[number], row: number, region: ImportedRegion }>()
  const regions: ImportedRegion[] = []
  for (const node of compiled.regions) {
    const parent = node.parentId ? entries.get(node.parentId) : undefined
    const row = compiled.authoredRow(node.row, FormulaEdge.Start)!
    const region: ImportedRegion = {
      id: createId('region'), xlsx: { address: node.address }, type: node.type, source: node.source, row: row - (parent ? parent.row - 1 : 0),
      ...(node.column !== 1 ? { column: node.column - (parent ? parent.node.column - 1 : 0) } : {}),
      width: node.width, ...(node.axis ? { axis: node.axis } : {}),
      height: compiled.authoredRow(node.endRow, FormulaEdge.Start)! - row, cells: [], rows: [], occupied: [],
    }
    if (parent) {
      (parent.region.regions ??= []).push(region)
    }
    else {
      regions.push(region)
    }
    entries.set(node.id, { node, row, region })
  }
  return { entries: [...entries.values()], regions }
}

/** Resolve values and placement without exposing the imported definition. */
export function resolveWorkbook(template: WorkbookTemplate, data: unknown, options: { readonly dictionaries?: Dictionaries } = {}): WorkbookLayout {
  const { definition } = WorkbookTemplate.content(template)
  assertInputData(data)
  const plan = withTemplateLocations(definition, () => planWorkbook(definition, data, { dictionaries: parseDictionaries(options.dictionaries, collectWorkbookDictionarySources(definition)) }))
  return { sheets: plan.sheets.map(({ sheet }) => ({
    ...structuredClone({ id: sheet.id, name: sheet.name, state: sheet.state, rows: sheet.rows, columns: sheet.columns, print: sheet.print }),
    cells: sheet.cells.map(cell => ({
      ...structuredClone({ id: cell.id, definitionId: cell.definitionId, at: cell.at, size: cell.size, rules: cell.rules,
        origin: cell.origin, contextPath: cell.contextPath, value: cell.value }),
      choice: cell.choice && { ...cell.choice, items: cell.choice.items.map(item => ({ ...item, value: jsonSnapshot(item.value) as Record<string, unknown> })) },
    })),
  })) }
}

export function workbookDictionarySources(template: WorkbookTemplate): string[] {
  return collectWorkbookDictionarySources(WorkbookTemplate.content(template).definition)
}

/** Render data into the imported workbook while preserving its native content. */
export async function renderWorkbookReport(template: WorkbookTemplate, data: unknown, options: { readonly dictionaries?: Dictionaries } = {}): Promise<Buffer> {
  const { definition, source, resources } = WorkbookTemplate.content(template)
  const dictionaries = parseDictionaries(options.dictionaries, collectWorkbookDictionarySources(definition))
  assertInputData(data)
  const plan = withTemplateLocations(definition, () => planWorkbook(definition, data, { dictionaries }))
  return writeWorkbookPackage({ source, plan, dictionaries, resources })
}
/** An imported XLSX and its compiled bindings, owned by the library. */
export class WorkbookTemplate {
  readonly #content: { readonly definition: WorkbookDefinition, readonly source: Uint8Array, readonly resources: WorkbookResourceSource }

  constructor(definition: WorkbookDefinition, source: Uint8Array, resources: WorkbookResourceSource) {
    this.#content = { definition, source: Uint8Array.from(source), resources }
  }

  static content(template: WorkbookTemplate) {
    if (!(template instanceof WorkbookTemplate)) {
      throw new TypeError('Load a tagged XLSX with importWorkbookXlsx first')
    }
    return template.#content
  }
}
/** Materialize shared formulas using string-aware reference offsets. */
function readWorkbookFormula(cell: Cell): string {
  const value = cell.value
  if (!value || typeof value !== 'object' || 'shareType' in value && value.shareType !== 'shared') {
    throw new SyntaxError('Expected an ordinary or shared formula')
  }
  if ('formula' in value) {
    return parseWorkbookFormula(value.formula)
  }
  if ('sharedFormula' in value && typeof value.sharedFormula === 'string') {
    const master = cell.worksheet.getCell(value.sharedFormula)
    const source = master.value
    if (source && typeof source === 'object' && 'formula' in source) {
      return copyWorkbookFormula(parseWorkbookFormula(source.formula), Number(cell.row) - Number(master.row), Number(cell.col) - Number(master.col))
    }
  }
  throw new SyntaxError('Shared formula master is missing')
}
export interface WorkbookTemplateFinding extends TaggedXlsxIssue {
  readonly severity: 'error' | 'warning'
}

/** Check render dependencies without data, formatter execution or value validation. */
export function inspectWorkbookTemplate(template: WorkbookTemplate, options: { readonly dictionaries?: readonly string[] } = {}): readonly WorkbookTemplateFinding[] {
  const findings: WorkbookTemplateFinding[] = []
  const known = options.dictionaries && new Set(options.dictionaries)
  for (const sheet of WorkbookTemplate.content(template).definition.sheets) {
    for (const cell of workbookCells(sheet)) {
      const location = { phase: 'template' as const, nodeId: cell.id, path: cell.xlsx.address, sheetName: sheet.name, address: cell.xlsx.address }
      if (cell.rules?.format) {
        try {
          prepareFormatting(cell.rules.format)
        }
        catch (error) {
          findings.push({ ...location, code: 'invalid-format', severity: 'warning', message: (error as Error).message })
        }
      }
      const choice = cell.rules?.choice?.source
      const dependencies = new Set([cell.rules?.list, choice && 'dictionary' in choice ? choice.dictionary : undefined])
      for (const name of dependencies) {
        if (name && known && !known.has(name)) {
          findings.push({ ...location, code: 'unknown-dict', severity: 'warning', message: `Unknown dictionary: ${name}` })
        }
      }
    }
  }
  return findings
}

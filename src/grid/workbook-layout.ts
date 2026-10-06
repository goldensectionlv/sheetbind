import { instantiate, TemplateError } from '../core/template'
import type { Fragment, Origin, ResolvedFragment, TemplateValue } from '../core/template'
import type { Dictionaries } from '../core/dictionaries'
import type { FieldRules } from '../core/field-rules'
import type { GridAddress, GridOffset } from './geometry'
import { createWorkbookChoiceDisplay } from './workbook-choice-display'
import type { WorkbookChoice } from './workbook-choice-display'
import { workbookCells, workbookRegions, workbookRows, workbookIssue, WorkbookAxis, WORKBOOK_LIMITS } from './workbook'
import type { WorkbookBody, WorkbookCell, WorkbookSheet, WorkbookDefinition, WorkbookRow, WorkbookColumn } from './workbook'
import { mapWorkbookPrint } from './workbook-print'
import type { WorkbookPrint } from './workbook-print'
import type { WorkbookFormula } from './workbook-formula'
import { resolveWorkbookFormulas } from './workbook-formulas'
import { axisPositions, mapAxis, planWorkbookAxes } from './workbook-axis'
import type { AxisPlan, WorkbookAxes } from './workbook-axis'
import { FormulaEdge } from './workbook-formula'
import { placeWorkbookPoint, workbookCoordinates } from './workbook-coordinates'
import type { WorkbookCoordinates, WorkbookIndexes } from './workbook-coordinates'

/** Public resolved cell; source-package details stay in the internal placement. */
export interface WorkbookCellInstance {
  readonly id: string
  readonly definitionId: string
  readonly at: GridAddress
  readonly size: GridOffset
  readonly rules?: FieldRules
  readonly origin: Origin
  readonly contextPath: string
  readonly choice?: WorkbookChoice
  readonly value: { readonly literal: TemplateValue } | WorkbookFormula
}
export interface WorkbookLayout {
  readonly sheets: readonly {
    readonly id: string
    readonly name: string
    readonly state?: 'visible' | 'hidden' | 'veryHidden'
    readonly rows?: readonly WorkbookRow[]
    readonly columns?: readonly WorkbookColumn[]
    readonly print?: WorkbookPrint
    readonly cells: readonly WorkbookCellInstance[]
  }[]
}

/** Internal placement retains the source references needed by the XLSX writer. */
export interface WorkbookPlacedCell extends WorkbookCellInstance { readonly xlsx?: WorkbookCell['xlsx'] }
export interface WorkbookPlacement { readonly sheets: readonly (Omit<WorkbookSheet, 'cells' | 'regions'> & { readonly cells: readonly WorkbookPlacedCell[] })[] }

/** Placement provenance for consumers that need the boundaries of expanded bodies. */
export interface WorkbookRegionLayout {
  readonly definitionId: string
  readonly dataPath: string
  readonly type: 'scope' | 'repeat'
  readonly row: number
  readonly height: number
  readonly column?: number
  readonly width?: number
  readonly instances: readonly { readonly row: number, readonly height: number, readonly regions: readonly WorkbookRegionLayout[] }[]
}
export interface WorkbookPlan {
  readonly layout: WorkbookPlacement
  readonly regions: readonly (readonly WorkbookRegionLayout[])[]
  readonly axes: readonly WorkbookAxes[]
  readonly coordinates: readonly WorkbookCoordinates[]
  readonly extents: readonly { readonly rows: number, readonly columns: number }[]
}

interface WorkbookExecution {
  readonly groups: readonly ResolvedFragment<null>[]
  readonly origin: (value: Origin) => Origin
}
type WorkbookExecutionOptions = { readonly dictionaries?: Dictionaries, readonly checkValues?: boolean }
function executionKey(kind: string, id: string): string {
  return JSON.stringify([kind, id])
}

/** Expose independent resolved cells without changing the compiled definition. */
export function resolveWorkbook(config: WorkbookDefinition, data: unknown, options: { dictionaries?: Dictionaries } = {}): WorkbookLayout {
  const { layout } = planWorkbook(config, data, { ...options, checkValues: true })
  return { sheets: layout.sheets.map(sheet => ({
    ...structuredClone({ id: sheet.id, name: sheet.name, state: sheet.state, rows: sheet.rows, columns: sheet.columns, print: sheet.print }),
    cells: sheet.cells.map(cell => structuredClone({
      id: cell.id, definitionId: cell.definitionId, at: cell.at, size: cell.size, rules: cell.rules,
      origin: cell.origin, contextPath: cell.contextPath, choice: cell.choice, value: cell.value,
    })),
  })) }
}

/** Accept a normalized definition. Placement never reparses or modifies the input. */
export function planWorkbook(config: WorkbookDefinition, data: unknown, options: WorkbookExecutionOptions = {}): WorkbookPlan {
  const { plan, sources } = placeWorkbook(config, data, options)
  return { ...plan, layout: resolveWorkbookFormulas(plan.layout, sources) }
}

/** Keep authored formulas until all placement transforms are known. */
export function placeWorkbook(config: WorkbookDefinition, data: unknown, options: WorkbookExecutionOptions = {}) {
  const execution = executeWorkbook(config, data, options)
  const placed = config.sheets.map((sheet, index) => placeWorkbookSheet(sheet, execution.groups[index], execution.origin))
  const sources = new Map(placed.map(value => [value.sheet.name.toLowerCase(), { axes: value.axes, authored: value.authored }]))
  const plan: WorkbookPlan = { layout: { sheets: placed.map(value => value.sheet) }, regions: placed.map(value => value.regions), axes: placed.map(value => value.axes), coordinates: placed.map(value => value.coordinates), extents: placed.map(value => value.extent) }
  return { plan, sources }
}

function executeWorkbook(config: WorkbookDefinition, data: unknown, options: WorkbookExecutionOptions): WorkbookExecution {
  const identities = new Map<string, string>()
  function key(kind: string, id: string): string {
    const key = executionKey(kind, id)
    identities.set(key, id)
    return key
  }
  function children(body: WorkbookBody): Fragment<null>[] {
    return [
      ...body.cells.map(cell => ({ type: 'value' as const, id: key('cell', cell.id), value: 'formula' in cell.value ? { literal: null } : cell.value, rules: cell.rules })),
      ...(body.regions ?? []).map(region => ({ type: region.type, id: key('region', region.id), source: region.source,
        body: { type: 'group' as const, id: key('body', region.id), content: null, children: children(region) } })),
    ]
  }
  const template: Fragment<null> = { type: 'group', id: key('workbook', '$template'), content: null, children: config.sheets.map(sheet => ({
    type: 'group', id: key('sheet', sheet.id), content: null, children: children(sheet),
  })) }
  let resolved: ResolvedFragment<null>
  try {
    resolved = instantiate(template, data, options)
  }
  catch (error) {
    if (error instanceof TemplateError) {
      throw new TemplateError(error.issues.map(issue => ({ ...issue, nodeId: identities.get(issue.nodeId) ?? issue.nodeId })))
    }
    throw error
  }
  if (resolved.type !== 'group') {
    throw new Error('Expected workbook group')
  }
  const iterations = new WeakMap<Origin['iterations'], Origin['iterations']>()
  function origin(value: Origin): Origin {
    let mapped = iterations.get(value.iterations)
    if (!mapped) {
      mapped = value.iterations.map(item => ({ ...item, nodeId: identities.get(item.nodeId)! }))
      iterations.set(value.iterations, mapped)
    }
    return { ...value, nodeId: identities.get(value.nodeId)!, iterations: mapped }
  }
  return { groups: resolved.children, origin }
}

function placeWorkbookSheet(sheet: WorkbookSheet, group: ResolvedFragment<null>, origin: (value: Origin) => Origin) {
  if (group.type !== 'group') {
    throw new Error('Expected sheet group')
  }
  const { regions, cells: definitions, ...settings } = sheet
  const cells: WorkbookPlacedCell[] = []
  const axes = planWorkbookAxes(sheet, group, id => executionKey('region', id))
  const extent = { rows: 0, columns: 0 }
  const authored = new Map(workbookCells(sheet).map(cell => [cell.id, cell]))
  const views = new Map(workbookRegions(sheet).map(region => [region.id, region]))
  const displayChoice = createWorkbookChoiceDisplay()
  const contexts = new WeakMap<Origin['iterations'], { indexes: ReadonlyMap<string, number>, suffix: string }>()
  const instances = new Map<string, WorkbookIndexes[]>()
  function context(value: Origin) {
    let found = contexts.get(value.iterations)
    if (!found) {
      found = { indexes: new Map(value.iterations.map(item => [item.nodeId, item.index])), suffix: JSON.stringify(value.iterations) + ']' }
      contexts.set(value.iterations, found)
    }
    return found
  }
  const ids = new Map(workbookCells(sheet).map(cell => [cell.id, '[' + JSON.stringify(cell.id) + ',']))
  const nodeKeys = new Map([...authored.keys()].map(id => [id, executionKey('cell', id)]))
  function placeCells(definitions: readonly WorkbookCell[], nodes: readonly ResolvedFragment<null>[], contextPath: string): WorkbookPlacedCell[] {
    const byId = new Map(nodes.filter(node => node.type === 'value').map(node => [node.origin.nodeId, node]))
    return definitions.map(cell => {
      const value = byId.get(nodeKeys.get(cell.id)!)!
      const source = origin(value.origin)
      const position = authored.get(cell.id)!.at
      const { indexes, suffix } = context(source)
      const at = placeWorkbookPoint(axes, position, indexes)
      if (!at) {
        workbookIssue('growth-crosses-cell', cell.id, 'A fixed cell occupies a removed band; put it inside the repeat or outside its band', source.dataPath, 'data')
      }
      const end = { row: cell.size.rows === 1 ? at.row : mapAxis(axes.rows, position.row + cell.size.rows - 1, FormulaEdge.End, indexes)!,
        column: cell.size.columns === 1 ? at.column : mapAxis(axes.columns, position.column + cell.size.columns - 1, FormulaEdge.End, indexes)! }
      if (at.row === undefined || at.column === undefined || end.row === undefined || end.column === undefined) {
        workbookIssue('growth-crosses-cell', cell.id, 'A fixed cell occupies a removed band; put it inside the repeat or outside its band', source.dataPath, 'data')
      }
      const rows = end.row - at.row + 1
      const columns = end.column - at.column + 1
      const size = rows === cell.size.rows && columns === cell.size.columns ? cell.size : { rows, columns }
      if (end.row > WORKBOOK_LIMITS.rows) {
        workbookIssue('row-limit', cell.id, 'Rendered content exceeds the XLSX row limit', source.dataPath, 'data')
      }
      if (end.column > WORKBOOK_LIMITS.columns) {
        workbookIssue('column-limit', cell.id, 'Rendered content exceeds the XLSX column limit', source.dataPath, 'data')
      }
      extent.rows = Math.max(extent.rows, end.row)
      extent.columns = Math.max(extent.columns, end.column)
      let choice: WorkbookChoice | undefined
      try {
        choice = value.choice ? displayChoice(value.choice) : undefined
      }
      catch (error) {
        workbookIssue('choice-display', cell.id, (error as Error).message, source.dataPath, 'data')
      }
      const literal = value.value !== null && typeof value.value === 'object' ? choice!.text : value.value
      return { id: ids.get(cell.id)! + suffix, definitionId: cell.id, at, size,
        value: 'path' in cell.value ? { literal } : cell.value, rules: cell.rules, xlsx: cell.xlsx, origin: source, contextPath, choice }
    })
  }
  function bounds(id: string, value: Origin) {
    const region = views.get(id)!
    const indexes = new Map(origin(value).iterations.map(item => [item.nodeId, item.index]))
    const row = mapAxis(axes.rows, region.row, FormulaEdge.Start, indexes)!
    const column = mapAxis(axes.columns, region.column ?? 1, FormulaEdge.Start, indexes)!
    return { row, column, height: mapAxis(axes.rows, region.row + region.height - 1, FormulaEdge.End, indexes)! - row + 1,
      width: mapAxis(axes.columns, (region.column ?? 1) + (region.width ?? WORKBOOK_LIMITS.columns) - 1, FormulaEdge.End, indexes)! - column + 1 }
  }
  function layout(definition: WorkbookBody, body: ResolvedFragment<null>, id?: string): WorkbookRegionLayout[] {
    if (body.type !== 'group') {
      throw new Error('Expected region body')
    }
    if (id) {
      const copies = instances.get(id) ?? []
      copies.push(context(origin(body.origin)).indexes)
      instances.set(id, copies)
    }
    const owner = id ? views.get(id) : undefined
    for (const range of definition.occupied ?? []) {
      const end = placeWorkbookPoint(axes, {
        row: range.end.row + (owner?.row ?? 1) - 1,
        column: range.end.column + (owner?.column ?? 1) - 1,
      }, context(origin(body.origin)).indexes)!
      extent.rows = Math.max(extent.rows, end.row)
      extent.columns = Math.max(extent.columns, end.column)
    }
    const nodes = new Map(body.children.map(node => [node.origin.nodeId, node]))
    const placedRegions: WorkbookRegionLayout[] = []
    for (const region of [...definition.regions ?? []].sort((a, b) => a.row - b.row)) {
      const node = nodes.get(executionKey('region', region.id))!
      const instances = node.type === 'repeat' ? node.instances : node.type === 'scope' ? [node.body] : []
      const placed = bounds(region.id, node.origin)
      if (placed.row + placed.height - 1 > WORKBOOK_LIMITS.rows) {
        workbookIssue('row-limit', region.id, 'Repeated rows exceed the XLSX row limit', node.origin.dataPath, 'data')
      }
      if (region.width !== undefined && placed.column + placed.width - 1 > WORKBOOK_LIMITS.columns) {
        workbookIssue('column-limit', region.id, 'Repeated columns exceed the XLSX column limit', node.origin.dataPath, 'data')
      }
      const placedInstances = instances.map(instance => ({ ...bounds(region.id, instance.origin), regions: layout(region, instance, region.id) }))
      const last = placedInstances.at(-1)
      const extent = region.type !== 'repeat'
        ? {}
        : region.axis === WorkbookAxis.Columns
          ? { width: last ? last.column + last.width - placed.column : 0 }
          : { height: last ? last.row + last.height - placed.row : 0 }
      placedRegions.push({ definitionId: region.id, dataPath: node.origin.dataPath, type: region.type, ...placed, ...extent, instances: placedInstances })
    }
    cells.push(...placeCells(definition.cells, body.children, body.origin.dataPath))
    return placedRegions
  }
  const placedRegions = layout({ cells: definitions, regions, occupied: sheet.occupied }, group)
  const print = mapWorkbookPrint(sheet.print, {
    rowStart: row => mapAxis(axes.rows, row, FormulaEdge.Start)!, rowEnd: row => mapAxis(axes.rows, row, FormulaEdge.End)!,
    columnStart: column => mapAxis(axes.columns, column, FormulaEdge.Start)!, columnEnd: column => mapAxis(axes.columns, column, FormulaEdge.End)!,
  })
  if (print?.area && (print.area.end.row > WORKBOOK_LIMITS.rows || print.area.end.column > WORKBOOK_LIMITS.columns) || print?.repeatRows && print.repeatRows.end > WORKBOOK_LIMITS.rows) {
    workbookIssue('print-limit', sheet.id, 'Rendered print range exceeds the XLSX grid', `$template.${sheet.id}.print`, 'data')
  }
  const rows = placeSettings(workbookRows(sheet), axes.rows, WORKBOOK_LIMITS.rows)
  const columns = placeSettings(sheet.columns ?? [], axes.columns, WORKBOOK_LIMITS.columns)
  const occupied = new Set<number>()
  for (const cell of cells) {
    for (let row = cell.at.row; row < cell.at.row + cell.size.rows; row++) {
      for (let column = cell.at.column; column < cell.at.column + cell.size.columns; column++) {
        const key = row * WORKBOOK_LIMITS.columns + column
        if (occupied.has(key)) {
          workbookIssue('placed-overlap', cell.definitionId, 'Rendered regions overlap', cell.origin.dataPath, 'data')
        }
        occupied.add(key)
      }
    }
  }
  const placed = { ...structuredClone(settings), ...(print ? { print } : {}), rows, columns, cells: cells.sort((a, b) => a.at.row - b.at.row || a.at.column - b.at.column) }
  return { sheet: placed, regions: placedRegions, axes, coordinates: workbookCoordinates(sheet, axes, instances), authored, extent }
}

function placeSettings<T extends { readonly index: number }>(settings: readonly T[], axis: AxisPlan, limit: number): T[] {
  const result = new Map<number, T>()
  for (const setting of settings) {
    for (const index of axisPositions(axis, setting.index)) {
      if (index > limit) {
        throw new RangeError('Rendered row/column settings exceed the XLSX grid')
      }
      const placed = { ...setting, index }
      const previous = result.get(index)
      if (previous && Object.entries(placed).some(([key, value]) => Reflect.get(previous, key) !== value)) {
        throw new RangeError('Conflicting settings on a shared row or column')
      }
      result.set(index, placed)
    }
  }
  return [...result.values()].sort((a, b) => a.index - b.index)
}

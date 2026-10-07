import { createValidation } from '../core/validation'
import type { ValidationOptions, ValidateValue } from '../core/validation'
import type { JsonValue } from '../core/json'
import { parseDictionaries } from '../core/dictionaries'
import type { Dictionaries } from '../core/dictionaries'
import { TemplateError } from '../core/template'
import type { DataReference } from '../core/template'
import { workbookCells, workbookRegions, WorkbookAxis } from '../grid/workbook'
import type { WorkbookBody, WorkbookCell, WorkbookDefinition, WorkbookRegion } from '../grid/workbook'
import { dataPath, readData, writeData } from './records'
import type { DataPath } from './records'

export interface WorkbookFormIssue {
  readonly rule?: string
  readonly args?: readonly JsonValue[]
  readonly index?: number
  readonly phase: 'structure' | 'value' | 'xlsx'
  readonly code: string
  readonly path: string
  readonly message: string
  readonly sheetName?: string
  readonly address?: string
  readonly nodeId?: string
}
export type WorkbookFormResult = { readonly success: true, readonly data: Readonly<Record<string, unknown>> }
  | { readonly success: false, readonly issues: readonly WorkbookFormIssue[], readonly data?: Readonly<Record<string, unknown>> }

/** Internal early exit for a malformed carrier; structural failures cannot provide trustworthy data. */
export class WorkbookFormInputError extends Error {
  constructor(readonly issue: WorkbookFormIssue) {
    super(issue.message)
  }
}

export function referencePath(reference: DataReference, context: DataPath): DataPath {
  return [...(reference.from === 'root' ? [] : context), ...reference.path.split('.')]
}

export type WorkbookFormOptions = ValidationOptions & { readonly dictionaries?: Dictionaries }
export interface PreparedWorkbookField extends Pick<WorkbookCell, 'id' | 'rules'> {
  readonly reference: DataReference
  readonly validate?: ValidateValue
}
export interface PreparedWorkbookForm {
  readonly template: WorkbookDefinition
  readonly fields: ReadonlyMap<string, PreparedWorkbookField>
  readonly dictionaries: Dictionaries
}
/** Data bindings define input fields; other worksheet content is not submitted. */
export function prepareWorkbookForm(value: WorkbookDefinition, purpose: 'issue' | 'read', options: WorkbookFormOptions = {}): PreparedWorkbookForm {
  const template = expandWorkbookScopes(value)
  const dictionaries = parseDictionaries(options.dictionaries === undefined ? {} : options.dictionaries)
  if (template.sheets.some(sheet => workbookRegions(sheet).some(region => region.axis === WorkbookAxis.Columns))) {
    throw new TemplateError([{ phase: 'template', code: 'report-only-region', nodeId: '', path: '$template', message: 'Form records repeat down rows; column repeats support reports' }])
  }
  const definitions = template.sheets.flatMap(sheet => workbookCells(sheet))
  const prepareValidation = purpose === 'read' ? createValidation(options) : undefined
  const fields = new Map<string, PreparedWorkbookField>()
  for (const cell of definitions) {
    try {
      const validate = prepareValidation?.(cell.rules?.validation, cell.rules?.validationMessages)
      if ('path' in cell.value) {
        fields.set(cell.id, { id: cell.id, reference: cell.value, rules: cell.rules, validate })
      }
    }
    catch (error) {
      throw new TemplateError([{ phase: 'template', code: 'invalid-rules', nodeId: cell.id, path: cell.id, message: (error as Error).message }])
    }
  }
  assertFormBindings(template)
  return { template, fields, dictionaries }
}

function assertFormBindings(template: WorkbookDefinition): void {
  const shapes = new Map<string, 'object' | 'array' | 'field'>()
  const lookups: string[] = []
  let fields = 0
  for (const sheet of template.sheets) {
    collect(sheet, '')
  }
  if (lookups.some(path => shapes.has(path) || [...shapes].some(([parent, kind]) => kind === 'field' && path.startsWith(parent + '.')))) {
    throw new TemplateError([{ phase: 'template', code: 'conflicting-source', nodeId: '', path: '$template', message: 'Form choice sources must be separate from submitted fields and collections' }])
  }
  if (!fields) {
    throw new TemplateError([{ phase: 'template', code: 'no-form-fields', nodeId: '', path: '$template', message: 'a form needs at least one data binding' }])
  }

  function claim(path: string, kind: 'object' | 'array' | 'field'): void {
    const previous = shapes.get(path)
    if (previous && previous !== kind) {
      throw new TemplateError([{ phase: 'template', code: 'conflicting-binding', nodeId: '', path,
        message: 'a form binding cannot also own an incompatible container or field' }])
    }
    shapes.set(path, kind)
  }
  function declare(path: string, kind: 'array' | 'field'): void {
    const parts = path.split('.')
    for (let index = 1; index < parts.length; index++) {
      claim(parts.slice(0, index).join('.'), parts[index] === '*' ? 'array' : 'object')
    }
    claim(path, kind)
  }
  function collect(body: WorkbookBody, context: string): void {
    const regions = [...body.regions ?? []].sort((a, b) => a.row - b.row)
    if (regions.some((region, index) => index && region.row < regions[index - 1].row + regions[index - 1].height)) {
      throw new TemplateError([{ phase: 'template', code: 'overlapping-form-regions', nodeId: '', path: '$template', message: 'Independent form regions need separate row bands' }])
    }
    const path = (ref: DataReference) => `${ref.from === 'root' ? '' : context}${ref.path}`
    for (const cell of body.cells) {
      if ('path' in cell.value) {
        fields++
        declare(path(cell.value), 'field')
      }
      if (cell.rules?.choice && 'path' in cell.rules.choice.source) {
        lookups.push(path(cell.rules.choice.source))
      }
    }
    for (const region of body.regions ?? []) {
      const source = path(region.source)
      declare(source, 'array')
      collect(region, `${source}.*.`)
    }
  }
}

/** Shapes come from validated template paths and bounded marker counts, never uploaded path strings. */
export function createWorkbookFormData() {
  const data: Record<string, unknown> = {}
  const fields = new Set<string>()
  const contexts = new Map<string, DataPath>()
  const lengths = new Map<string, number>()
  const fail = (path: DataPath): never => {
    throw new WorkbookFormInputError({ phase: 'structure', code: 'conflicting-shape', path: dataPath(path), message: 'repeated views disagree about the data structure' })
  }
  const set = (path: DataPath, value: unknown) => {
    try {
      writeData(data, path, value, true)
    }
    catch {
      fail(path)
    }
  }
  return {
    data, contexts,
    context(path: DataPath) {
      contexts.set(dataPath(path), path)
    },
    collection(path: DataPath, count: number) {
      const name = dataPath(path)
      const previous = lengths.get(name)
      if (previous !== undefined) {
        if (previous !== count) {
          fail(path)
        }
        return
      }
      if (fields.has(name) || readData(data, path) !== undefined) {
        fail(path)
      }
      lengths.set(name, count)
      set(path, Array.from({ length: count }, () => ({})))
    },
    field(path: DataPath) {
      const name = dataPath(path)
      if (fields.has(name)) {
        return
      }
      if (readData(data, path) !== undefined) {
        fail(path)
      }
      fields.add(name)
      set(path, null)
    },
  }
}

/** Remove object scopes using explicit paths; repeats establish a new current item. */
export function expandWorkbookScopes(template: WorkbookDefinition): WorkbookDefinition {
  const expand = <T extends WorkbookBody>(body: T, prefix?: WorkbookRegion['source'], local = false): T => {
    const qualify = (reference: WorkbookRegion['source']) => reference.from === 'root' || !prefix
      ? reference
      : { ...reference, path: `${prefix.path}.${reference.path}`, from: prefix.from }
    const cells = body.cells.map(cell => ({ ...cell, value: 'path' in cell.value ? qualify(cell.value) : cell.value,
      ...(cell.rules?.choice && 'path' in cell.rules.choice.source ? { rules: { ...cell.rules, choice: { ...cell.rules.choice, source: qualify(cell.rules.choice.source) } } } : {}),
    }))
    const rows = [...body.rows ?? []]
    const occupied = [...body.occupied ?? []]
    const regions: WorkbookRegion[] = []
    for (const region of body.regions ?? []) {
      const source = qualify(region.source)
      if (region.type === 'repeat') {
        regions.push({ ...expand(region, undefined, true), source })
        continue
      }
      const expanded = expand(region, { ...source, from: source.from === 'root' || !local ? 'root' : 'current' }, local)
      cells.push(...expanded.cells.map(cell => ({ ...cell, at: { row: cell.at.row + region.row - 1, column: cell.at.column + (region.column ?? 1) - 1 } })))
      rows.push(...(expanded.rows ?? []).map(row => ({ ...row, index: row.index + region.row - 1 })))
      occupied.push(...(expanded.occupied ?? []).map(range => ({
        start: { row: range.start.row + region.row - 1, column: range.start.column + (region.column ?? 1) - 1 },
        end: { row: range.end.row + region.row - 1, column: range.end.column + (region.column ?? 1) - 1 },
      })))
      regions.push(...(expanded.regions ?? []).map(child => ({ ...child, row: child.row + region.row - 1,
        ...(region.column !== undefined || child.column !== undefined ? { column: (child.column ?? 1) + (region.column ?? 1) - 1 } : {}),
        ...(region.width !== undefined && child.width === undefined ? { width: region.width } : {}) })))
    }
    return { ...body, cells, ...(occupied.length ? { occupied } : {}), ...(body.rows || rows.length ? { rows: rows.sort((a, b) => a.index - b.index) } : {}), ...(body.regions ? { regions } : {}) }
  }
  return { ...template, sheets: template.sheets.map(sheet => expand(sheet)) }
}

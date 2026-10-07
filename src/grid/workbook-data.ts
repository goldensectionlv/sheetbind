import { isDataObject } from '../core/json'
import { isBlank } from '../core/validation'
import { compileDataPath } from '../core/template'
import { TemplateError } from '../core/template'
import type { DataReference, Origin, TemplateIssue, TemplateValue } from '../core/template'
import { validateList } from '../core/field-rules'
import type { Dictionaries } from '../core/dictionaries'
import { allowsChoiceInput, choiceKey, createChoiceResolver, createChoiceLabels, returnsObject } from '../core/choices'
import type { WorkbookChoice } from '../core/choices'
import { prepareFormatting } from '../core/formatters'
import { workbookCells, workbookIssue } from './workbook'
import type { WorkbookBody, WorkbookCell, WorkbookDefinition, WorkbookRegion } from './workbook'

interface WorkbookCellData {
  readonly definition: WorkbookCell
  readonly dataPath: string
  readonly value: TemplateValue
  readonly choice?: WorkbookChoice
}
/** Expanded workbook bodies retain their definitions, without translating node identities. */
export interface WorkbookData {
  readonly definition: WorkbookBody
  readonly path: string
  readonly iterations: Origin['iterations']
  readonly cells: readonly WorkbookCellData[]
  readonly regions: readonly { readonly definition: WorkbookRegion, readonly path: string, readonly instances: readonly WorkbookData[] }[]
}
export interface WorkbookDataOptions { readonly dictionaries?: Dictionaries, readonly purpose?: 'report' | 'issue' }
interface DataContext { readonly value: unknown, readonly path: string, readonly iterations: Origin['iterations'] }

function scalar(value: unknown): value is TemplateValue {
  return value === null || typeof value === 'string' || typeof value === 'boolean'
    || typeof value === 'number' && Number.isFinite(value)
}

/** Resolve bindings and region instances directly from the imported workbook definition. */
export function resolveWorkbookData(template: WorkbookDefinition, data: unknown, options: WorkbookDataOptions = {}): WorkbookData[] {
  const dictionaries = options.dictionaries ?? {}
  const purpose = options.purpose ?? 'report'
  const formatters = prepareWorkbookValues(template, dictionaries)
  if (!isDataObject(data)) {
    workbookIssue('invalid-data', '$template', 'root data must be an object', '$data', 'data')
  }
  const root: DataContext = { value: data, path: '$data', iterations: [] }
  const references = new Map<DataReference, ReturnType<typeof compileDataPath>>()
  const choiceOptions = createChoiceResolver()
  const choiceLabels = createChoiceLabels()
  const issues: TemplateIssue[] = []
  function reference(ref: DataReference, current: DataContext) {
    let read = references.get(ref)
    if (!read) {
      read = compileDataPath(ref.path)
      references.set(ref, read)
    }
    const source = ref.from === 'root' ? root : current
    return { value: read(source.value), path: `${source.path}.${ref.path}` }
  }
  function value(cell: WorkbookCell, context: DataContext): WorkbookCellData {
    const rules = cell.rules
    const format = formatters.get(cell.id)
    const result = 'path' in cell.value
      ? reference(cell.value, context)
      : { value: 'literal' in cell.value ? cell.value.literal : null, path: context.path }
    if (result.value === undefined && 'path' in cell.value && (cell.value.optional || purpose === 'issue')) {
      result.value = null
    }
    if (result.value === undefined) {
      workbookIssue('missing-source', cell.id, 'source is missing', result.path, 'data')
    }
    if (!scalar(result.value) && !(returnsObject(rules?.choice) && isDataObject(result.value))) {
      workbookIssue('non-scalar', cell.id, 'expected a scalar or a declared object choice', result.path, 'data')
    }
    if (rules?.list && purpose === 'report') {
      const issue = validateList(result.value, rules, dictionaries[rules.list] as readonly string[])
      if (issue) {
        issues.push({ ...issue, phase: 'data', path: result.path, nodeId: cell.id })
      }
    }
    let choice: WorkbookChoice | undefined
    if (rules?.choice) {
      let items
      try {
        const source = rules.choice.source
        items = choiceOptions(rules.choice, 'dictionary' in source ? dictionaries[source.dictionary] : reference(source, context).value ?? [])
      }
      catch (error) {
        workbookIssue('choice-source', cell.id, (error as Error).message, result.path, 'data')
      }
      if (!allowsChoiceInput(rules.choice, items)) {
        let labeled
        try {
          labeled = choiceLabels(items)
        }
        catch (error) {
          workbookIssue('choice-display', cell.id, (error as Error).message, result.path, 'data')
        }
        const blank = isBlank(result.value)
        const key = choiceKey(rules.choice, result.value)
        const displayKey = typeof key === 'string' || typeof key === 'number' ? key : String(result.value)
        const selected = labeled.find(item => item.key === displayKey)
        if (!blank && (key !== displayKey || !selected) && purpose === 'report') {
          issues.push({ phase: 'data', code: 'choice', path: result.path, nodeId: cell.id, message: 'select a key from the declared choice source' })
        }
        choice = { items: labeled, text: blank ? null : selected?.text ?? String(displayKey) }
      }
    }
    let resolved = scalar(result.value) ? result.value : choice!.text
    if (format) {
      try {
        resolved = format(resolved)
      }
      catch (error) {
        workbookIssue('format', cell.id, (error as Error).message, result.path, 'data')
      }
    }
    return { definition: cell, dataPath: result.path, value: resolved, ...(choice ? { choice } : {}) }
  }
  function expand(definition: WorkbookBody, context: DataContext): WorkbookData {
    const cells = definition.cells.map(cell => value(cell, context))
    const regions = (definition.regions ?? []).map(region => {
      const source = reference(region.source, context)
      if (region.type === 'scope') {
        if (source.value === undefined) {
          workbookIssue('missing-source', region.id, 'source is missing', source.path, 'data')
        }
        if (!isDataObject(source.value)) {
          workbookIssue('invalid-object', region.id, 'scope source must be an object', source.path, 'data')
        }
        return { definition: region, path: source.path, instances: [expand(region, { ...context, ...source })] }
      }
      const items = source.value ?? []
      if (!Array.isArray(items)) {
        workbookIssue('invalid-collection', region.id, 'repeat source must be an array of objects', source.path, 'data')
      }
      const instances = Array.from(items, (item, index) => {
        const path = `${source.path}[${index}]`
        if (!isDataObject(item)) {
          workbookIssue('invalid-item', region.id, 'repeat item must be an object', path, 'data')
        }
        return expand(region, { value: item, path, iterations: [...context.iterations, { nodeId: region.id, index }] })
      })
      return { definition: region, path: source.path, instances }
    })
    return { definition, path: context.path, iterations: context.iterations, cells, regions }
  }
  const sheets = template.sheets.map(sheet => expand(sheet, root))
  if (issues.length) {
    throw new TemplateError(issues)
  }
  return sheets
}

/** Check dependencies even when their fields belong to empty repeats. */
function prepareWorkbookValues(template: WorkbookDefinition, dictionaries: Dictionaries) {
  const formatters = new Map<string, ReturnType<typeof prepareFormatting>>()
  const lists = new Map<string, string[]>()
  const choiceSources = new Set<string>()
  for (const sheet of template.sheets) {
    for (const cell of workbookCells(sheet)) {
      const rules = cell.rules
      try {
        if (rules?.format) {
          formatters.set(cell.id, prepareFormatting(rules.format))
        }
      }
      catch (error) {
        workbookIssue('invalid-rules', cell.id, (error as Error).message)
      }
      if (rules?.list) {
        lists.set(rules.list, [...lists.get(rules.list) ?? [], cell.id])
      }
      if (rules?.choice && 'dictionary' in rules.choice.source) {
        choiceSources.add(rules.choice.source.dictionary)
      }
    }
  }
  const missing: TemplateIssue[] = []
  for (const [name, consumers] of lists) {
    const values = Object.hasOwn(dictionaries, name) ? dictionaries[name] : undefined
    if (!values?.length) {
      missing.push(...consumers.map(nodeId => ({ phase: 'data' as const, nodeId, path: `$dictionaries.${name}`,
        code: values ? 'empty-dictionary' : 'missing-dictionary', message: values ? `dictionary ${name} is empty` : `dictionary ${name} was not supplied` })))
    }
    else if (!values.every(value => typeof value === 'string')) {
      workbookIssue('invalid-dictionary', consumers[0], 'a string list requires string values', `$dictionaries.${name}`, 'data')
    }
  }
  for (const name of choiceSources) {
    if (!Object.hasOwn(dictionaries, name)) {
      workbookIssue('missing-dictionary', '$template', 'choice dictionary was not supplied', `$dictionaries.${name}`, 'data')
    }
  }
  if (missing.length) {
    throw new TemplateError(missing)
  }
  return formatters
}

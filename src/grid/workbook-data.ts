import { isDataObject } from '../core/json'
import { isBlank } from '../core/validation'
import { compileDataPath, readDataPath } from '../core/template'
import type { DataReference, Origin, TemplateValue } from '../core/template'
import type { Dictionaries } from '../core/dictionaries'
import { choiceKey, createChoiceResolver, returnsObject } from '../core/choices'
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
  const formatters = new Map<string, ReturnType<typeof prepareFormatting> | null>()
  if (!isDataObject(data)) {
    workbookIssue('invalid-data', '$template', 'root data must be an object', '$data', 'data')
  }
  const root: DataContext = { value: data, path: '$data', iterations: [] }
  const references = new Map<DataReference, ReturnType<typeof compileDataPath>>()
  const choiceOptions = createChoiceResolver()
  const warnings = new Set<string>()
  function warn(cell: WorkbookCell, path: string, message: string): void {
    const key = `${cell.id}:${message}`
    if (!warnings.has(key)) {
      console.warn(`sheetbind: ${cell.xlsx.address} (${path}): ${message}`)
      warnings.add(key)
    }
  }
  function reference(ref: DataReference, current: DataContext) {
    let read = references.get(ref)
    if (!read) {
      read = compileDataPath(ref.path)
      references.set(ref, read)
    }
    const source = ref.from === 'root' ? root : current
    return { value: read(source.value), path: `${source.path}.${ref.path}` }
  }
  function choices(cell: WorkbookCell, context: DataContext, path: string) {
    const rule = cell.rules!.choice!
    const source = rule.source
    const { items, problem } = choiceOptions(rule, 'dictionary' in source ? dictionaries[source.dictionary] : reference(source, context).value)
    if (problem || !items.length) {
      if (problem || !('dictionary' in source)) {
        warn(cell, path, `${problem ?? 'Choice source is missing or empty'}; its dropdown and lookup are skipped`)
      }
      return undefined
    }
    return items
  }
  function value(cell: WorkbookCell, context: DataContext): WorkbookCellData {
    const rules = cell.rules
    const result = 'path' in cell.value
      ? reference(cell.value, context)
      : { value: 'literal' in cell.value ? cell.value.literal : null, path: context.path }
    if (result.value === undefined) {
      if ('path' in cell.value && !cell.value.optional && purpose === 'report') {
        warn(cell, result.path, 'Source is missing; the cell is left blank')
      }
      result.value = null
    }
    if (!scalar(result.value) && !(returnsObject(rules?.choice) && isDataObject(result.value))) {
      workbookIssue('non-scalar', cell.id, 'expected a scalar or a declared object choice', result.path, 'data')
    }
    if (rules?.list && dictionaries[rules.list]?.some(item => typeof item !== 'string')) {
      warn(cell, result.path, `Dictionary ${rules.list} does not contain strings; its dropdown is skipped`)
    }
    let choice: WorkbookChoice | undefined
    if (rules?.choice) {
      const labeled = choices(cell, context, result.path)
      if (labeled) {
        const blank = isBlank(result.value)
        const key = choiceKey(rules.choice, result.value)
        const selected = labeled.find(item => item.key === key)
        const fallback = scalar(result.value) ? result.value : readDataPath(result.value, rules.choice.label)
        if (!blank && !selected) {
          warn(cell, result.path, 'Value has no choice label; the supplied value is retained')
        }
        if (!blank && !selected && !scalar(fallback)) {
          workbookIssue('non-scalar', cell.id, 'expected a scalar choice label', result.path, 'data')
        }
        if (purpose === 'issue' && !blank && !selected && labeled.some(item => item.text === String(fallback))) {
          workbookIssue('ambiguous-choice', cell.id, 'unmatched input equals another choice label and would read back as a different value', result.path, 'data')
        }
        choice = { items: labeled, text: blank ? null : selected?.text ?? fallback as TemplateValue }
      }
    }
    const selectedValue = scalar(result.value) ? result.value : choice?.text ?? readDataPath(result.value, rules!.choice!.label)
    if (!scalar(selectedValue)) {
      workbookIssue('non-scalar', cell.id, 'expected a scalar choice label', result.path, 'data')
    }
    let resolved = selectedValue
    if (rules?.format) {
      try {
        if (!formatters.has(cell.id)) {
          formatters.set(cell.id, null)
          formatters.set(cell.id, prepareFormatting(rules.format))
        }
        const format = formatters.get(cell.id)
        if (format) {
          resolved = format(resolved)
        }
      }
      catch (error) {
        warn(cell, result.path, `Formatting was skipped: ${error instanceof Error ? error.message : String(error)}; the original value is retained`)
      }
    }
    return { definition: cell, dataPath: result.path, value: resolved, ...(choice ? { choice } : {}) }
  }
  function expand(definition: WorkbookBody, context: DataContext): WorkbookData {
    const cells = definition.cells.map(cell => value(cell, context))
    const regions = (definition.regions ?? []).map(region => {
      const source = reference(region.source, context)
      if (region.type === 'scope') {
        if (source.value == null) {
          console.warn(`sheetbind: ${region.xlsx.address} (${source.path}): scope is missing; its fields are left blank`)
          source.value = {}
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
  // Shared sources can also feed formula ranges or rows inserted into an issued form.
  for (const sheet of template.sheets) {
    for (const cell of workbookCells(sheet)) {
      const source = cell.rules?.choice?.source
      if (source && ('dictionary' in source || source.from === 'root')) {
        choices(cell, root, 'dictionary' in source ? `$dictionaries.${source.dictionary}` : `$data.${source.path}`)
      }
    }
  }
  return sheets
}

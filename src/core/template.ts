import { isDataObject } from './json'
import { isBlank } from './validation'
import type { JsonValue } from './json'
import { isDataPath, compileDataPath } from './reference'
import type { DataReference } from './reference'
export type { DataReference } from './reference'
import { parseFieldRules, validateList } from './field-rules'
import type { FieldRules } from './field-rules'
import { parseDictionaries } from './dictionaries'
import type { Dictionaries } from './dictionaries'
import { allowsChoiceInput, choiceKey, createChoiceResolver, selectedChoice, returnsObject } from './choices'
import type { ResolvedChoice } from './choices'
import { prepareFormatting } from './formatters'

/** Template semantics, independent of coordinates, document formats and I/O. */
export type TemplateValue = string | number | boolean | null
export type FieldValue = TemplateValue | Readonly<Record<string, unknown>>
export type ValueExpression = { readonly literal: TemplateValue }
  | (DataReference & { readonly optional?: boolean })

/** Content belongs to the consumer. The core only composes it and resolves data. */
export type Fragment<Content> =
  | { readonly type: 'value', readonly id: string, readonly value: ValueExpression, readonly rules?: FieldRules }
  | { readonly type: 'group', readonly id: string, readonly content: Content, readonly children: readonly Fragment<Content>[] }
  | { readonly type: 'scope' | 'repeat', readonly id: string, readonly source: DataReference, readonly body: Fragment<Content> }

export interface Origin {
  readonly nodeId: string
  /** A concrete source path, not a stable record key for editable forms. */
  readonly dataPath: string
  readonly iterations: readonly { readonly nodeId: string, readonly index: number }[]
}
export type ResolvedFragment<Content> =
  | { readonly type: 'value', readonly origin: Origin, readonly value: FieldValue, readonly choice?: ResolvedChoice }
  | { readonly type: 'group', readonly origin: Origin, readonly content: Content, readonly children: readonly ResolvedFragment<Content>[] }
  | { readonly type: 'scope', readonly origin: Origin, readonly body: ResolvedFragment<Content> }
  | { readonly type: 'repeat', readonly origin: Origin, readonly instances: readonly ResolvedFragment<Content>[] }

export interface TemplateIssue {
  readonly rule?: string
  readonly args?: readonly JsonValue[]
  readonly index?: number
  readonly phase: 'template' | 'data'
  readonly code: string
  readonly path: string
  readonly nodeId: string
  readonly message: string
}
export class TemplateError extends Error {
  constructor(readonly issues: readonly TemplateIssue[]) {
    super(issues.map(issue => `${issue.path}: ${issue.message}`).join('\n'))
    this.name = 'TemplateError'
  }
}
export interface ExecutionOptions {
  readonly maxNodes?: number
  readonly maxDepth?: number
  readonly dictionaries?: Dictionaries
  readonly checkValues?: boolean
}

function scalar(value: unknown): value is TemplateValue {
  return value === null || typeof value === 'string' || typeof value === 'boolean'
    || typeof value === 'number' && Number.isFinite(value)
}

/** Typed, consumer-built templates; consumer content must support structuredClone. */
export function instantiate<Content>(template: Fragment<Content>, data: unknown, options: ExecutionOptions = {}): ResolvedFragment<Content> {
  const { fields, dictionaries, formatters } = prepareTemplate(template, options)
  const choiceOptions = createChoiceResolver()
  const { maxNodes } = options
  if (!isDataObject(data)) {
    fail('data', 'invalid-data', '$data', template.id, 'root data must be an object')
  }
  const root: DataContext = { value: data, path: '$data', iterations: [] }
  const references = new Map<DataReference, ReturnType<typeof compileDataPath>>()
  function resolveReference(reference: DataReference, current: DataContext) {
    let read = references.get(reference)
    if (!read) {
      read = compileDataPath(reference.path)
      references.set(reference, read)
    }
    const source = reference.from === 'root' ? root : current
    return { value: read(source.value), path: `${source.path}.${reference.path}` }
  }
  let count = 0
  const valueIssues: TemplateIssue[] = []
  function visit(node: Fragment<Content>, context: DataContext): ResolvedFragment<Content> {
    if (++count > (maxNodes ?? Infinity)) {
      fail('data', 'node-budget', context.path, node.id, 'execution exceeds maxNodes')
    }
    const origin: Origin = { nodeId: node.id, dataPath: context.path, iterations: context.iterations }
    switch (node.type) {
      case 'value': {
        const rules = fields.get(node.id)
        const result = 'literal' in node.value ? { value: node.value.literal, path: context.path } : resolveReference(node.value, context)
        if (result.value === undefined && 'path' in node.value && node.value.optional) {
          result.value = null
        }
        if (result.value === undefined) {
          fail('data', 'missing-source', result.path, node.id, 'source is missing')
        }
        if (!scalar(result.value) && !(returnsObject(rules?.choice) && isDataObject(result.value))) {
          return fail('data', 'non-scalar', result.path, node.id, 'expected a scalar or a declared object choice')
        }
        if (rules?.list && options.checkValues !== false) {
          const issue = validateList(result.value, rules, dictionaries[rules.list] as readonly string[])
          if (issue) {
            valueIssues.push({ ...issue, phase: 'data', path: result.path, nodeId: node.id })
          }
        }
        let choice: ResolvedChoice | undefined
        if (rules?.choice) {
          let items
          try {
            items = choiceOptions(rules.choice, data, context.value, dictionaries)
          }
          catch (error) {
            return fail('data', 'choice-source', result.path, node.id, (error as Error).message)
          }
          const selected = selectedChoice(rules.choice, items, result.value)
          const blank = isBlank(result.value)
          if (!blank && !selected && !allowsChoiceInput(rules.choice, items) && options.checkValues !== false) {
            valueIssues.push({ phase: 'data', code: 'choice', path: result.path, nodeId: node.id, message: 'select a key from the declared choice source' })
          }
          const key = choiceKey(rules.choice, result.value)
          if (!allowsChoiceInput(rules.choice, items)) {
            choice = { items, key: blank ? null : typeof key === 'string' || typeof key === 'number' ? key : String(result.value) }
          }
        }
        let value = scalar(result.value) ? result.value : structuredClone(result.value)
        const format = formatters.get(node.id)
        if (format && scalar(value)) {
          try {
            value = format(value)
          }
          catch (error) {
            return fail('data', 'format', result.path, node.id, (error as Error).message)
          }
        }
        return { type: 'value', origin: { ...origin, dataPath: result.path }, value, ...(choice ? { choice } : {}) }
      }
      case 'group': return { type: 'group', origin, content: structuredClone(node.content), children: node.children.map(child => visit(child, context)) }
      case 'scope': case 'repeat': {
        const source = resolveReference(node.source, context)
        if (node.type === 'scope') {
          if (source.value === undefined) {
            fail('data', 'missing-source', source.path, node.id, 'source is missing')
          }
          if (!isDataObject(source.value)) {
            fail('data', 'invalid-object', source.path, node.id, 'scope source must be an object')
          }
          return { type: 'scope', origin: { ...origin, dataPath: source.path }, body: visit(node.body, { ...context, ...source }) }
        }
        const items = source.value ?? []
        if (!Array.isArray(items)) {
          return fail('data', 'invalid-collection', source.path, node.id, 'repeat source must be an array of objects')
        }
        // Every instance costs at least one node when the caller sets a budget.
        if (maxNodes !== undefined && items.length > maxNodes - count) {
          fail('data', 'node-budget', source.path, node.id, 'collection exceeds remaining maxNodes')
        }
        const instances = Array.from(items, (item, index) => {
          const path = `${source.path}[${index}]`
          if (!isDataObject(item)) {
            fail('data', 'invalid-item', path, node.id, 'repeat item must be an object')
          }
          return visit(node.body, { value: item, path, iterations: [...context.iterations, { nodeId: node.id, index }] })
        })
        return { type: 'repeat', origin: { ...origin, dataPath: source.path }, instances }
      }
    }
  }
  const result = visit(template, root)
  if (valueIssues.length) {
    throw new TemplateError(valueIssues)
  }
  return result
}

interface DataContext { readonly value: unknown, readonly path: string, readonly iterations: Origin['iterations'] }

function fail(phase: TemplateIssue['phase'], code: string, path: string, nodeId: string, message: string): never {
  throw new TemplateError([{ phase, code, path, nodeId, message }])
}

/** Validate the complete definition, including empty repeats, before any data is executed. */
function prepareTemplate<Content>(template: Fragment<Content>, options: ExecutionOptions) {
  const fields = new Map<string, FieldRules>()
  const { maxNodes, maxDepth } = options
  for (const value of [maxNodes, maxDepth]) {
    if (value !== undefined && (!Number.isSafeInteger(value) || value < 1)) {
      throw new RangeError('execution limits must be positive safe integers')
    }
  }
  const ids = new Set<string>()
  const lists = new Map<string, string[]>()
  const formatters = new Map<string, ReturnType<typeof prepareFormatting>>()
  const choiceSources = new Set<string>()
  function reference(ref: DataReference, id: string): void {
    if (!isDataPath(ref.path) || ref.from !== undefined && ref.from !== 'current' && ref.from !== 'root') {
      fail('template', 'invalid-reference', id, id, 'expected a safe path and current/root context')
    }
  }
  function validate(node: Fragment<Content>, depth: number): void {
    if (maxDepth !== undefined && depth > maxDepth) {
      fail('template', 'depth-limit', node.id, node.id, 'template nesting exceeds maxDepth')
    }
    if (!node.id || ids.has(node.id)) {
      fail('template', 'duplicate-identity', node.id, node.id, 'node IDs must be nonempty and unique')
    }
    ids.add(node.id)
    if (maxNodes !== undefined && ids.size > maxNodes) {
      fail('template', 'node-budget', node.id, node.id, 'template exceeds maxNodes')
    }
    switch (node.type) {
      case 'value':
        if (node.rules !== undefined) {
          if ('literal' in node.value) {
            fail('template', 'invalid-rules', node.id, node.id, 'field rules require a data binding')
          }
          let rules: FieldRules
          try {
            rules = parseFieldRules(node.rules)
            fields.set(node.id, rules)
            if (rules.format) {
              formatters.set(node.id, prepareFormatting(rules.format))
            }
          }
          catch (error) {
            fail('template', 'invalid-rules', node.id, node.id, (error as Error).message)
          }
          if (rules.list) {
            const consumers = lists.get(rules.list) ?? []
            consumers.push(node.id)
            lists.set(rules.list, consumers)
          }
          if (rules.choice && 'dictionary' in rules.choice.source) {
            choiceSources.add(rules.choice.source.dictionary)
          }
        }
        if ('literal' in node.value) {
          if (!scalar(node.value.literal)) {
            fail('template', 'non-scalar', node.id, node.id, 'literal must be a finite scalar')
          }
        }
        else {
          reference(node.value, node.id)
        }
        break
      case 'group': node.children.forEach(child => validate(child, depth + 1))
        break
      case 'scope': case 'repeat': reference(node.source, node.id)
        validate(node.body, depth + 1)
        break
      default: fail('template', 'unknown-node', '$template', '', 'unknown template operation')
    }
  }
  validate(template, 1)
  let dictionaries: Dictionaries
  try {
    dictionaries = parseDictionaries(options.dictionaries ?? {})
  }
  catch (error) {
    return fail('data', 'invalid-dictionaries', '$dictionaries', template.id, (error as Error).message)
  }
  const missing: TemplateIssue[] = []
  for (const [name, fields] of lists) {
    const values = Object.hasOwn(dictionaries, name) ? dictionaries[name] : undefined
    if (!values?.length) {
      missing.push(...fields.map(nodeId => ({ phase: 'data' as const, nodeId, path: `$dictionaries.${name}`,
        code: values ? 'empty-dictionary' : 'missing-dictionary', message: values ? `dictionary ${name} is empty` : `dictionary ${name} was not supplied` })))
    }
    else if (!values.every(value => typeof value === 'string')) {
      fail('data', 'invalid-dictionary', `$dictionaries.${name}`, fields[0], 'a string list requires string values')
    }
  }
  for (const name of choiceSources) {
    if (!Object.hasOwn(dictionaries, name)) {
      fail('data', 'missing-dictionary', `$dictionaries.${name}`, template.id, 'choice dictionary was not supplied')
    }
  }
  if (missing.length) {
    throw new TemplateError(missing)
  }
  return { fields, dictionaries, formatters }
}

import { allowsChoiceInput, createChoiceResolver, returnsObject } from '../core/choices'
import { equalJson, isDataObject } from '../core/json'
import { validateList } from '../core/field-rules'
import { isBlank } from '../core/validation'
import { readDataPath } from '../core/reference'
import type { FieldValue, TemplateValue } from '../core/template'
import type { Dictionaries } from '../core/dictionaries'
import { createWorkbookChoiceDisplay } from '../grid/workbook-choice-display'
import type { PreparedWorkbookField, PreparedWorkbookForm, WorkbookFormIssue, WorkbookFormResult } from './workbook'
import { WorkbookFormInputError } from './workbook'
import { readWorkbookRows } from './workbook-rows'
import type { WorkbookFormRows } from './workbook-rows'
import { dataPath, readData, writeData } from './records'
import type { DataPath } from './records'

export interface WorkbookFormField extends PreparedWorkbookField {
  readonly path: DataPath
  readonly context: DataPath
  readonly raw: TemplateValue
  /** Original input text for label lookup, separate from a normalized free value. */
  readonly text?: string | null
  readonly issue?: Pick<WorkbookFormIssue, 'code' | 'message'>
  readonly location: { readonly sheetName: string, readonly address: string, readonly nodeId: string, readonly path: string }
}
export interface WorkbookFormSubmission {
  readonly data: Readonly<Record<string, unknown>>
  readonly rows: readonly WorkbookFormRows[]
  readonly fields: readonly WorkbookFormField[]
}
interface ReadOptions {
  readonly dictionaries: Dictionaries
  readonly context: Readonly<Record<string, unknown>>
  readonly local: Readonly<Record<string, unknown>>
  readonly validateText?: (text: string) => readonly { code: string, message: string }[]
}

/** Decode labels and omit empty input lines before validating the submitted data. */
export function readWorkbookFormFields(prepared: PreparedWorkbookForm, submission: WorkbookFormSubmission, options: ReadOptions): WorkbookFormResult {
  const decoded = decodeFormFields(prepared, submission, options)
  const rows = readWorkbookRows(decoded.data, submission.rows, [...decoded.issues.keys()].map(field => field.path))
  const { data } = rows
  const issues: WorkbookFormIssue[] = []
  const invalid = new Set([...decoded.issues.keys()].map(field => field.location.path))
  for (const field of submission.fields) {
    const path = rows.remapPath(field.path)
    if (!path) {
      continue
    }
    const location = { ...field.location, path: dataPath(path) }
    const decodedIssue = decoded.issues.get(field)
    if (decodedIssue) {
      issues.push({ ...decodedIssue, ...location })
    }
    // Failed decoding already explains this value; validating a placeholder adds false errors.
    if (invalid.has(field.location.path)) {
      continue
    }
    const context = field.reference.from === 'root' ? [] : rows.remapPath(field.context)
    if (!context) {
      continue
    }
    const rules = field.rules ?? {}
    const raw = readData(data, path)
    const issue = field.validate?.(raw, { root: data, current: readData(data, context) as Record<string, unknown>, path: location.path })
      ?? validateList(raw, rules, rules.list ? options.dictionaries[rules.list] as readonly string[] : undefined)
    if (issue) {
      issues.push({ phase: 'value', ...issue, ...location })
    }
    if (typeof raw === 'string') {
      for (const issue of options.validateText?.(raw) ?? []) {
        issues.push({ phase: 'value', ...issue, ...location })
      }
    }
  }
  if (issues.length) {
    return { success: false, issues, data }
  }
  return { success: true, data }
}

function decodeFormFields(prepared: PreparedWorkbookForm, submission: WorkbookFormSubmission, options: ReadOptions) {
  const data = structuredClone(submission.data) as Record<string, unknown>
  const issues = new Map<WorkbookFormField, WorkbookFormIssue>()
  const fields = new Map<string, FieldValue>()
  const choiceOptions = createChoiceResolver()
  const displayChoice = createWorkbookChoiceDisplay()
  const locations = new Map<string, WorkbookFormField['location']>()
  for (const field of submission.fields) {
    if (!locations.has(field.id)) {
      locations.set(field.id, field.location)
    }
  }
  function savedSource<T>(field: PreparedWorkbookField, read: () => T, location = locations.get(field.id)): T {
    try {
      return read()
    }
    catch (error) {
      if (!(error instanceof SyntaxError)) {
        throw error
      }
      throw new WorkbookFormInputError({ phase: 'xlsx', code: 'choice-source',
        ...location ?? { nodeId: field.id, path: '$workbook' },
        message: `saved choice source is missing or malformed: ${error.message}` })
    }
  }
  function resolve(field: PreparedWorkbookField, location?: WorkbookFormField['location']) {
    const rule = field.rules!.choice!
    const source = rule.source
    if ('dictionary' in source) {
      return choiceOptions(rule, options.dictionaries[source.dictionary])
    }
    if (source.from === 'root') {
      return choiceOptions(rule, readDataPath(options.context, source.path))
    }
    const stored = options.local[field.id]
    if (!location || !isDataObject(stored) || !Array.isArray(stored[location.path])) {
      throw new SyntaxError('local source was not issued for this field')
    }
    return choiceOptions(rule, stored[location.path])
  }
  // Validate the issued sources before interpreting user values, including blank fields.
  for (const field of prepared.fields.values()) {
    savedSource(field, () => {
      const list = field.rules?.list
      if (list && (!options.dictionaries[list]?.length || !options.dictionaries[list].every(item => typeof item === 'string'))) {
        throw new SyntaxError(`dictionary ${list} must contain strings`)
      }
      const source = field.rules?.choice?.source
      if (source && ('dictionary' in source || source.from === 'root')) {
        resolve(field)
      }
    })
  }
  for (const field of submission.fields) {
    const { location, rules } = field
    const items = rules?.choice ? savedSource(field, () => displayChoice({ key: null, items: resolve(field, location) }).items, location) : undefined
    if (field.issue) {
      issues.set(field, { phase: 'value', ...field.issue, ...location })
      continue
    }
    let raw: FieldValue = rules?.list ? field.text ?? field.raw : field.raw
    if (rules?.choice && items && !isBlank(raw)) {
      const selected = items.find(item => item.text === (field.text ?? raw))
      if (!selected && !allowsChoiceInput(rules.choice, items)) {
        raw = field.text ?? raw
        issues.set(field, { phase: 'value', code: 'choice', message: 'select a label from the declared choice source', ...location })
      }
      if (selected) {
        raw = returnsObject(rules.choice) ? structuredClone(selected.value) : selected.key
      }
    }
    if (fields.has(location.path) && !equalJson(fields.get(location.path), raw)) {
      issues.set(field, { phase: 'value', code: 'conflicting-field', message: 'all occurrences of the same data field must agree', ...location })
    }
    else {
      fields.set(location.path, raw)
      writeData(data, field.path, raw)
    }
  }
  return { data, issues }
}

import { createChoiceResolver, selectedChoice, returnsObject } from '../core/choices'
import { equalJson } from '../core/json'
import { validateList } from '../core/field-rules'
import { isBlank } from '../core/validation'
import type { FieldValue, TemplateValue } from '../core/template'
import { createWorkbookChoiceDisplay } from '../grid/workbook-choice-display'
import type { WorkbookChoiceOption } from '../grid/workbook-choice-display'
import type { PreparedWorkbookField, PreparedWorkbookForm, WorkbookFormIssue, WorkbookFormResult } from './workbook'
import { readWorkbookRows } from './workbook-rows'
import type { WorkbookFormRows } from './workbook-rows'
import { dataPath, readData, writeData } from './records'
import type { DataPath } from './records'
import type { WorkbookChoiceSources } from './workbook-choice-sources'

export interface WorkbookFormField extends PreparedWorkbookField {
  readonly path: DataPath
  readonly context: DataPath
  readonly raw: TemplateValue
  readonly issue?: Pick<WorkbookFormIssue, 'code' | 'message'>
  readonly location: { readonly sheetName: string, readonly address: string, readonly nodeId: string, readonly path: string }
}
export interface WorkbookFormSubmission {
  readonly data: Readonly<Record<string, unknown>>
  readonly rows: readonly WorkbookFormRows[]
  readonly fields: readonly WorkbookFormField[]
}
interface ReadOptions {
  readonly context?: Readonly<Record<string, unknown>>
  readonly objectSources?: WorkbookChoiceSources
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
    const choices = decoded.choices.get(field)
    if (rules.choice && !isBlank(raw) && choices) {
      const selected = selectedChoice(rules.choice, choices, raw)
      if (!selected || returnsObject(rules.choice) && !equalJson(raw, selected.value)) {
        issues.push({ phase: 'value', code: 'choice', message: 'select a value from the declared choice source', ...location })
      }
    }
    const issue = field.validate(raw, { root: data, current: readData(data, context) as Record<string, unknown>, path: location.path })
      ?? validateList(raw, rules, rules.list ? prepared.dictionaries[rules.list] as readonly string[] : undefined)
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
    return { success: false, issues }
  }
  return { success: true, data }
}

function decodeFormFields(prepared: PreparedWorkbookForm, submission: WorkbookFormSubmission, options: ReadOptions) {
  const { context, objectSources } = options
  const data = structuredClone(submission.data) as Record<string, unknown>
  const issues = new Map<WorkbookFormField, WorkbookFormIssue>()
  const fields = new Map<string, FieldValue>()
  const choiceOptions = createChoiceResolver()
  const displayChoice = createWorkbookChoiceDisplay()
  const choices = new Map<WorkbookFormField, readonly WorkbookChoiceOption[]>()
  for (const field of submission.fields) {
    const { location, rules } = field
    if (field.issue) {
      issues.set(field, { phase: 'value', ...field.issue, ...location })
      continue
    }
    let raw: FieldValue = field.raw
    if (rules?.choice && !isBlank(raw)) {
      try {
        const object = returnsObject(rules.choice)
        const source = object ? objectSources!.context : context ?? {}
        const dictionaries = object ? objectSources!.dictionaries : prepared.dictionaries
        const resolved = choiceOptions(rules.choice, source, readData(source, field.context), dictionaries)
        const items = displayChoice({ key: null, items: resolved }).items
        choices.set(field, items)
        const selected = items.find(item => item.text === raw)
        if (!selected) {
          issues.set(field, { phase: 'value', code: 'choice', message: 'select a label from the declared choice source', ...location })
          continue
        }
        raw = object ? structuredClone(selected.value) : selected.key
      }
      catch (error) {
        issues.set(field, { phase: 'value', code: 'choice-source', message: (error as Error).message, ...location })
        continue
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
  return { data, choices, issues }
}

import ExcelJS from 'exceljs'
import type { ValidationOptions } from '../core/validation'
import { parseDictionaries, warnMissingDictionaries } from '../core/dictionaries'
import type { Dictionaries } from '../core/dictionaries'
import { TemplateError, readDataPath } from '../core/template'
import { prepareWorkbookForm, WorkbookFormInputError } from './form-definition'
import type { PreparedWorkbookField, PreparedWorkbookForm, WorkbookFormIssue, WorkbookFormResult } from './form-definition'
import { issueWorkbookFormData, readWorkbookRows, referencePath, dataPath, readData, writeData } from './form-records'
import { serializeWorkbookChoiceSources, readWorkbookChoiceSources, workbookListSheetName } from './workbook-lists'
import type { DataPath, WorkbookFormRows } from './form-records'
import { placeWorkbook } from '../grid/workbook-layout'
import type { WorkbookPlacedSheet } from '../grid/workbook-layout'
import { WorkbookTemplate } from './workbook-template'
import { resolveWorkbookFormulas } from '../grid/workbook-formulas'
import { formatAddress, formatRange, parseRange } from '../grid/geometry'
import { loadWorkbook } from './workbook-input'
import { xlsxTextIssues } from './report-text'
import { readFormValue } from './form-value'
import { formCarrierDefinition, readFormStructure, placeFormMarkers, readFormMarkers } from './workbook-form-markers'
import { writeWorkbookPackage } from './workbook-source'
import { TaggedXlsxError, withTemplateLocations } from './tagged-template'
import { createChoiceResolver, returnsObject } from '../core/choices'
import { equalJson, isDataObject } from '../core/json'
import { validateList } from '../core/field-rules'
import { isBlank } from '../core/validation'
import type { FieldValue, TemplateValue } from '../core/template'
import { collectWorkbookDictionarySources } from '../grid/workbook'

/** Issue input fields and structural boundaries. Blank required fields can be completed later. */
export async function renderWorkbookForm(value: WorkbookTemplate, data: unknown, options: { readonly dictionaries?: Dictionaries } = {}): Promise<Buffer> {
  const { definition, source, resources } = WorkbookTemplate.content(value)
  const dictionaries = parseDictionaries(options.dictionaries, collectWorkbookDictionarySources(definition))
  const prepared = withTemplateLocations(definition, () => prepareWorkbookForm(definition, 'issue'))
  const issued = issueWorkbookFormData(prepared.template, data)
  const placed = withTemplateLocations(prepared.template, () => placeWorkbook(prepared.template, issued, { dictionaries, purpose: 'issue' }))
  const carrier = placeFormMarkers(placed, formCarrierDefinition(prepared.template))
  const plan = resolveWorkbookFormulas(carrier.plan, carrier.formulaRows)
  const choices = serializeWorkbookChoiceSources(plan, issued, dictionaries)
  return writeWorkbookPackage({ source, plan, dictionaries, resources, form: { markers: carrier.markers, rows: carrier.formulaRows, choices } })
}

/** Definition + completed form -> data/issues, using only the sources issued with the form. */
export async function readWorkbookForm(value: WorkbookTemplate, bytes: Uint8Array, options: ValidationOptions = {}): Promise<WorkbookFormResult> {
  const { definition } = WorkbookTemplate.content(value)
  const prepared = withTemplateLocations(definition, () => prepareWorkbookForm(definition, 'read', options))
  let returned: ExcelJS.Workbook
  try {
    returned = await loadWorkbook(bytes)
  }
  catch {
    return { success: false, issues: [{ phase: 'xlsx', code: 'invalid-workbook', path: '$workbook', message: 'cannot read a supported unencrypted XLSX form' }] }
  }
  try {
    const result = readFormSubmission(returned, prepared)
    if (!result.success) {
      return result
    }
    const sources = [...prepared.fields.values()].some(field => field.rules?.choice || field.rules?.list)
      ? readWorkbookChoiceSources(returned, workbookListSheetName(prepared.template.sheets.map(sheet => sheet.name)))
      : { dictionaries: {}, context: {}, local: {}, skipped: new Set<string>() }
    return readWorkbookFormFields(prepared, result, sources)
  }
  catch (error) {
    if (error instanceof WorkbookFormInputError) {
      return { success: false, issues: [error.issue] }
    }
    // Returned boundaries must still satisfy the shared layout contract.
    if (error instanceof TaggedXlsxError) {
      return { success: false, issues: error.issues.map(issue => ({ ...issue, phase: 'structure', code: 'form-layout' })) }
    }
    if (error instanceof RangeError || error instanceof TemplateError) {
      return { success: false, issues: [{ phase: 'structure', code: 'form-layout', path: '$workbook', message: error.message }] }
    }
    throw error
  }
}

type FormSubmissionResult = ({ readonly success: true } & WorkbookFormSubmission) | Extract<WorkbookFormResult, { success: false }>

function readFormSubmission(workbook: ExcelJS.Workbook, prepared: PreparedWorkbookForm): FormSubmissionResult {
  const definition = formCarrierDefinition(prepared.template)
  const issues: WorkbookFormIssue[] = []
  for (const name of definition.sheets.keys()) {
    if (!workbook.getWorksheet(name)) {
      issues.push(structureIssue('sheets', 'an input worksheet is missing or renamed', name))
    }
  }
  if (issues.length) {
    return { success: false, issues }
  }
  const markers = new Map([...definition.sheets].map(([name, column]) => [name, readFormMarkers(workbook.getWorksheet(name)!, column)]))
  const shape = withTemplateLocations(prepared.template, () => readFormStructure(prepared.template, definition, markers))
  const expected = placeFormMarkers(shape.plan, definition)
  const fields: WorkbookFormField[][] = []
  for (const { sheet: plan } of expected.plan.sheets) {
    const actual = markers.get(plan.name)!
    const wanted = expected.markers.get(plan.name)!
    if (JSON.stringify(actual) !== JSON.stringify(wanted)) {
      const index = actual.column !== wanted.column ? 0 : wanted.rows.findIndex((marker, index) => JSON.stringify(marker) !== JSON.stringify(actual.rows[index]))
      const found = actual.rows[index < 0 ? wanted.rows.length : index]
      const target = wanted.rows[index]
      const address = found ? formatAddress({ row: found.row, column: actual.column }) : undefined
      const expectedAddress = target ? formatAddress({ row: target.row, column: wanted.column }) : undefined
      issues.push(structureIssue('form-layout', `form boundary expected at ${expectedAddress ?? 'end of sheet'}, found ${address ?? 'none'}; copy or delete the whole record block`, plan.name, address ?? expectedAddress))
      continue
    }
    const result = readSheetFields(workbook.getWorksheet(plan.name)!, plan, prepared.fields, shape.contexts)
    fields.push(result.fields)
    for (const issue of result.issues) {
      issues.push(issue)
    }
  }
  return issues.length
    ? { success: false, issues }
    : { success: true, data: shape.data, rows: shape.rows, fields: fields.flat() }
}

function readSheetFields(sheet: ExcelJS.Worksheet, plan: WorkbookPlacedSheet, definitions: ReadonlyMap<string, PreparedWorkbookField>, contexts: ReadonlyMap<string, DataPath>) {
  const fields: WorkbookFormField[] = []
  const issues: WorkbookFormIssue[] = []
  const merges = new Map(sheet.model.merges.map(ref => {
    const range = parseRange(ref)
    return [formatAddress(range.start), formatRange(range)]
  }))
  for (const placed of plan.cells) {
    const field = definitions.get(placed.definitionId)
    if (!field) {
      continue
    }
    const address = formatAddress(placed.at)
    const cell = sheet.getCell(address)
    const expectedMerge = placed.size.rows > 1 || placed.size.columns > 1
      ? formatRange({ start: placed.at, end: { row: placed.at.row + placed.size.rows - 1, column: placed.at.column + placed.size.columns - 1 } })
      : undefined
    if (merges.get(cell.master.address) !== expectedMerge) {
      issues.push(structureIssue('merges', 'an input field has a different merged range', sheet.name, address))
      continue
    }
    const { raw, text } = readFormValue(cell)
    const location = { sheetName: sheet.name, address, nodeId: field.id, path: placed.origin.dataPath }
    const issue = raw === undefined
      ? cell.type === ExcelJS.ValueType.Formula
        ? { code: 'formula', message: 'recalculate and save the formula in Excel; its saved result must be a string, number, boolean, date or blank value' }
        : { code: 'non-scalar', message: 'enter a string, number, boolean, Excel date or blank value' }
      : undefined
    const context = contexts.get(placed.contextPath)!
    fields.push({ ...field, path: referencePath(field.reference, context), context, raw: raw ?? null, location,
      ...(field.rules?.list || field.rules?.choice ? { text } : {}), ...(issue ? { issue } : {}) })
  }
  return { fields, issues }
}

function structureIssue(code: string, message: string, sheetName: string, address?: string): WorkbookFormIssue {
  return { phase: 'structure', code, path: '$workbook', message, sheetName, address }
}
interface WorkbookFormField extends PreparedWorkbookField {
  readonly path: DataPath
  readonly context: DataPath
  readonly raw: TemplateValue
  /** Original input text for label lookup, separate from a normalized free value. */
  readonly text?: string | null
  readonly issue?: Pick<WorkbookFormIssue, 'code' | 'message'>
  readonly location: { readonly sheetName: string, readonly address: string, readonly nodeId: string, readonly path: string }
}
interface WorkbookFormSubmission {
  readonly data: Record<string, unknown>
  readonly rows: readonly WorkbookFormRows[]
  readonly fields: readonly WorkbookFormField[]
}
type ReadOptions = ReturnType<typeof readWorkbookChoiceSources>

/** Decode labels and omit empty input lines before validating the submitted data. */
function readWorkbookFormFields(prepared: PreparedWorkbookForm, submission: WorkbookFormSubmission, options: ReadOptions): WorkbookFormResult {
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
      for (const issue of xlsxTextIssues(raw)) {
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
  warnMissingDictionaries(collectWorkbookDictionarySources(prepared.template), options.dictionaries)
  const { data } = submission
  const issues = new Map<WorkbookFormField, WorkbookFormIssue>()
  const fields = new Map<string, FieldValue>()
  const choiceOptions = createChoiceResolver()
  const warnings = new Set<string>()
  function warn(field: PreparedWorkbookField, message = 'choice source is missing or empty'): void {
    if (!warnings.has(field.id)) {
      console.warn(`sheetbind: ${field.id}: ${message}; its lookup and list validation are skipped`)
      warnings.add(field.id)
    }
  }
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
    if (options.skipped.has(field.id)) {
      warn(field, 'choice source was unavailable when the form was issued')
      return []
    }
    let values: unknown
    if ('dictionary' in source) {
      values = options.dictionaries[source.dictionary]
    }
    else if (source.from === 'root') {
      values = readDataPath(options.context, source.path)
    }
    else {
      const stored = options.local[field.id]
      if (!location || !isDataObject(stored) || !Array.isArray(stored[location.path])) {
        throw new SyntaxError('local source was not issued for this field')
      }
      values = stored[location.path]
    }
    const { items, problem } = choiceOptions(rule, values)
    if (problem) {
      throw new SyntaxError(problem)
    }
    if (!items.length && !('dictionary' in source)) {
      warn(field)
    }
    return items
  }
  // Validate the issued sources before interpreting user values, including blank fields.
  for (const field of prepared.fields.values()) {
    savedSource(field, () => {
      const list = field.rules?.list
      if (list && options.dictionaries[list]?.some(item => typeof item !== 'string')) {
        warn(field, `dictionary ${list} does not contain strings`)
      }
      const source = field.rules?.choice?.source
      if (source && ('dictionary' in source || source.from === 'root')) {
        resolve(field)
      }
    })
  }
  for (const field of submission.fields) {
    const { location, rules } = field
    const items = rules?.choice ? savedSource(field, () => resolve(field, location), location) : undefined
    if (field.issue) {
      issues.set(field, { phase: 'value', ...field.issue, ...location })
      continue
    }
    const list = rules?.list && options.dictionaries[rules.list]
    let raw: FieldValue = list && list.length && list.every(item => typeof item === 'string') ? field.text ?? field.raw : field.raw
    if (rules?.choice && items?.length && !isBlank(raw)) {
      const selected = items.find(item => item.text === (field.text ?? raw))
      if (!selected) {
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

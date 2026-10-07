import { hasValidation } from '../core/field-rules'
import { readWorkbookFormFields } from '../form/workbook-read'
import type { WorkbookFormField, WorkbookFormSubmission } from '../form/workbook-read'
import ExcelJS from 'exceljs'
import type { ValidationOptions } from '../core/validation'
import { parseDictionaries } from '../core/dictionaries'
import type { Dictionaries } from '../core/dictionaries'
import { TemplateError } from '../core/template'
import { prepareWorkbookForm, referencePath, WorkbookFormInputError } from '../form/workbook'
import type { PreparedWorkbookField, PreparedWorkbookForm, WorkbookFormIssue, WorkbookFormResult } from '../form/workbook'
import { issueWorkbookFormData } from '../form/workbook-rows'
import { workbookChoiceSources } from '../form/workbook-choice-sources'
import type { DataPath } from '../form/records'
import { placeWorkbook } from '../grid/workbook-layout'
import type { WorkbookLayout } from '../grid/workbook-layout'
import { WorkbookTemplate } from './template'
import { resolveWorkbookFormulas } from '../grid/workbook-formulas'
import { createWorkbookOutput } from './workbook-output'
import { formatAddress, formatRange, parseRange } from './addresses'
import { loadWorkbook } from './workbook-input'
import { createWorkbookResources } from './workbook-resources'
import { workbookListSheetName } from './workbook-dropdowns'
import { readWorkbookChoiceSources, writeWorkbookChoiceSources } from './workbook-choice-sources'
import { xlsxTextIssues } from './report-text'
import { readFormValue } from './form-value'
import { formCarrierDefinition, formDataFromMarkers, placeFormMarkers, readFormMarkers, writeFormMarkers } from './workbook-form-markers'
import { writeWorkbookPackage } from './workbook-source'
import { TaggedXlsxError, withTemplateLocations } from './tagged-template'

/** Issue input fields and structural boundaries. Blank required fields can be completed later. */
export async function renderWorkbookForm(value: WorkbookTemplate, data: unknown, options: { readonly dictionaries?: Dictionaries } = {}): Promise<Buffer> {
  const { definition, source } = WorkbookTemplate.content(value)
  const dictionaries = parseDictionaries(options.dictionaries === undefined ? {} : options.dictionaries)
  const prepared = withTemplateLocations(definition, () => prepareWorkbookForm(definition, 'issue'))
  const issued = issueWorkbookFormData(prepared.template, data)
  const placed = withTemplateLocations(prepared.template, () => placeWorkbook(prepared.template, issued, { dictionaries, purpose: 'issue' }))
  const carrier = placeFormMarkers(placed, formCarrierDefinition(prepared.template))
  const plan = resolveWorkbookFormulas(carrier.plan, carrier.formulaRows)
  const sheets = plan.sheets.map(({ sheet }) => ({ ...sheet, cells: sheet.cells.map(cell => {
    const field = prepared.fields.get(cell.definitionId)
    return field && (hasValidation(field.rules, 'string') || cell.choice) ? { ...cell, text: true } : cell
  }) }))
  const resources = createWorkbookResources(WorkbookTemplate.content(value).resources)
  const workbook = createWorkbookOutput(plan, dictionaries, resources)
  for (const [name, markers] of carrier.markers) {
    writeFormMarkers(workbook.getWorksheet(name)!, markers)
  }
  const choices = workbookChoiceSources(prepared.template, issued, dictionaries, sheets)
  if (choices) {
    writeWorkbookChoiceSources(workbook, workbookListSheetName(prepared.template.sheets.map(sheet => sheet.name)), choices, resources)
  }
  return writeWorkbookPackage(workbook, { source, plan, sheets, carriers: carrier.formulaRows })
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
      : { dictionaries: {}, context: {}, local: {} }
    return readWorkbookFormFields(prepared, result.submission, { ...sources, validateText: xlsxTextIssues })
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

type FormSubmissionResult = { readonly success: true, readonly submission: WorkbookFormSubmission } | Extract<WorkbookFormResult, { success: false }>

function readFormSubmission(workbook: ExcelJS.Workbook, prepared: PreparedWorkbookForm): FormSubmissionResult {
  const definition = formCarrierDefinition(prepared.template)
  const issues: WorkbookFormIssue[] = []
  for (const name of definition.sheets) {
    if (!workbook.getWorksheet(name)) {
      issues.push(structureIssue('sheets', 'an input worksheet is missing or renamed', name))
    }
  }
  if (issues.length) {
    return { success: false, issues }
  }
  const markers = new Map([...definition.sheets].map(name => [name, readFormMarkers(workbook.getWorksheet(name)!)]))
  const shape = formDataFromMarkers(prepared.template, definition, markers)
  const plan = withTemplateLocations(prepared.template, () => placeWorkbook(prepared.template, shape.data, { purpose: 'read' }))
  const expected = placeFormMarkers(plan, definition)
  const fields: WorkbookFormField[][] = []
  for (const { sheet: plan } of expected.plan.sheets) {
    if (!definition.sheets.has(plan.name)) {
      continue
    }
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
    : { success: true, submission: { data: shape.data, rows: shape.rows, fields: fields.flat() } }
}

function readSheetFields(sheet: ExcelJS.Worksheet, plan: WorkbookLayout['sheets'][number], definitions: ReadonlyMap<string, PreparedWorkbookField>, contexts: ReadonlyMap<string, DataPath>) {
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

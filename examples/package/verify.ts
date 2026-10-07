import assert from 'node:assert/strict'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import ExcelJS from 'exceljs'
import * as api from 'sheetbind'
import sample from './template.data.json' with { type: 'json' }
import { definition as validationDefinition, data as validationData, options as validationOptions } from './validation.js'

const { data, dictionaries } = sample
const unchanged = structuredClone(sample)
function freeze(value: unknown): void {
  if (value && typeof value === 'object') {
    Object.values(value).forEach(freeze)
    Object.freeze(value)
  }
}
freeze(sample)
const directory = resolve(process.argv[2] ?? 'temp/package/verify')
await mkdir(directory, { recursive: true })
const source = await readFile(new URL('./template.xlsx', import.meta.url))
const template = await api.importWorkbookXlsx(source)
assert.deepEqual(api.workbookDictionarySources(template), ['statuses'])
const options = { dictionaries }
const layout = api.resolveWorkbook(template, data, options)
assert.throws(() => api.resolveWorkbook(template, { ...data, unused: new Date() }, options), SyntaxError)
assert(layout.sheets.every(sheet => !('xlsx' in sheet) && sheet.cells.every(cell => !('xlsx' in cell))))
const another = { ...data, customer: { name: 'Another customer' }, items: data.items.slice(0, 1) }
freeze(another)
const [output, form, anotherReport, anotherForm] = await Promise.all([
  api.renderWorkbookReport(template, data, options), api.renderWorkbookForm(template, data, options),
  api.renderWorkbookReport(template, another, options), api.renderWorkbookForm(template, another, options),
])
const separate = new ExcelJS.Workbook()
await separate.xlsx.load(Uint8Array.from(anotherReport).buffer)
assert.equal(separate.getWorksheet('Report')!.getCell('A1').value, 'Another customer')
assert.equal(separate.getWorksheet('Report')!.getCell('A6').value, 'Prepared for review')
assert.deepEqual(await api.readWorkbookForm(template, anotherForm), { success: true, data: another })
assert.deepEqual(api.resolveWorkbook(template, data, options), layout)
assert.deepEqual(sample, unchanged)
const runtimeData = { ...data, optional: undefined, items: data.items.map(item => ({ ...item, unused: { optional: undefined } })) }
const runtimeBefore = structuredClone(runtimeData)
assert.deepEqual(api.resolveWorkbook(template, runtimeData, options), layout)
assert.deepEqual(await api.readWorkbookForm(template, await api.renderWorkbookForm(template, runtimeData, options)), { success: true, data })
assert.deepEqual(runtimeData, runtimeBefore)
for (const empty of [{ customer: data.customer }, { customer: data.customer, items: null }]) {
  const before = structuredClone(empty)
  const report = new ExcelJS.Workbook()
  await report.xlsx.load(Uint8Array.from(await api.renderWorkbookReport(template, empty, options)).buffer)
  assert.equal(report.getWorksheet('Report')!.getCell('A5').value, 'Prepared for review')
  const form = await api.renderWorkbookForm(template, empty, options)
  assert.deepEqual(await api.readWorkbookForm(template, form), { success: true, data: { customer: data.customer, items: [] } })
  assert.deepEqual(empty, before)
}
await writeFile(resolve(directory, 'report.xlsx'), output)
const book = new ExcelJS.Workbook()
await book.xlsx.load(Uint8Array.from(output).buffer)
const report = book.getWorksheet('Report')!
assert.equal(report.getCell('A1').value, 'Sample customer')
assert.equal(report.getCell('A1').font.bold, true)
assert.equal(report.getCell('A1').font.size, 16)
assert.deepEqual([report.getCell('C4').value, report.getCell('C5').value], [0, 2.5])
assert.equal(report.getCell('A7').value, 'Prepared for review')
assert.equal(report.getCell('B4').dataValidation.type, 'list')
assert.equal(book.worksheets.filter(sheet => sheet.state === 'veryHidden').length, 1)

await writeFile(resolve(directory, 'form.xlsx'), form)
assert.deepEqual(await api.readWorkbookForm(template, await readFile(resolve(directory, 'form.xlsx'))), { success: true, data })
const completed = new ExcelJS.Workbook()
await completed.xlsx.load(Uint8Array.from(form).buffer)
completed.getWorksheet('Report')!.eachRow(row => row.eachCell(cell => {
  if (cell.value === 'Sample customer') {
    cell.value = 'Edited customer'
  }
}))
assert.deepEqual(await api.readWorkbookForm(template, Buffer.from(await completed.xlsx.writeBuffer())), {
  success: true, data: { ...data, customer: { name: 'Edited customer' } },
})

const invalidData = { ...data, items: [{ name: 'Bad', status: 'Done', hours: -1 }] }
await api.renderWorkbookReport(template, invalidData, options)
const invalidResult = await api.readWorkbookForm(template, await api.renderWorkbookForm(template, invalidData, options))
assert.equal(invalidResult.success, false)
if (!invalidResult.success) {
  assert(invalidResult.issues.some(issue => issue.path === '$data.items[0].hours' && issue.rule === 'min'))
}
await assert.rejects(api.renderWorkbookReport(template, undefined, options), SyntaxError)
await assert.rejects(api.renderWorkbookReport(template, data), api.TemplateError)
await assert.rejects(api.renderWorkbookReport(template, data, { dictionaries: { statuses: [5] } } as unknown as typeof options), SyntaxError)
console.log('Installed package: tagged XLSX, reports, edited forms and diagnostics: ok')

const choiceBook = new ExcelJS.Workbook()
choiceBook.addWorksheet('Input').addRows([
  ['{#questions}'], ['{.answer}{@choice:.answers; key=id; label=label; return=key; emptySource=input}'], ['{/questions}'],
])
const choiceTemplate = await api.importWorkbookXlsx(Buffer.from(await choiceBook.xlsx.writeBuffer()))
const choiceData = { questions: [{ answer: 'Omitted' }, { answer: 'Undefined', answers: undefined }, { answer: 'Null', answers: null }] }
assert.deepEqual(await api.readWorkbookForm(choiceTemplate, await api.renderWorkbookForm(choiceTemplate, choiceData)), {
  success: true, data: { questions: [{ answer: 'Omitted' }, { answer: 'Undefined' }, { answer: 'Null' }] },
})

const customForm = await api.renderWorkbookForm(validationDefinition, validationData)
assert.deepEqual(await api.readWorkbookForm(validationDefinition, customForm, validationOptions), { success: true, data: validationData })
const invalidCustomData = { ...validationData, items: [{ ...validationData.items[0], quantity: 1.234 }] }
await api.renderWorkbookReport(validationDefinition, invalidCustomData)
const invalidCustom = await api.renderWorkbookForm(validationDefinition, invalidCustomData)
await assert.rejects(api.readWorkbookForm(validationDefinition, invalidCustom), error => error instanceof api.TemplateError
  && error.issues.some(issue => issue.code === 'invalid-rules' && issue.message === 'Unknown validation rule: decimalPlaces'))
const customResult = await api.readWorkbookForm(validationDefinition, invalidCustom, validationOptions)
assert.equal(customResult.success, false)
if (!customResult.success) {
  assert(customResult.issues.some(issue => issue.rule === 'decimalPlaces' && issue.index === 3 && issue.message === 'Use at most 2 decimal places'))
}
console.log('Installed custom validation: typed handlers, messages and form read: ok')

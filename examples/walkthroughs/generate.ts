import assert from 'node:assert/strict'
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import ExcelJS from 'exceljs'
import { importWorkbookXlsx, renderWorkbookReport, renderWorkbookForm, readWorkbookForm } from 'sheetbind'
import budgetData from './budget.data.json' with { type: 'json' }
import studyData from './study-plan.data.json' with { type: 'json' }
import registrationData from './registration.data.json' with { type: 'json' }
import dictionaries from './registration-dictionaries.json' with { type: 'json' }

const directory = resolve(process.argv[2] ?? 'temp/walkthroughs')
await mkdir(directory, { recursive: true })
const file = (name: string) => resolve(directory, name)
const save = async (name: string, book: ExcelJS.Workbook) => writeFile(file(name), Buffer.from(await book.xlsx.writeBuffer()))
const json = async (name: string, value: unknown) => writeFile(file(name), `${JSON.stringify(value, null, 2)}\n`)
async function open(name: string): Promise<ExcelJS.Workbook> {
  const book = new ExcelJS.Workbook()
  await book.xlsx.load(Uint8Array.from(await readFile(file(name))).buffer)
  return book
}
function style(sheet: ExcelJS.Worksheet, widths: number[]): void {
  sheet.columns = widths.map(width => ({ width }))
  for (let row = 1; row <= sheet.rowCount; row++) {
    sheet.getRow(row).height = 25
    for (let column = 1; column <= widths.length; column++) {
      const cell = sheet.getCell(row, column)
      if (cell.value === null) {
        continue
      }
      cell.font = { name: 'Arial', size: 11, color: { argb: 'FF17352F' } }
      cell.alignment = { vertical: 'middle', wrapText: true }
    }
  }
  sheet.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
}
function band(sheet: ExcelJS.Worksheet, row: number, columns: number, color = 'FF087F6D'): void {
  for (let column = 1; column <= columns; column++) {
    const cell = sheet.getCell(row, column)
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } }
    cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } }
  }
}

const budgetBook = new ExcelJS.Workbook()
const budget = budgetBook.addWorksheet('Budget')
budget.addRows([
  ['Monthly budget'],
  [null, '{#months | axis=columns}'],
  ['Category', '{.name}', 'All months'],
  ['{#categories}'],
  [null, '{#.amounts}'],
  [null, null, '{#categories}'],
  ['{.name}', '{?.value}', { formula: 'IFERROR(SUM(B7:B7),0)' }],
  [null, null, '{/categories}'],
  [null, '{/.amounts}'],
  ['{/categories}'],
  ['Total', { formula: 'SUM(B3:B10)' }, { formula: 'SUM(C3:C10)' }],
  [null, '{/months}'],
])
style(budget, [20, 14, 20])
band(budget, 3, 3)
band(budget, 11, 3, 'FF2457A7')
budget.getRow(1).height = 42
budget.getRow(2).height = 42
budget.getRow(7).height = 40
for (const address of ['B7', 'C7', 'B11', 'C11']) {
  budget.getCell(address).numFmt = '#,##0.00'
}
await save('budget-template.xlsx', budgetBook)
await copyFile(new URL('./budget.data.json', import.meta.url), file('budget.data.json'))
const budgetTemplate = await importWorkbookXlsx(await readFile(file('budget-template.xlsx')))
for (const rows of [0, 1, 3]) {
  for (const columns of [0, 1, 3]) {
    const data = {
      categories: budgetData.categories.slice(0, rows),
      months: budgetData.months.slice(0, columns).map(month => ({ ...month, amounts: month.amounts.slice(0, rows) })),
    }
    const name = rows === 3 && columns === 3 ? 'budget.xlsx' : `budget-${rows}x${columns}.xlsx`
    await writeFile(file(name), await renderWorkbookReport(budgetTemplate, data))
    const sheet = (await open(name)).worksheets[0]
    assert.equal(sheet.getCell('A2').value, 'Category')
    assert.equal(sheet.getCell(3 + rows, 1).value, 'Total')
    assert.equal(sheet.getCell(2, 2 + columns).value, 'All months')
    for (let index = 0; index < rows; index++) {
      assert.equal(sheet.getCell(3 + index, 1).value, data.categories[index].name)
    }
    if (rows === 3 && columns === 3) {
      assert.equal(sheet.getCell('C4').value, 0)
      assert.equal(sheet.getCell('C5').value, null)
      assert.equal(sheet.getCell('E3').formula, 'IFERROR(SUM(B3:D3),0)')
      assert.equal(sheet.getCell('B6').formula, 'SUM(B2:B5)')
    }
  }
}

const studyBook = new ExcelJS.Workbook()
const study = studyBook.addWorksheet('Plan')
study.addRows([
  ['Learning plan', '{learner}'],
  ['{#courses}'],
  ['{.title}'],
  ['Lesson', 'Minutes', 'Note'],
  ['{#.lessons}'],
  ['{.name}', '{.minutes}', '{?.note}'],
  [null, null, '{/.lessons}'],
  ['Subtotal', { formula: 'SUBTOTAL(9,B4:B7)' }],
  ['{.note}'],
  [null, null, '{/courses}'],
  ['Total minutes', { formula: 'IFERROR(SUBTOTAL(9,B3:B10),0)' }],
])
style(study, [30, 18, 24])
study.mergeCells('A3:C3')
study.mergeCells('A9:C9')
study.mergeCells('B11:C11')
band(study, 3, 3)
band(study, 11, 3, 'FF2457A7')
study.getRow(6).height = 40
study.getRow(11).height = 42
study.getCell('A9').font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF576C66' } }
study.getColumn(2).numFmt = '0'
await save('study-plan-template.xlsx', studyBook)
await copyFile(new URL('./study-plan.data.json', import.meta.url), file('study-plan.data.json'))
const studyTemplate = await importWorkbookXlsx(await readFile(file('study-plan-template.xlsx')))
for (const count of [0, 1, 3]) {
  const name = count === 3 ? 'study-plan.xlsx' : `study-plan-${count}.xlsx`
  await writeFile(file(name), await renderWorkbookReport(studyTemplate, { ...studyData, courses: studyData.courses.slice(0, count) }))
  const sheet = (await open(name)).worksheets[0]
  assert.equal(sheet.getCell('B1').value, 'Alex')
  const titles = new Set(studyData.courses.map(course => course.title))
  let found = 0
  sheet.eachRow(row => {
    if (titles.has(String(row.getCell(1).value))) {
      found++
    }
  })
  assert.equal(found, count)
}

const registrationBook = new ExcelJS.Workbook()
const registration = registrationBook.addWorksheet('Registration')
registration.addRows([
  ['Event registration'],
  ['{#groups}'],
  ['Group', '{.name}\n{@validate:required|string}'],
  ['Participant', 'Ticket', 'Sessions', 'Note'],
  ['{#.participants}'],
  [
    '{.name}\n{@validate:required|string}',
    '{.ticket}\n{@choice:Tickets; key=id; label=label; return=key}\n{@validate:required|string}',
    '{.sessions}\n{@validate:required|number|min:0}',
    '{?.note}\n{@validate:string}',
  ],
  [null, null, null, '{/.participants}'],
  ['Contact', '{?.contact}\n{@validate:string}'],
  [null, null, null, '{/groups}'],
  ['Return the completed XLSX file.'],
])
style(registration, [25, 38, 20, 24])
registration.mergeCells('A1:D1')
registration.mergeCells('B3:D3')
registration.mergeCells('B8:D8')
registration.mergeCells('A10:D10')
band(registration, 1, 4)
band(registration, 4, 4, 'FF2457A7')
registration.getRow(3).height = 40
registration.getRow(6).height = 66
registration.getRow(8).height = 40
registration.getCell('C6').numFmt = '0'
for (const address of ['B3', 'A6', 'B6', 'C6', 'D6', 'B8']) {
  registration.getCell(address).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF5F1' } }
}
await save('registration-template.xlsx', registrationBook)
await copyFile(new URL('./registration.data.json', import.meta.url), file('registration.data.json'))
await copyFile(new URL('./registration-dictionaries.json', import.meta.url), file('registration-dictionaries.json'))
const registrationTemplate = await importWorkbookXlsx(await readFile(file('registration-template.xlsx')))
const options = { dictionaries }
const issued = await renderWorkbookForm(registrationTemplate, registrationData, options)
await writeFile(file('registration-issued.xlsx'), issued)
assert.deepEqual(await readWorkbookForm(registrationTemplate, issued, options), {
  success: true,
  data: { groups: [{ name: 'Friends', participants: [], contact: null }, { name: 'Family', participants: [], contact: null }] },
})
const completed = await open('registration-issued.xlsx')
const input = completed.worksheets[0]
const inputRows: number[] = []
input.eachRow({ includeEmpty: true }, row => {
  if (row.getCell(2).dataValidation?.type === 'list') {
    inputRows.push(row.number)
  }
})
assert.equal(inputRows.length, 4)
const entries = [
  ['Alex', 'Standard', 2, null],
  null,
  ['Sam', 'Student', 0, 'First visit'],
  ['Taylor', 'Standard', 1, null],
]
entries.forEach((values, index) => values?.forEach((value, column) => {
  input.getCell(inputRows[index], column + 1).value = value
}))
await save('registration-completed.xlsx', completed)
const expected = {
  groups: [
    { name: 'Friends', participants: [{ name: 'Alex', ticket: 'standard', sessions: 2, note: null }, { name: 'Sam', ticket: 'student', sessions: 0, note: 'First visit' }], contact: null },
    { name: 'Family', participants: [{ name: 'Taylor', ticket: 'standard', sessions: 1, note: null }], contact: null },
  ],
}
assert.deepEqual(await readWorkbookForm(registrationTemplate, await readFile(file('registration-completed.xlsx')), options), { success: true, data: expected })
await json('registration-completed.json', expected)
input.getCell(inputRows[2], 3).value = -1
await save('registration-invalid.xlsx', completed)
const invalid = await readWorkbookForm(registrationTemplate, await readFile(file('registration-invalid.xlsx')), options)
assert(!invalid.success)
assert.equal(invalid.issues.length, 1)
assert.equal(invalid.issues[0].code, 'min')
assert.equal(invalid.issues[0].path, '$data.groups[0].participants[1].sessions')
assert.equal(invalid.issues[0].address, `C${inputRows[2]}`)
await json('registration-invalid.json', invalid)
await json('registration-layout.json', { inputRows, rowCount: input.rowCount })
console.log(`Examples: ${directory}`)

import assert from 'node:assert/strict'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import ExcelJS from 'exceljs'
import { importWorkbookXlsx, renderWorkbookReport, renderWorkbookForm, readWorkbookForm } from 'sheetbind'

const directory = resolve(process.argv[2] ?? 'temp/tutorials')
await mkdir(directory, { recursive: true })
const save = async (name: string, book: ExcelJS.Workbook) => writeFile(resolve(directory, name), Buffer.from(await book.xlsx.writeBuffer()))
const open = async (name: string) => {
  const book = new ExcelJS.Workbook()
  await book.xlsx.load(Uint8Array.from(await readFile(resolve(directory, name))).buffer)
  return book
}

const firstReport = new ExcelJS.Workbook()
firstReport.addWorksheet('Report').addRow(['Customer', '{customer.name}'])
await save('first-report-template.xlsx', firstReport)
const firstReportTemplate = await importWorkbookXlsx(await readFile(resolve(directory, 'first-report-template.xlsx')))
await writeFile(resolve(directory, 'first-report.xlsx'), await renderWorkbookReport(firstReportTemplate, { customer: { name: 'Sample customer' } }))
assert.equal((await open('first-report.xlsx')).worksheets[0].getCell('B1').value, 'Sample customer')

const firstForm = new ExcelJS.Workbook()
firstForm.addWorksheet('Input').addRow(['Customer', '{customer.name}{@validate:required|string}'])
await save('first-form-template.xlsx', firstForm)
const firstFormTemplate = await importWorkbookXlsx(await readFile(resolve(directory, 'first-form-template.xlsx')))
await writeFile(resolve(directory, 'first-form-issued.xlsx'), await renderWorkbookForm(firstFormTemplate, { customer: { name: 'Sample customer' } }))
const firstCompleted = await open('first-form-issued.xlsx')
firstCompleted.worksheets[0].getCell('B2').value = 'Alex'
await save('first-form-completed.xlsx', firstCompleted)
assert.deepEqual(await readWorkbookForm(firstFormTemplate, await readFile(resolve(directory, 'first-form-completed.xlsx'))), {
  success: true, data: { customer: { name: 'Alex' } },
})
firstCompleted.worksheets[0].getCell('B2').value = null
await save('first-form-invalid.xlsx', firstCompleted)
const firstInvalid = await readWorkbookForm(firstFormTemplate, await readFile(resolve(directory, 'first-form-invalid.xlsx')))
assert.equal(firstInvalid.success, false)
if (!firstInvalid.success) {
  assert.equal(firstInvalid.issues[0].code, 'required')
  assert.equal(firstInvalid.issues[0].address, 'B2')
}

const rules = new ExcelJS.Workbook()
rules.addWorksheet('Input').addRow(['Quantity', '{quantity}{@validate:required|number|multipleOf:2}'])
await save('rules-template.xlsx', rules)

const listBook = new ExcelJS.Workbook()
listBook.addWorksheet('Input').addRow(['Status', '{status}{@list:Statuses}'])
await save('list-template.xlsx', listBook)
const choiceBook = new ExcelJS.Workbook()
choiceBook.addWorksheet('Input').addRow(['Product', '{product}{@choice:Products; key=id; label=name}'])
await save('choice-template.xlsx', choiceBook)

// Ordinary workbook -> tagged template. Only the selected cells and rows change.
const ordinary = new ExcelJS.Workbook()
const original = ordinary.addWorksheet('Order')
original.addRows([
  ['Customer', 'Sample customer'],
  ['Item', 'Quantity', 'Price', 'Amount', 'Note'],
  ['Paper', 2, 5, { formula: 'B3*C3' }, null],
  ['Total', null, null, { formula: 'SUM(D2:D3)' }],
])
original.columns = [{ width: 16 }, { width: 20 }, { width: 12 }, { width: 16 }, { width: 15 }]
for (let row = 1; row <= 4; row++) {
  original.getRow(row).height = 24
  for (let column = 1; column <= 5; column++) {
    const cell = original.getCell(row, column)
    cell.font = { name: 'Arial', size: 11, color: { argb: 'FF17352F' } }
    cell.alignment = { vertical: 'middle' }
  }
}
for (let column = 1; column <= 5; column++) {
  original.getCell(2, column).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF087F6D' } }
  original.getCell(2, column).font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } }
  original.getCell(3, column).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF5F1' } }
  original.getCell(4, column).font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF17352F' } }
}
for (const column of [3, 4]) {
  original.getColumn(column).numFmt = '0.00'
}
original.getCell('E3').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFCC' } }
original.getCell('E3').alignment = { vertical: 'middle', wrapText: true }
original.pageSetup.orientation = 'landscape'
await save('ordinary.xlsx', ordinary)

const tagged = await open('ordinary.xlsx')
const reportSheet = tagged.worksheets[0]
reportSheet.spliceRows(3, 0, ['{#items}'])
reportSheet.spliceRows(5, 0, [null, null, null, null, '{/items}'])
reportSheet.getCell('B1').value = '{customer.name}'
for (const [address, value] of Object.entries({ A4: '{.name}', B4: '{.quantity}', C4: '{.price}', E4: '{?.note}' })) {
  reportSheet.getCell(address).value = value
}
// ExcelJS row insertion does not rebase formulas; author their final template references explicitly.
reportSheet.getCell('D4').value = { formula: 'B4*C4' }
reportSheet.getCell('D6').value = { formula: 'SUM(D2:D5)' }
for (const row of [3, 5]) {
  reportSheet.getRow(row).height = 24
  for (let column = 1; column <= 5; column++) {
    const cell = reportSheet.getCell(row, column)
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF0FB' } }
    cell.font = { name: 'Arial', size: 11, color: { argb: 'FF2457A7' } }
    cell.alignment = { vertical: 'middle' }
  }
}
await save('report-template.xlsx', tagged)

const reportTemplate = await importWorkbookXlsx(await readFile(resolve(directory, 'report-template.xlsx')))
const items = [
  { name: 'Paper', quantity: 2, price: 5 },
  { name: 'Pen', quantity: 0, price: 3, note: 'Spare' },
  { name: 'Folder', quantity: 3, price: 4 },
]
for (const count of [0, 1, 3]) {
  const data = { customer: { name: 'Sample customer' }, items: items.slice(0, count) }
  await writeFile(resolve(directory, `report-${count}.json`), JSON.stringify(data, null, 2))
  await writeFile(resolve(directory, `report-${count}.xlsx`), await renderWorkbookReport(reportTemplate, data))
  const result = (await open(`report-${count}.xlsx`)).worksheets[0]
  assert.equal(result.getCell(count + 3, 1).value, 'Total')
  assert.equal(result.getCell(count + 3, 4).formula, `SUM(D2:D${count + 2})`)
  if (count) {
    assert.equal(result.getCell('D3').formula, 'B3*C3')
    assert.equal(result.getCell('C3').numFmt, '0.00')
    assert.deepEqual(result.getCell('E3').fill, { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFCC' } })
    assert.equal(result.getCell('E3').value, null)
  }
}

// Issue twenty empty lines. Persist this exact template and dictionary for the return=key reader.
const formBook = new ExcelJS.Workbook()
const formSheet = formBook.addWorksheet('Input')
formSheet.addRows([
  ['Product', 'Quantity', 'Date'],
  ['{#items}'],
  ['{.product}{@choice:Products; key=id; label=name; return=key}{@validate:required|string}', '{.quantity}{@validate:required|number|min:0}', '{?.date}{@validate:string}'],
  [null, null, '{/items}'],
])
formSheet.columns = [{ width: 24 }, { width: 16 }, { width: 18 }]
formSheet.getRow(1).font = { bold: true }
formSheet.getRow(3).height = 26
formSheet.getCell('B3').numFmt = '0.00'
formSheet.getCell('C3').numFmt = '@'
Reflect.set(formSheet.getCell('A3'), 'dataValidation', {
  type: 'any', formulae: [], showInputMessage: true, promptTitle: 'Product', prompt: 'Select a product from this form.',
})
const dictionaries = { Products: [{ id: '001', name: 'Paper' }, { id: '002', name: 'Paper' }] }
await save('form-template.xlsx', formBook)
await writeFile(resolve(directory, 'issued-dictionaries.json'), JSON.stringify(dictionaries, null, 2))
const formTemplate = await importWorkbookXlsx(await readFile(resolve(directory, 'form-template.xlsx')))
const issued = await renderWorkbookForm(formTemplate, { items: Array.from({ length: 20 }, () => ({})) }, { dictionaries })
await writeFile(resolve(directory, 'issued.xlsx'), issued)
assert.deepEqual(await readWorkbookForm(formTemplate, issued, { dictionaries }), { success: true, data: { items: [] } })

// Simulate two user edits separated by blank lines; read the separately saved upload.
const completed = await open('issued.xlsx')
const input = completed.worksheets[0]
input.getCell('A4').value = 'Paper [001]'
input.getCell('B4').value = 2
input.getCell('C4').value = '2026-01-15'
input.getCell('A6').value = 'Paper [002]'
input.getCell('B6').value = 0
await save('completed.xlsx', completed)
const expected = { items: [{ product: '001', quantity: 2, date: '2026-01-15' }, { product: '002', quantity: 0, date: null }] }
assert.deepEqual(await readWorkbookForm(formTemplate, await readFile(resolve(directory, 'completed.xlsx')), { dictionaries }), { success: true, data: expected })
await writeFile(resolve(directory, 'completed.json'), JSON.stringify(expected, null, 2))
input.getCell('A6').value = 'Unknown'
await save('invalid.xlsx', completed)
const invalid = await readWorkbookForm(formTemplate, await readFile(resolve(directory, 'invalid.xlsx')), { dictionaries })
assert.equal(invalid.success, false)
if (!invalid.success) {
  assert.equal(invalid.issues.length, 1)
  assert.equal(invalid.issues[0].path, '$data.items[1].product')
  assert.equal(invalid.issues[0].address, 'A6')
  assert.equal(invalid.issues[0].code, 'choice')
}
await writeFile(resolve(directory, 'invalid.json'), JSON.stringify(invalid, null, 2))

const contactsBook = new ExcelJS.Workbook()
contactsBook.addWorksheet('Contacts').addRows([
  ['{#contacts}'], ['Name', '{.name}'], ['Email', '{.email}'], [null, '{/contacts}'],
])
await save('contacts-template.xlsx', contactsBook)
const contactsTemplate = await importWorkbookXlsx(await readFile(resolve(directory, 'contacts-template.xlsx')))
const contacts = { contacts: [{ name: 'Alex', email: 'alex@example.com' }, { name: 'Sam', email: 'sam@example.com' }] }
const contactsForm = await renderWorkbookForm(contactsTemplate, contacts)
await writeFile(resolve(directory, 'contacts-issued.xlsx'), contactsForm)
assert.deepEqual(await readWorkbookForm(contactsTemplate, contactsForm), { success: true, data: contacts })
console.log(`Tutorial files: ${directory}`)

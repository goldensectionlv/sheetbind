import { writeFile } from 'node:fs/promises'
import ExcelJS from 'exceljs'
import { importWorkbookXlsx, readWorkbookForm, registerFormatter, registerValidationRule, renderWorkbookForm } from 'sheetbind'

registerFormatter('uppercase', value => value === null ? null : String(value).toUpperCase())
registerValidationRule('twoLetters', {
  validate: value => typeof value === 'string' && /^[A-Z]{2}$/.test(value),
  message: 'Enter two uppercase letters',
})

const source = new ExcelJS.Workbook()
source.addWorksheet('Input').addRow([
  '{date | format_date:DD.MM.YYYY}',
  '{enabled | bool_replace:"Yes","No"}',
  '{amount | float}',
  '{code | uppercase}{@validate:twoLetters}',
])
const template = await importWorkbookXlsx(Buffer.from(await source.xlsx.writeBuffer()))
const form = await renderWorkbookForm(template, { date: '2026-10-01', enabled: false, amount: '12.50', code: 'ab' })
await writeFile('formatting-form.xlsx', form)
console.log(JSON.stringify(await readWorkbookForm(template, form), null, 2))

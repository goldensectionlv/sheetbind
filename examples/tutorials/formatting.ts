import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, readWorkbookForm, registerFormatter, registerValidationRule, renderWorkbookForm } from 'sheetbind'

async function main() {
  registerFormatter('uppercase', value => value === null ? null : String(value).toUpperCase())
  registerValidationRule('twoLetters', {
    validate: value => typeof value === 'string' && /^[A-Z]{2}$/.test(value),
    message: 'Enter two uppercase letters',
  })

  const template = await importWorkbookXlsx(await readFile('formatting-template.xlsx'))
  const form = await renderWorkbookForm(template, { date: '2026-10-01', enabled: false, amount: '12.50', code: 'ab' })
  await writeFile('formatting-form.xlsx', form)
  console.log(JSON.stringify(await readWorkbookForm(template, form), null, 2))
}

main().catch(error => {
  console.error(error)
  process.exitCode = 1
})

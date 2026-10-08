import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookForm } from 'sheetbind'

async function main() {
  const template = await importWorkbookXlsx(await readFile('first-form-template.xlsx'))
  const data = { customer: { name: 'Sample customer' } }
  await writeFile('first-form-issued.xlsx', await renderWorkbookForm(template, data))
}

main().catch(error => {
  console.error(error)
  process.exitCode = 1
})

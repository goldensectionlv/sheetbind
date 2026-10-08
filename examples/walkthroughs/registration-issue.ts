import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookForm } from 'sheetbind'

async function main() {
  const template = await importWorkbookXlsx(await readFile('registration-template.xlsx'))
  const data = JSON.parse(await readFile('registration.data.json', 'utf8'))
  const dictionaries = JSON.parse(await readFile('registration-dictionaries.json', 'utf8'))
  await writeFile('registration-issued.xlsx', await renderWorkbookForm(template, data, { dictionaries }))
}

main().catch(error => {
  console.error(error)
  process.exitCode = 1
})

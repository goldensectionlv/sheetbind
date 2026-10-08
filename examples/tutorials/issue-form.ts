import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookForm } from 'sheetbind'

async function main() {
  const source = await readFile('form-template.xlsx')
  const template = await importWorkbookXlsx(source)
  const dictionaries = {
    Products: [{ id: '001', name: 'Paper' }, { id: '002', name: 'Paper' }],
  }
  const data = { items: Array.from({ length: 20 }, () => ({})) }
  const issued = await renderWorkbookForm(template, data, { dictionaries })

  await mkdir('saved-form', { recursive: true })
  await writeFile('saved-form/template.xlsx', source)
  await writeFile('saved-form/issued.xlsx', issued)
}

main().catch(error => {
  console.error(error)
  process.exitCode = 1
})

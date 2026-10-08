import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

async function main() {
  const template = await importWorkbookXlsx(await readFile('choice-template.xlsx'))
  const dictionaries = {
    Products: [
      { id: '001', name: 'Paper' },
      { id: '002', name: 'Pen' },
    ],
  }
  const data = { product: dictionaries.Products[0] }
  await writeFile('choice-report.xlsx', await renderWorkbookReport(template, data, { dictionaries }))
}

main().catch(error => {
  console.error(error)
  process.exitCode = 1
})

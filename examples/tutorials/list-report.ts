import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

async function main() {
  const template = await importWorkbookXlsx(await readFile('list-template.xlsx'))
  const data = { status: 'Draft' }
  const dictionaries = { Statuses: ['Draft', 'Ready'] }
  await writeFile('list-report.xlsx', await renderWorkbookReport(template, data, { dictionaries }))
}

main().catch(error => {
  console.error(error)
  process.exitCode = 1
})

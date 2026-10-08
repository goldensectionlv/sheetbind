import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

async function main() {
  const template = await importWorkbookXlsx(await readFile('first-report-template.xlsx'))
  const data = { customer: { name: 'Sample customer' } }
  await writeFile('report.xlsx', await renderWorkbookReport(template, data))
}

main().catch(error => {
  console.error(error)
  process.exitCode = 1
})

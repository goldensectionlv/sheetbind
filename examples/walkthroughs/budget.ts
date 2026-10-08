import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

async function main() {
  const template = await importWorkbookXlsx(await readFile('budget-template.xlsx'))
  const data = JSON.parse(await readFile('budget.data.json', 'utf8'))
  await writeFile('budget.xlsx', await renderWorkbookReport(template, data))
}

main().catch(error => {
  console.error(error)
  process.exitCode = 1
})

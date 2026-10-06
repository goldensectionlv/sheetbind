import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('study-plan-template.xlsx'))
const data = JSON.parse(await readFile('study-plan.data.json', 'utf8'))
await writeFile('study-plan.xlsx', await renderWorkbookReport(template, data))

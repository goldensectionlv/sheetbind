import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('report-template.xlsx'))
const data = JSON.parse(await readFile('report-3.json', 'utf8'))
await writeFile('report.xlsx', await renderWorkbookReport(template, data))

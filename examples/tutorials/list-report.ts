import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('list-template.xlsx'))
const data = { status: 'Draft' }
const dictionaries = { Statuses: ['Draft', 'Ready'] }
await writeFile('list-report.xlsx', await renderWorkbookReport(template, data, { dictionaries }))

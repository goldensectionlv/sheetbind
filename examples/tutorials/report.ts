import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('report-template.xlsx'))
const items = [
  { name: 'Paper', quantity: 2, price: 5 },
  { name: 'Pen', quantity: 0, price: 3, note: 'Spare' },
  { name: 'Folder', quantity: 3, price: 4 },
]

for (const count of [0, 1, 3]) {
  const data = { customer: { name: 'Sample customer' }, items: items.slice(0, count) }
  await writeFile(`report-${count}.xlsx`, await renderWorkbookReport(template, data))
}

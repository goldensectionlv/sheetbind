import { readFile } from 'node:fs/promises'
import { importWorkbookXlsx } from 'sheetbind'

type Example = 'invoice' | 'service' | 'comparison'

export async function loadTemplate(name: Example) {
  return importWorkbookXlsx(await readFile(new URL(`./${name}-tags.xlsx`, import.meta.url)))
}

export const invoiceData = {
  customer: 'Sample customer',
  items: [
    { id: '0001', description: 'Inspection', quantity: 2, price: 10 },
    { id: '0002', description: 'Adjustment', quantity: 0, price: 15 },
    { id: '0003', description: 'Review\nwith notes', quantity: 3, price: 12 },
  ],
}
export const serviceData = { sites: [
  { id: 'N-01', name: 'North site', note: 'No work scheduled', work: [] },
  { id: 'S-02', name: 'South site', note: 'Check access first', work: [
    { code: '001', description: 'Inspection', hours: 0, rate: 10 },
    { code: '002', description: 'Adjustment', hours: 2.5, rate: 20 },
    { code: '003', description: 'Review', hours: 1, rate: 15 },
  ] },
] }
export const comparisonData = { offers: [
  { name: 'Option A', price: 10, days: 2 }, { name: 'Option B', price: 12, days: 0 }, { name: 'Option C', price: 9, days: 5 },
] }

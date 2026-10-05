import { readFile } from 'node:fs/promises'
import { importWorkbookXlsx } from 'sheetbind'
import sample from './nested.data.json' with { type: 'json' }

export const definition = await importWorkbookXlsx(await readFile(new URL('./nested.xlsx', import.meta.url)))
export const data = sample.data

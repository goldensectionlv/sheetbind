import { readFile } from 'node:fs/promises'
import { importWorkbookXlsx } from 'sheetbind'
import sample from './template.data.json' with { type: 'json' }

export const definition = await importWorkbookXlsx(await readFile(new URL('./template.xlsx', import.meta.url)))
export const data = sample.data
export const dictionaries = sample.dictionaries

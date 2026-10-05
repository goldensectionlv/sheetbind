import { readFile } from 'node:fs/promises'
import { importWorkbookXlsx } from 'sheetbind'
import sample from './template.data.json' with { type: 'json' }

export const definition = await importWorkbookXlsx(await readFile(new URL('./template.xlsx', import.meta.url)))
export const data = sample.data
export const dictionaries = sample.dictionaries
export const declaredData = sample.declaredData
export const objectData = sample.objectData
export const objectResult = sample.objectResult
export const objectDefinition = await importWorkbookXlsx(await readFile(new URL('./object-template.xlsx', import.meta.url)))

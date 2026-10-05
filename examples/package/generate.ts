import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'
import sample from './template.data.json' with { type: 'json' }

const { data, dictionaries } = sample
const directory = resolve(process.argv[2] ?? 'temp/package')
await mkdir(directory, { recursive: true })
const template = await importWorkbookXlsx(await readFile(new URL('./template.xlsx', import.meta.url)))
await writeFile(resolve(directory, 'report.xlsx'), await renderWorkbookReport(template, data, { dictionaries }))
console.log(`Report output: ${directory}`)

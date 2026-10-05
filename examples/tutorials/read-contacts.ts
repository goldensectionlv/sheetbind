import { readFile } from 'node:fs/promises'
import { importWorkbookXlsx, readWorkbookForm } from 'sheetbind'

try {
  const template = await importWorkbookXlsx(await readFile('contacts-template.xlsx'))
  const result = await readWorkbookForm(template, await readFile('contacts-completed.xlsx'))
  if (result.success) {
    console.log(JSON.stringify(result.data, null, 2))
  }
  else {
    console.error(result.issues)
    process.exitCode = 1
  }
}
catch (error) {
  console.error('Cannot process the form:', error)
  process.exitCode = 1
}

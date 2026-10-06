import { readFile } from 'node:fs/promises'
import { importWorkbookXlsx, readWorkbookForm } from 'sheetbind'

try {
  const template = await importWorkbookXlsx(await readFile('registration-template.xlsx'))
  const dictionaries = JSON.parse(await readFile('registration-dictionaries.json', 'utf8'))
  const result = await readWorkbookForm(template, await readFile('registration-completed.xlsx'), { dictionaries })
  if (result.success) {
    console.log(JSON.stringify(result.data, null, 2))
  }
  else {
    console.error(JSON.stringify(result.issues, null, 2))
    process.exitCode = 1
  }
}
catch (error) {
  console.error('Cannot read the registration form:', error)
  process.exitCode = 1
}

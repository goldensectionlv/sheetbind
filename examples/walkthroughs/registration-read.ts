import { readFile } from 'node:fs/promises'
import { importWorkbookXlsx, readWorkbookForm } from 'sheetbind'

async function main() {
  const template = await importWorkbookXlsx(await readFile('registration-template.xlsx'))
  const result = await readWorkbookForm(template, await readFile('registration-completed.xlsx'))
  if (result.success) {
    console.log(JSON.stringify(result.data, null, 2))
  }
  else {
    console.error(JSON.stringify(result.issues, null, 2))
    process.exitCode = 1
  }
}

main().catch(error => {
  console.error('Cannot read the registration form:', error)
  process.exitCode = 1
})

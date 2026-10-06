import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { importWorkbookXlsx, readWorkbookForm, TemplateError, TaggedXlsxError, ValidationExecutionError } from 'sheetbind'

const directory = resolve(process.argv[2] ?? 'temp/tutorials')
const uploaded = resolve(process.argv[3] ?? resolve(directory, 'completed.xlsx'))
try {
  const template = await importWorkbookXlsx(await readFile(resolve(directory, 'form-template.xlsx')))
  const result = await readWorkbookForm(template, await readFile(uploaded))
  if (result.success) {
    console.log(JSON.stringify(result.data, null, 2))
  }
  else {
    console.error(JSON.stringify(result.issues, null, 2))
    process.exitCode = 1
  }
}
catch (error) {
  if (error instanceof TaggedXlsxError || error instanceof TemplateError) {
    console.error('Template or issuance configuration:', error.issues)
  }
  else if (error instanceof ValidationExecutionError) {
    console.error('Validation handler failed:', error.message)
  }
  else {
    console.error('Cannot process the file:', error)
  }
  process.exitCode = 1
}

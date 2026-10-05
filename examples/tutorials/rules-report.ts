import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'
import type { ValidationOptions } from 'sheetbind'

const options: ValidationOptions = {
  validationRules: {
    multipleOf: {
      validateArgs: args => args.length === 1 && typeof args[0] === 'number'
        && Number.isSafeInteger(args[0]) && args[0] > 0,
      validate(value, [step]) {
        return typeof value === 'number' && Number.isSafeInteger(value)
          && typeof step === 'number' && value % step === 0
      },
      message: ({ args }) => `Enter a whole number divisible by ${args[0]}`,
    },
  },
}
const template = await importWorkbookXlsx(await readFile('rules-template.xlsx'))
const data = { quantity: 4 }
await writeFile('rules-report.xlsx', await renderWorkbookReport(template, data, options))

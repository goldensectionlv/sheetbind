# Sheetbind

[English](./README.md) | [Русский](./README.ru.md)

Sheetbind fills XLSX templates with data. Add tags to cells in Excel, pass the workbook and data to your code, and save a report. You can also issue a form, let a user fill it in, and read the completed file.

Requires Node.js 22.13 or newer and ExcelJS 4.4.0. Licensed under MIT.

## First report

Create `first-report-template.xlsx`: put `Customer` in A1 and `{customer.name}` in B1. Set fonts, colours and column widths in Excel.

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('first-report-template.xlsx'))
const data = { customer: { name: 'Sample customer' } }
await writeFile('report.xlsx', await renderWorkbookReport(template, data))
```

The result contains `Sample customer` in B1. Follow [First report](./docs/getting-started.md) for package installation, a ready-made template and the command to run this code.

## Guide

- [Templates](./docs/templates.md): values, repeated rows, nested data and repeated columns.
- [Validation and lists](./docs/fields.md): field rules and dropdown choices.
- [Filling in a form](./docs/forms.md): issue an XLSX, edit it and read the returned data.
- [API](./docs/api.md): functions, options, results and errors.
- [Excel and limitations](./docs/xlsx.md): formatting, formulas and supported workbook features.

Keep the tagged XLSX as your template. After editing it, load the file again with `importWorkbookXlsx`. Formula calculation belongs to Excel; Sheetbind does not calculate formula results.

To work on the library itself, see [Development](./docs/development.md) and [Architecture](./docs/architecture.md).

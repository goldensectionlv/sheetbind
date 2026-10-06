# Sheetbind

[English](./README.md) | [Русский](./README.ru.md)

Sheetbind fills XLSX templates with data. Add tags to cells in Excel, pass the workbook and data to your code, and save a report. You can also issue a form, let a user fill it in, and read the completed file.

Requires Node.js 22.13 or newer and ExcelJS 4.4.0. Licensed under MIT.

[Read the documentation](https://goldensectionlv.github.io/sheetbind/) for illustrated examples and downloadable workbooks.

The package is not published on npm yet. The [installation guide](https://goldensectionlv.github.io/sheetbind/getting-started) shows how to build and install an archive from source.

## First report

Create `first-report-template.xlsx`: put `Customer` in A1 and `{customer.name}` in B1. Set fonts, colours and column widths in Excel.

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('first-report-template.xlsx'))
const data = { customer: { name: 'Sample customer' } }
await writeFile('report.xlsx', await renderWorkbookReport(template, data))
```

The result contains `Sample customer` in B1. Follow [First report](https://goldensectionlv.github.io/sheetbind/getting-started) for package installation, a ready-made template and the command to run this code.

## Guide

- [Templates](https://goldensectionlv.github.io/sheetbind/templates): values, repeated rows, nested data and repeated columns.
- [Validation and lists](https://goldensectionlv.github.io/sheetbind/fields): field rules and dropdown choices.
- [Filling in a form](https://goldensectionlv.github.io/sheetbind/forms): issue an XLSX, edit it and read the returned data.
- [API](https://goldensectionlv.github.io/sheetbind/api): functions, options, results and errors.
- [Excel and limitations](https://goldensectionlv.github.io/sheetbind/xlsx): formatting, formulas and supported workbook features.

Keep the tagged XLSX as your template. After editing it, load the file again with `importWorkbookXlsx`. Formula calculation belongs to Excel; Sheetbind does not calculate formula results.

To work on the library itself, see [Development](https://goldensectionlv.github.io/sheetbind/development) and [Architecture](https://goldensectionlv.github.io/sheetbind/architecture).

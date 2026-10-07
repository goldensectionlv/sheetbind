# First report

Create an XLSX with a customer name. You need Node.js 22.13 or newer, [pnpm 11.5.2](https://pnpm.io/installation), and Excel to create or open the workbook. Check `pnpm --version`: the commands below use pnpm 11.5.2.

To install that pnpm version, run `npm install --global pnpm@11.5.2`.

## 1. Install the package

These pages and their downloadable files describe unreleased changes. The published npm package `0.2.0` does not match them. Build an archive from the same Sheetbind checkout as the documentation. From the repository root:

```sh
pnpm install --frozen-lockfile
pnpm pack --out ./temp/sheetbind.tgz
```

Create a `sheetbind-example` directory outside the Sheetbind checkout and copy `temp/sheetbind.tgz` into it. From that new directory, install the archive, ExcelJS and a TypeScript runner:

```sh
pnpm init --init-type module
pnpm add ./sheetbind.tgz exceljs@4.4.0
pnpm add -D tsx
```

`--init-type module` adds `"type": "module"` to `package.json`; the examples use `import` and top-level `await`. For an existing application, install the dependencies there and use its module setup. ExcelJS is a peer dependency and must be installed alongside Sheetbind.

Run the rest of this guide from `sheetbind-example`. Keep form issuance and reading on the same package build.

## 2. Create the template

Download [first-report-template.xlsx](/examples/tutorials/first-report-template.xlsx) into the application directory, or create the workbook in Excel:

| Cell | Content |
| --- | --- |
| A1 | `Customer` |
| B1 | `{customer.name}` |

Save it as `first-report-template.xlsx`. You can set the font, borders and column widths in Excel.

## 3. Pass data and save the report

Save this code as `report.ts` beside the template. You can also [download the script](/examples/tutorials/first-report.ts) and save it under that name.

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('first-report-template.xlsx'))
const data = { customer: { name: 'Sample customer' } }
await writeFile('report.xlsx', await renderWorkbookReport(template, data))
```

Run it:

```sh
pnpm exec tsx report.ts
```

Open `report.xlsx` in the same directory. A1 still contains `Customer`; B1 now contains `Sample customer`.

If rendering warns that `$data.customer.name` is missing, check the data structure. `{ name: 'Sample customer' }` does not match `{customer.name}`; rendering leaves that cell blank. The example above supplies the required `customer` object.

Keep the tagged file as your template. After changing its tags or formatting, save it and call `importWorkbookXlsx` again.

## Next step

Continue with [Templates](./templates.md) to add a table, repeated rows and a total. If users need to fill in and return the file, follow [Filling in a form](./forms.md).

# First report

Create an XLSX with a customer name. Sheetbind runs on Node.js 22.13 or newer. Use Excel to edit the template or view the result.

## 1. Install the package

From your application directory:

```sh
npm install sheetbind
```

Dependencies, including ExcelJS, are installed automatically. TypeScript declarations are included. Use your application's existing module and build settings:

::: code-group

```ts [TypeScript / ESM]
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'
```

```js [CommonJS]
const { importWorkbookXlsx, renderWorkbookReport } = require('sheetbind')
```

:::

Keep form issuance and reading on the same package build. To try the downloadable scripts in a separate directory, use the optional [example setup](#run-examples).

## 2. Create the template

Download [first-report-template.xlsx](/examples/tutorials/first-report-template.xlsx) into the application directory, or create the workbook in Excel:

| Cell | Content |
| --- | --- |
| A1 | `Customer` |
| B1 | `{customer.name}` |

Save it as `first-report-template.xlsx`. You can set the font, borders and column widths in Excel.

## 3. Pass data and save the report

Save this code as `report.ts` beside the template. You can also [download the script](/examples/tutorials/first-report.ts) and save it under that name.

<<< @/public/examples/tutorials/first-report.ts

Run it with your application's usual command. File paths in this example are relative to the working directory. For a standalone run, follow the [example setup](#run-examples) below.

Open `report.xlsx` in the same directory. A1 still contains `Customer`; B1 now contains `Sample customer`.

If rendering warns that `$data.customer.name` is missing, check the data structure. `{ name: 'Sample customer' }` does not match `{customer.name}`; rendering leaves that cell blank. The example above supplies the required `customer` object.

Keep the tagged file as your template. After changing its tags or formatting, save it and call `importWorkbookXlsx` again.

## Run the downloadable examples {#run-examples}

This setup is for trying the guide outside an existing application. Create a directory and install a TypeScript runner for the `.ts` example files:

```sh
mkdir sheetbind-example
cd sheetbind-example
npm init -y
npm install sheetbind
npm install --save-dev tsx
```

Save the template and `report.ts` here, then run:

```sh
npx tsx report.ts
```

The other downloadable examples use the same setup. In an existing application, use its own TypeScript tools; `tsx` is only a runner for these standalone scripts.

## Next step

Continue with [Templates](./templates.md) to add a table, repeated rows and a total. If users need to fill in and return the file, follow [Filling in a form](./forms.md).

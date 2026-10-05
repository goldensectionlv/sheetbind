# Filling a form

A form is an XLSX file that a user fills in and returns to the application. `renderWorkbookForm` creates it; `readWorkbookForm` reads it against the original template. First [install the package and set up TypeScript execution](./getting-started.md) to follow these examples.

## Your first form

### 1. Create the template

Create an `Input` sheet in `first-form-template.xlsx`, or download the [template](/examples/tutorials/first-form-template.xlsx):

| Cell | Content |
| --- | --- |
| A1 | `Customer` |
| B1 | `{customer.name}{@validate:required\|string}` |

The field accepts nonempty text. Ordinary captions such as `Customer` are not included in the returned data.

### 2. Issue the file

Save [first-form.ts](/examples/tutorials/first-form.ts) next to the template:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookForm } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('first-form-template.xlsx'))
const data = { customer: { name: 'Sample customer' } }
await writeFile('first-form-issued.xlsx', await renderWorkbookForm(template, data))
```

Run from that directory:

```sh
pnpm exec tsx first-form.ts
```

Open `first-form-issued.xlsx`. The field is in B2: the form has added a hidden control row before the content. Change the name to `Alex` and save a **separate file named `completed.xlsx`** next to the script. Keep the original template for reading.

### 3. Read the completed file

Save [read-first-form.ts](/examples/tutorials/read-first-form.ts) in the same directory:

```ts
import { readFile } from 'node:fs/promises'
import { importWorkbookXlsx, readWorkbookForm } from 'sheetbind'

try {
  const template = await importWorkbookXlsx(await readFile('first-form-template.xlsx'))
  const result = await readWorkbookForm(template, await readFile('completed.xlsx'))
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
```

```sh
pnpm exec tsx read-first-form.ts
```

Output:

```json
{
  "customer": { "name": "Alex" }
}
```

Clear B2 and save the file again to get `success: false`. The `issues` array contains a `required` issue with `sheetName: "Input"`, `address: "B2"`, and `path: "$data.customer.name"`. Failed results contain no partial `data`.

`issues` describes problems with the returned file or entered values. Invalid application settings and file operations can throw exceptions, so use both `result.success` and `try/catch`. See the [API](./api.md) for details.

## A table with 20 input rows

Next, issue blank rows with a product list, fill two rows, and leave a gap. This example uses [validation and lists](./fields.md): `required` requires a value, `number|min:0` requires a nonnegative number, and `@choice` displays a product and returns its ID.

### 1. Prepare the template

Create a subdirectory inside the application from the first tutorial:

```sh
mkdir table-form
cd table-form
```

It uses the application's installed packages and module settings. Save the following files and run commands in `table-form`.

Create an `Input` sheet in `form-template.xlsx`, or download the [template](/examples/tutorials/form-template.xlsx):

| Cell | Content |
| --- | --- |
| A1:C1 | Headers: `Product`, `Quantity`, `Date` |
| A2 | `{#items}` |
| A3 | `{.product}{@choice:Products; key=id; label=name; return=key}{@validate:required\|string}` |
| B3 | `{.quantity}{@validate:required\|number\|min:0}` |
| C3 | `{?.date}{@validate:string}` |
| C4 | `{/items}` |

The repeat covers row A3:C3. Style it and set B3 to number format `0.00`. C3 needs text format `@`: enter the date as the string `2026-01-15`. Native Excel dates are unsupported in input fields; the `string` rule checks the type, not calendar validity.

### 2. Issue the rows

Save [issue-form.ts](/examples/tutorials/issue-form.ts) next to the template:

```ts
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookForm } from 'sheetbind'

const source = await readFile('form-template.xlsx')
const template = await importWorkbookXlsx(source)
const dictionaries = {
  Products: [{ id: '001', name: 'Paper' }, { id: '002', name: 'Paper' }],
}
const data = { items: Array.from({ length: 20 }, () => ({})) }
const issued = await renderWorkbookForm(template, data, { dictionaries })

await mkdir('saved-form', { recursive: true })
await writeFile('saved-form/template.xlsx', source)
await writeFile('saved-form/dictionaries.json', JSON.stringify(dictionaries, null, 2))
await writeFile('saved-form/issued.xlsx', issued)
```

```sh
pnpm exec tsx issue-form.ts
```

Twenty `{}` objects create twenty input rows. Required fields may be empty at issuance. The products have identical labels, so the dropdown shows `Paper [001]` and `Paper [002]`.

This example uses `return=key`. Keep **the dictionary used for issuance** with the template: the reader converts a label to an ID using the supplied dictionary. Current API data may give a different mapping. Object mode and other sources are covered in [validation and lists](./fields.md).

### 3. Fill and read

Open `saved-form/issued.xlsx`. Input rows occupy rows 4–23:

| Row | Product | Quantity | Date |
| --- | --- | --- | --- |
| 4 | Select `Paper [001]` | Number `2` | Text `2026-01-15` |
| 5 | Leave empty | | |
| 6 | Select `Paper [002]` | Number `0` | Leave empty |

Save the file as `completed.xlsx` next to `issue-form.ts`. Download [read-completed-form.ts](/examples/tutorials/read-completed-form.ts) into that directory and run:

```sh
pnpm exec tsx read-completed-form.ts
```

This is a complete handler with `result.success` and `try/catch`, as in the first form. It loads `saved-form/template.xlsx`, `saved-form/dictionaries.json`, and `completed.xlsx` from the current directory.

```json
{
  "items": [
    { "product": "001", "quantity": 2, "date": "2026-01-15" },
    { "product": "002", "quantity": 0, "date": null }
  ]
}
```

Empty row 5 and the remaining unfilled rows are omitted; `0` is preserved. Entering `Unknown` in A6 produces a `choice` issue with address `A6` and path `$data.items[1].product`. The address refers to Excel; the index refers to data after empty rows are omitted. Select a value from the list and read the file again.

## Empty rows and values

For a single-row repeat with input fields, `items: []` creates one blank row. To get a specific number of rows, pass that many `{}` objects as shown above.

Completely blank single-row records are omitted on read. Partially completed rows are validated; a row with a nonempty invalid value is also retained. `null`, empty text, and whitespace-only text count as empty; `0` and `false` do not. A blank cell in a retained record is returned as `null`.

Types are not converted: the text `"2"` remains a string and fails `number`. A formula in an input field produces a `formula` issue even if Excel has cached its result. Calculated cells without a binding are not included in submitted data.

## Adding, deleting, and reordering records

### Single-row records

In the table above, copy an entire Excel row and insert the copy inside the input region. Delete an entire row to remove it; cut and insert an entire row in the same region to reorder it. Copying retains formatting, formulas, and lists. Clearing every field of a single-row record omits it from the result.

Forms support repetition down rows. Sideways repetition is for reports. Place independent form lists one below another: their row ranges must not overlap. Nested lists are supported.

Do not delete hidden control rows or columns separately from records, and do not rename input sheets.

### Multirow records

When a record occupies several rows, copy the whole block with its hidden boundaries. Download [contacts-template.xlsx](/examples/tutorials/contacts-template.xlsx) and the [issued form with two records](/examples/tutorials/contacts-issued.xlsx). The `Contacts` sheet has these tags:

| Cell | Content |
| --- | --- |
| A1 | `{#contacts}` |
| A2 | `Name` |
| B2 | `{.name}` |
| A3 | `Email` |
| B3 | `{.email}` |
| B4 | `{/contacts}` |

In the issued form, the first record occupies entire rows **3:6**, and the second occupies **7:10**. Fields are visible on rows 4:5 and 8:9; the other rows in each block are hidden.

1. Enter `3:6` in Excel's Name Box to the left of the formula bar and press Enter. This selects entire rows, including hidden boundaries.
2. Copy them. Select row 7 and use Insert Copied Cells to insert entire rows.

After this single copy operation, save the workbook as `contacts-completed.xlsx` next to `contacts-template.xlsx`. Download [read-contacts.ts](/examples/tutorials/read-contacts.ts) into that directory and run:

```sh
pnpm exec tsx read-contacts.ts
```

The result contains three records:

```json
{
  "contacts": [
    { "name": "Alex", "email": "alex@example.com" },
    { "name": "Alex", "email": "alex@example.com" },
    { "name": "Sam", "email": "sam@example.com" }
  ]
}
```

Deleting and moving records are separate actions:

- To delete a record, select the whole block and use Delete Sheet Rows. The Delete key only clears cells.
- To move a record, cut the whole block and insert the cut rows before another block's boundary.

These numbers apply to the original example: inserting rows shifts later blocks. Do not copy only the visible fields. **Clearing a multirow block retains an empty record**, which may fail `required`; deleting the block removes the record. When copying a parent with a nested list, include the whole list and the parent's boundaries.

Reissue forms after changing template bindings, rules, or structure. Keep the original template and rule handlers for reading. Your application defines comparison with earlier data and permitted changes. Other file behavior is covered in [Excel and limitations](./xlsx.md).

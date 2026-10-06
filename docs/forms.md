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

Open `first-form-issued.xlsx`. The field is in B1. The form preserves this address. Change the name to `Alex` and save a **separate file named `completed.xlsx`** next to the script. Keep the original template for reading.

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

Clear B1 and save the file again to get `success: false`. The `issues` array contains a `required` issue with `sheetName: "Input"`, `address: "B1"`, and `path: "$data.customer.name"`. Parsed fields remain in `data` so they can be shown for correction; `success: false` means the data have not passed validation. Invalid file structure produces no `data`.

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

The repeat covers row A3:C3. Style it and set B3 to number format `0.00`. C3 needs text format `@`: enter the date as the string `2026-01-15`. Native Excel dates are also supported and read as ISO strings. This example uses text to preserve the `YYYY-MM-DD` representation; the `string` rule checks the type, not calendar validity.

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
await writeFile('saved-form/issued.xlsx', issued)
```

```sh
pnpm exec tsx issue-form.ts
```

Twenty `{}` objects create twenty input rows. Required fields may be empty at issuance. The products have identical labels, so the dropdown shows `Paper [001]` and `Paper [002]`.

This is the top of the [issued form](/examples/tutorials/issued.xlsx). Input rows start at 3 and continue through 22; control row 2 is hidden.

<div class="workbook-preview" tabindex="0" role="region" aria-label="Blank form, columns A–C, top of the table">

[![Issued form: Product, Quantity and Date headers on row 1, with blank input rows starting at 3. Column A provides product lists.](/images/form-issued.png)](/images/form-issued.png)

</div>

The form stores the issuance dictionaries in its hidden sheet. Reading needs the original template and the completed file; `return=key` returns the selected ID. See [validation and lists](./fields.md) for other sources.

### 3. Fill and read

Open `saved-form/issued.xlsx`. Input rows occupy rows 3–22:

| Row | Product | Quantity | Date |
| --- | --- | --- | --- |
| 3 | Select `Paper [001]` | Number `2` | Text `2026-01-15` |
| 4 | Leave empty | | |
| 5 | Select `Paper [002]` | Number `0` | Leave empty |

After filling two rows, the form looks like this:

<div class="workbook-preview" tabindex="0" role="region" aria-label="Completed form, columns A–C, rows 3–5">

[![Row 3 contains Paper [001], quantity 2.00 and date 2026-01-15. Row 4 is empty. Row 5 contains Paper [002] and quantity 0.00, with no date.](/images/form-completed.png)](/images/form-completed.png)

</div>

Save the file as `completed.xlsx` next to `issue-form.ts`, or download the [completed form](/examples/tutorials/completed.xlsx). Reading requires `saved-form/template.xlsx` from the previous step. Download [read-completed-form.ts](/examples/tutorials/read-completed-form.ts) next to `completed.xlsx` and run:

```sh
pnpm exec tsx read-completed-form.ts
```

This is a complete handler with `result.success` and `try/catch`, as in the first form. It loads `saved-form/template.xlsx` and `completed.xlsx` from the current directory.

The result is also available as a [JSON download](/examples/tutorials/completed.json):

<<< @/public/examples/tutorials/completed.json

Empty row 4 and the remaining unfilled rows are omitted; `0` is preserved. Cell addresses refer to Excel; data indexes refer to the array after empty rows are omitted.

To try a reading error, download [invalid.xlsx](/examples/tutorials/invalid.xlsx) and save it in the example directory as `completed.xlsx`. It already contains `Unknown` in A5. Excel's list validation blocks normal entry of that value, so this step uses a prepared file. Running the script again returns `success: false` and a `choice` issue with address `A5` and path `$data.items[1].product`; the [full result](/examples/tutorials/invalid.json) is available to download. Select `Paper [002]` in A5, save the file and read it again.

## Empty rows and values

For a single-row repeat with input fields, `items: []`, `items: null` or an omitted `items` property creates one blank row. Nested input tables behave the same way. To get a specific number of rows, pass that many `{}` objects as shown above.

Completely blank single-row records are omitted on read. Partially completed rows are validated; a row with a nonempty invalid value is also retained. `null`, empty text, and whitespace-only text count as empty; `0` and `false` do not. A blank cell in a retained record is returned as `null`.

Types and precision follow the [submitted cell's format](./xlsx.md#input-formats-and-value-types): for example, `"2"` with a numeric format reads as a number, while `@` keeps it as text. A formula in an input field uses the result saved by Excel. A missing result or an Excel error produces a `formula` issue. Calculated cells without a binding are not included in submitted data.

## Adding, deleting, and reordering records

### Single-row records

In the table above, copy an entire Excel row and insert the copy inside the input region. Delete an entire row to remove it; cut and insert an entire row in the same region to reorder it. Copying retains formatting, formulas, and lists. Clearing every field of a single-row record omits it from the result.

This example uses a shared dictionary. Forms with [local options](./fields.md) support filling the issued rows; changing those records or their order requires issuing a new form.

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

In the issued form, the first record occupies entire rows **2:5**, and the second occupies **6:9**. Fields are visible on rows 3:4 and 7:8; the other rows in each block are hidden.

1. Enter `2:5` in Excel's Name Box to the left of the formula bar and press Enter. This selects entire rows, including hidden boundaries.
2. Copy them. Enter `6:6` in the Name Box and press Enter to select hidden row 6. Use Insert Copied Cells to insert entire rows.

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

## File structure and version

Forms do not prepend a control row or store a template hash. Hidden repeat and sheet-end boundaries describe the current records and allow layout checks. Reading uses bindings and rules from the supplied template without requiring a matching hash. Applications own document type and version checks. Reissue files with the older control-marker format.

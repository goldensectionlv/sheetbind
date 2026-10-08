# Templates

A template is an XLSX file with tags in its cells. Prepare one item row, pass three records and get a table with a total.

To run the example, complete the [setup in First report](./getting-started.md). The images below show the files used in this example. Open an image for its full size; on narrow screens, scroll the tables horizontally.

## Repeated rows and a total {#report-example}

### 1. Template: one item row

Download [report-template.xlsx](/examples/tutorials/report-template.xlsx). Its `Order` sheet contains a header, an item row and a total:

<div class="workbook-preview" tabindex="0" role="region" aria-label="Report template, columns A–E">

[![Template: opening marker in A3, item row A4:E4, closing marker in E5 and total in D6.](/images/report-template.png)](/images/report-template.png)

</div>

B1 contains `{customer.name}`. The customer's name will replace this tag.

Two markers define the repeated area:

- **A3 — `{#items}`:** starts the `items` list at column A.
- **E5 — `{/items}`:** ends the list at column E.
- **Between them — A4:E4:** this entire row repeats for each item. `{.name}`, `{.quantity}` and `{.price}` read the current item; `{?.note}` leaves a blank cell when there is no note.

The blue marker rows disappear from the report. Colors only illustrate the example; tags determine its behavior. Each marker must be the only value on its entire row.

The template image displays **formula text**: D4 contains `=B4*C4`, D6 contains `=SUM(D2:D5)`. In a localized Excel application, use its local function names.

::: details Create this template manually: cells and tags to copy

Create an `Order` sheet in a new workbook:

| Cell | Content |
| --- | --- |
| A1 | `Customer` |
| B1 | `{customer.name}` |
| A2:E2 | Headers: `Item`, `Quantity`, `Price`, `Amount`, `Note` |
| A3 | `{#items}` |
| A4 | `{.name}` |
| B4 | `{.quantity}` |
| C4 | `{.price}` |
| D4 | Formula `=B4*C4` |
| E4 | `{?.note}` |
| E5 | `{/items}` |
| A6 | `Total` |
| D6 | Formula `=SUM(D2:D5)` |

Save as `report-template.xlsx`. Style row 4 and set the number format `0.00` for price and amount. Keep labels, formatting and formulas in Excel.

:::

### 2. Data: three items

Download [report-3.json](/examples/tutorials/report-3.json) beside the template. This is the complete object passed to the renderer:

<<< @/public/examples/tutorials/report-3.json

The `items` array contains three objects, so one template row produces three report rows. `Paper` and `Folder` have no `note`; the quantity of `Pen` is zero.

### 3. Result: table and total

<div class="workbook-preview" tabindex="0" role="region" aria-label="Report with three items, columns A–E">

[![Report: Paper — 2 at 5.00, amount 10.00; Pen — 0 at 3.00, amount 0.00; Folder — 3 at 4.00, amount 12.00. The total in D6 is 22.00.](/images/report-3.png)](/images/report-3.png)

</div>

Download the [completed report-3.xlsx](/examples/tutorials/report-3.xlsx) to inspect it immediately.

Items occupy rows **3–5**. The marker rows are gone and the item formatting repeats. Missing notes remain blank and the zero quantity stays zero.

D3 now contains `=B3*C3`, D4 contains `=B4*C4`, and D5 contains `=B5*C5`. The total in D6 is `=SUM(D2:D5)`.

**The image shows recalculated values: 10 + 0 + 12 = 22.** Sheetbind moves formula references but does not calculate them. Excel recalculates on opening; another viewer may show blank or stale values until recalculation. See [Excel and limitations](./xlsx.md).

### 4. Generate this file in your application

Save [report.ts](/examples/tutorials/report.ts) beside `report-template.xlsx` and `report-3.json`. Run commands from this directory, with the packages installed as described in [First report](./getting-started.md).

<<< @/public/examples/tutorials/report.ts

```sh
pnpm exec tsx report.ts
```

Open the generated `report.xlsx`. Its contents should match the result above.

## Fewer items

The same template handles arrays of different lengths. To try them with your script, download the JSON you want, replace `report-3.json` in `report.ts` with its filename and run the command again. The script overwrites `report.xlsx`.

**One item** — [report-1.json data](/examples/tutorials/report-1.json), [completed report-1.xlsx](/examples/tutorials/report-1.xlsx):

<div class="workbook-preview" tabindex="0" role="region" aria-label="Report with one item, columns A–E">

[![One Paper item on row 3. The total moves to D4 and equals 10.00 after recalculation.](/images/report-1.png)](/images/report-1.png)

</div>

**No items, `items: []`** — [report-0.json data](/examples/tutorials/report-0.json), [completed report-0.xlsx](/examples/tutorials/report-0.xlsx):

<div class="workbook-preview" tabindex="0" role="region" aria-label="Report without items, columns A–E">

[![Empty report: customer, headers and a total of 0.00 in D3. There are no item rows.](/images/report-0.png)](/images/report-0.png)

</div>

The total follows the table: D4 for one item, D3 for an empty array. `SUM` includes the header and ignores its text, so the empty report keeps the valid formula `=SUM(D2:D2)`.

Omitting `items` or passing `items: null` has the same effect as `items: []`: the report contains no item rows. This also applies to nested repeats and repeats across columns; you do not need to add empty arrays to the data. A supplied nonempty array must contain objects. A number, string or object in place of the array remains an error. For **blank rows that a user will fill in**, use a [form](./forms.md#a-table-with-20-input-rows).

## Repeat boundaries

The closing marker must be below the opening marker, in the same column or to its right. The path must match: `{#items}` closes with `{/items}`, while `{#.items}` closes with `{/.items}`.

If only column A repeats, check the closing marker: a closing tag in A defines a one-column body. Our example puts it in E to repeat A:E.

For a block several rows high, leave all its body rows between the markers. Blank rows in that rectangle repeat too. Merged cells and nested blocks must fit inside their block; they cannot cross its boundary.


## Values and paths

A field cell contains one binding and, optionally, its [validation or list directives](./fields.md).

| Tag | Meaning |
| --- | --- |
| `{customer.name}` | The `name` property inside `customer` in the current context |
| `{.name}` | The `name` property of the current record |
| `{$root.customer.name}` | The path from the root data object, including inside a repeat |
| `{?customer.note}` | A missing value is allowed; the output cell is blank |
| `{?.note}` | A missing `note` on the current record is allowed |

Outside a block, the current context is the root data object. Inside a repeat, it is one array item; inside `with`, it is the selected object. `{name}` and `{.name}` select the same property. The leading dot makes a local reference easier to see.

For example, `{customer.name}` and `{?customer.note}` accept:

```json
{
  "customer": {
    "name": "Sample customer"
  }
}
```

The name cell contains `Sample customer`; the note cell is blank. A missing property without `?` emits a warning and leaves the cell blank. `null` is blank; `0` and `false` retain their values. Ordinary fields accept strings, finite numbers, booleans and `null`. To display an object, select a property or use an [object choice](./fields.md).

Paths use property names separated by dots. Tags do not run JavaScript: expressions such as `{price * quantity}` or `{items[0].name}` are unsupported. Calculate values in your application or use an Excel formula.

Text outside bindings stays as written. `Customer: {customer.name}` is not interpolated. Put the label in a separate cell, or prepare the whole text in your data. To display a literal `{customer.name}`, write <code v-pre>{{customer.name}}</code>. A cell that starts with a binding cannot append other text or a second binding.

## If the workbook already exists

Work on a copy of the file. For the table above, start with [ordinary.xlsx](/examples/tutorials/ordinary.xlsx):

<div class="workbook-preview" tabindex="0" role="region" aria-label="Original workbook before adding tags, columns A–E">

[![Ordinary workbook before tagging: Sample customer, one Paper item, two units at 5.00. The total in D4 is 10.00.](/images/ordinary.png)](/images/ordinary.png)

</div>

Formulas are shown as recalculated values: D3 contains `=B3*C3`, D4 contains `=SUM(D2:D3)`.

1. Replace B1 with `{customer.name}`.
2. Insert a whole blank row before the item and put `{#items}` in A3.
3. Insert a whole blank row between the item and total and put `{/items}` in E5. Leave the other cells of both marker rows blank.
4. Replace A4, B4, C4 and E4 with `{.name}`, `{.quantity}`, `{.price}` and `{?.note}`.
5. Check the formulas: `=B4*C4` in D4 and `=SUM(D2:D5)` in D6.
6. Save as `report-template.xlsx` and run `report.ts` above.

If your workbook has several sample items, keep one as the repeated body and remove the others from the template copy. Keep its formatting. Move the closing marker right when more columns belong to the item. Check the generated values, formulas and workbook features you use in Excel.

Sheetbind accepts an already tagged file. These are editing instructions; there is no automatic conversion of an arbitrary workbook into a template.

## An object block with `with`

Use `with` to avoid repeating a long object path. It changes the current context without repeating the body.

| Cell | Content |
| --- | --- |
| A1 | `{#with customer}` |
| A2 | `Name` |
| B2 | `{.name}` |
| A3 | `City` |
| B3 | `{.address.city}` |
| B4 | `{/with}` |

```json
{
  "customer": {
    "name": "Sample customer",
    "address": { "city": "Oslo" }
  }
}
```

The report contains the name and city in B1 and B2 after removing the marker rows. `with` expects an object; it always closes with `{/with}`. Use `{$root.some.path}` to read outside the selected object.

A larger example combines an object block and repeated work rows: [template](/examples/regions/template.xlsx), [data](/examples/regions/template.data.json).

## Nested repeats

Put one complete block inside another. Each inner array belongs to the current outer record.

| Cell | Content |
| --- | --- |
| A1 | `{#groups}` |
| A2 | `{.name}` |
| A3 | `{#.items}` |
| A4 | `{.name}` |
| B4 | `{.quantity}` |
| B5 | `{/.items}` |
| B6 | `{/groups}` |

```json
{
  "groups": [
    {
      "name": "Office",
      "items": [{ "name": "Paper", "quantity": 2 }, { "name": "Pen", "quantity": 0 }]
    },
    { "name": "Storage", "items": [] }
  ]
}
```

The report shows `Office`, its two item rows, then `Storage`. The empty inner array removes only its item body; the group name remains. The opening and closing markers of the inner block must both be inside the outer rectangle.

Download a larger [nested template](/examples/regions/nested.xlsx) and its [data](/examples/regions/nested.data.json).

## Repeated columns

Add `axis=columns` to repeat the body to the right. The markers still sit above and below the source rectangle:

| Cell | Content |
| --- | --- |
| A1 | `{#offers \| axis=columns}` |
| A2 | `{.name}` |
| A3 | `{.price}` |
| A4 | `{/offers}` |

```json
{
  "offers": [
    { "name": "Option A", "price": 10 },
    { "name": "Option B", "price": 12 },
    { "name": "Option C", "price": 9 }
  ]
}
```

The body is `A2:A3`. The report places names in A1:C1 and prices in A2:C2. To repeat a wider block, put its closing marker in the last column of that block. Set column widths in Excel.

Column repeats work in reports. Editable forms support row repeats. Download the [comparison template](/examples/comparison/template.xlsx) and [data](/examples/comparison/template.data.json) for a report combining both directions.

## Run the additional examples

The JSON files for the region and comparison examples contain a `data` property. Save the desired workbook as `template.xlsx` and its JSON as `template.data.json`, then use this `report.ts`:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('template.xlsx'))
const { data } = JSON.parse(await readFile('template.data.json', 'utf8'))
await writeFile('report.xlsx', await renderWorkbookReport(template, data))
```

Run `pnpm exec tsx report.ts` and open `report.xlsx`. The [nested example data](/examples/regions/nested.data.json) contains three sites with 0, 1 and 3 work records. The [comparison data](/examples/comparison/template.data.json) contains three items and three offers; a missing price stays blank.

## Next step

Add [validation and lists](./fields.md) to fields, or continue with [filling in a form](./forms.md).

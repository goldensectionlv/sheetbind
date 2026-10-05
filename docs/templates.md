# Templates

A template is an XLSX file with tags in its cells. Keep labels, formatting and formulas in Excel. Tags select values from the data and mark areas to repeat.

This guide continues [First report](./getting-started.md). Run its scripts from the application directory where Sheetbind, ExcelJS and `tsx` are installed.

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

The name cell contains `Sample customer`; the note cell is blank. Without `?`, a missing property causes `missing-source`. `null` produces a blank cell; `0` and `false` keep their values. Ordinary fields accept strings, finite numbers, booleans and `null`. To display an object, select one of its properties or use an [object choice](./fields.md).

Paths use property names separated by dots. Tags do not run JavaScript: expressions such as `{price * quantity}` or `{items[0].name}` are unsupported. Calculate values in your application or use an Excel formula.

Text outside bindings stays as written. `Customer: {customer.name}` is not interpolated. Put the label in a separate cell, or prepare the whole text in your data. To display a literal `{customer.name}`, write <code v-pre>{{customer.name}}</code>. A cell that starts with a binding cannot append other text or a second binding.

## Repeated rows and a total

Build a table of items with quantity, price, amount and an optional note. The same template will produce reports for 0, 1 and 3 records.

### Create the template

Download [report-template.xlsx](/examples/tutorials/report-template.xlsx), or create an `Order` sheet with these cells:

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

Save the file as `report-template.xlsx`. Style row 4 and set the number format `0.00` for price and amount. In a localized Excel application, enter formulas using its local function names.

The two markers define the rectangle:

- `{#items}` marks the row above the body and its leftmost column.
- `{/items}` marks the row below the body and its rightmost column.
- Here the body is `A4:E4`, including blank cells. All of it repeats for each item.

Each marker must be the only value on its entire row. The closing marker must be below the opening marker, in the same column or to its right. The path must match: `{#items}` closes with `{/items}`, while `{#.items}` closes with `{/.items}`. Both marker rows disappear from the report.

For a block several rows high, leave all its body rows between the markers. Blank rows in that rectangle repeat too. Merged cells and nested blocks must fit inside their block; they cannot cross its boundary.

If only column A repeats, check the closing marker: a closing tag in A defines a one-column body. Move it to E when A:E should repeat.

### Pass data and run

Save this code as `report.ts` beside the template, replacing the script from the first guide. You can also [download it](/examples/tutorials/report.ts).

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('report-template.xlsx'))
const items = [
  { name: 'Paper', quantity: 2, price: 5 },
  { name: 'Pen', quantity: 0, price: 3, note: 'Spare' },
  { name: 'Folder', quantity: 3, price: 4 },
]

for (const count of [0, 1, 3]) {
  const data = { customer: { name: 'Sample customer' }, items: items.slice(0, count) }
  await writeFile(`report-${count}.xlsx`, await renderWorkbookReport(template, data))
}
```

```sh
pnpm exec tsx report.ts
```

The full inputs are also available as JSON: [0 records](/examples/tutorials/report-0.json), [1 record](/examples/tutorials/report-1.json), [3 records](/examples/tutorials/report-3.json).

### Check the results

Open the three files from the application directory:

| File | Item rows in Excel | Total cell | Total after recalculation |
| --- | --- | --- | --- |
| `report-0.xlsx` | None | D3 | 0 |
| `report-1.xlsx` | Row 3 | D4 | 10 |
| `report-3.xlsx` | Rows 3–5 | D6 | 22 |

`items: []` removes the repeated body. Omitting `items` is an error; the repeat expects an array of objects. The zero quantity for `Pen` remains zero.

The total formula includes the header and the body. `SUM` ignores header text, so the empty report still has a valid range. Sheetbind moves formula references but does not calculate results. Excel recalculates on opening; another viewer may show an empty or stale result until recalculation. See [Excel and limitations](./xlsx.md) for formula support.

## If the workbook already exists

Work on a copy of the file. For the table above, start with [ordinary.xlsx](/examples/tutorials/ordinary.xlsx):

| Row | A | B | C | D | E |
| --- | --- | --- | --- | --- | --- |
| 1 | Customer | Sample customer | | | |
| 2 | Item | Quantity | Price | Amount | Note |
| 3 | Paper | 2 | 5 | `=B3*C3` | |
| 4 | Total | | | `=SUM(D2:D3)` | |

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

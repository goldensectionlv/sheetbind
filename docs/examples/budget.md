# Monthly budget

Build an expense table with categories down the rows, months across the columns, and formulas for both sets of totals. The data determines the number of categories and months.

First [install the package](../getting-started.md). Create a separate directory inside your application for this example:

```sh
mkdir budget-example
cd budget-example
```

## 1. Inspect the template

Download [budget-template.xlsx](/examples/walkthroughs/budget-template.xlsx). Its `Budget` sheet has one category row and one month column:

<div class="workbook-preview" tabindex="0" role="region" aria-label="Budget template, A1:C12">

[![Template with categories in A, a month in B and category totals in C. The data row is row 7.](/images/example-budget-template.png)](/images/example-budget-template.png)

</div>

The image displays formulas as text. The downloadable XLSX contains ordinary Excel formulas.

| Template region | What repeats |
| --- | --- |
| B2:B12, `{#months \| axis=columns}` | The entire month column: heading, amounts and bottom total |
| A4:A10, `{#categories}` | Category names down the rows |
| B5:B9, `{#.amounts}` | The current month's amounts down the rows |
| C6:C8, `{#categories}` | Each category's total down the rows |

All three vertical repeats contain row 7. Their markers occupy separate rows: a marker row must contain only one structural tag. Rendering removes the marker rows and keeps the values aligned.

Both `categories` repeats use the same array: one produces names, the other total formulas. `C7` contains `=IFERROR(SUM(B7:B7),0)`. Expanding the months extends its range across their columns. `B11` contains `=SUM(B3:B10)` and totals its own month.

## 2. Prepare the data

Save [budget.data.json](/examples/walkthroughs/budget.data.json) beside the template:

<<< @/public/examples/walkthroughs/budget.data.json

**The order of `amounts` must match `categories`.** The application owns this alignment. Keep `{}` in place for a missing amount: deleting the element shifts later amounts to different categories.

February has `0` for `Transport` and no value for `Books`. The `{?.value}` binding permits the missing property and leaves the cell blank. `SUM` adds the supplied numbers.

## 3. Create the report

Download [budget.ts](/examples/walkthroughs/budget.ts) into the same directory:

<<< @/public/examples/walkthroughs/budget.ts

```sh
pnpm exec tsx budget.ts
```

Open `budget.xlsx`, or download the [rendered report](/examples/walkthroughs/budget.xlsx):

<div class="workbook-preview" tabindex="0" role="region" aria-label="Budget for three months and three categories, A1:E6">

[![January totals 335, February 225 and March 325, with a grand total of 885. February Books is blank; Transport is zero.](/images/example-budget.png)](/images/example-budget.png)

</div>

`Food` totals 725, `Transport` 105 and `Books` 55. The grand total is **885**. The first category's formula moved from C7 to E3 and became `=IFERROR(SUM(B3:D3),0)`.

The image shows recalculated values. **Sheetbind relocates formulas; Excel calculates them when opening the workbook.** Other viewers may display blank values before recalculation.

## 4. Change the table size

- Add a category and a matching element to every `amounts` array to add a row.
- Add a `months` object with an amount for each category to add a column.
- For no categories, pass `categories: []` and an empty `amounts` array in every month. Month headings and zero totals remain: [example](/examples/walkthroughs/budget-0x3.xlsx).
- With `months: []`, category names and zero totals remain: [example](/examples/walkthroughs/budget-3x0.xlsx). `IFERROR` handles the removed amount range.

This example creates a report. [Editable forms](./registration.md) use row repeats; column repeats are for reports.

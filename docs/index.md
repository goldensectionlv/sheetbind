# Sheetbind

This guide and its downloadable examples use Sheetbind 0.3.1. Start with the [installation guide](./getting-started.md).

Sheetbind fills XLSX templates with data. You prepare and format the workbook in Excel; your application supplies the values. The result is a report or a form that users can fill in and return.

## From template to report

**Template.** One item row between `{#items}` and `{/items}`. Its cells contain tags and formulas:

<div class="workbook-preview" tabindex="0" role="region" aria-label="Item table template, columns A–E">

[![Template: item row A4:E4 between markers in A3 and E5. A total formula is in D6.](/images/report-template.png)](/images/report-template.png)

</div>

**Result for three items.** Pass `Paper`, `Pen` and `Folder` in the `items` array. Get three rows with the same formatting:

<div class="workbook-preview" tabindex="0" role="region" aria-label="Completed item table, columns A–E">

[![Report with Paper, Pen and Folder on rows 3–5. The recalculated total is 22.00.](/images/report-3.png)](/images/report-3.png)

</div>

The images show the example files: formula text in the template, recalculated values in the report. Sheetbind moves formulas; Excel calculates them. Open an image for a larger view, or scroll it horizontally on a narrow screen.

[Follow the example: template, JSON, code and completed XLSX](./templates.md#report-example).

## Start here

1. [First report](./getting-started.md) — install the package and create one XLSX file.
2. [Templates](./templates.md) — add repeated rows, totals and nested data.
3. [Validation and lists](./fields.md) — check field values and offer dropdown choices.
4. [Filling in a form](./forms.md) — issue a file, edit it in Excel and read the result.

The instructions include downloadable templates, code and expected results. The same example files are used in both languages.

## Examples

- [Monthly budget](./examples/budget.md) — grow rows and columns, keep blanks and zero values, calculate totals.
- [Learning plan](./examples/study-plan.md) — courses with different numbers of lessons, merged headings and subtotals.
- [Event registration](./examples/registration.md) — nested participant lists, shared ticket choices, validation and reading completed forms.

Each example includes a template, data, a runnable script and the expected result. All data is fictional.

## Reference

- [API](./api.md) — function arguments, results and errors.
- [Excel and limitations](./xlsx.md) — formatting, formulas and workbook features.
- [Development](./development.md) and [Architecture](./architecture.md) — for changes to the library itself.

Requires Node.js 22.13 or newer. Licensed under MIT.

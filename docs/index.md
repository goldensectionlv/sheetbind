# Sheetbind

Sheetbind fills XLSX templates with data. You prepare and format the workbook in Excel; your application supplies the values. The result is a report or a form that users can fill in and return.

For example, put `{customer.name}` in a cell and pass `{ customer: { name: 'Sample customer' } }`. The output cell contains `Sample customer`.

## Start here

1. [First report](./getting-started.md) — install the package and create one XLSX file.
2. [Templates](./templates.md) — add repeated rows, totals and nested data.
3. [Validation and lists](./fields.md) — check field values and offer dropdown choices.
4. [Filling in a form](./forms.md) — issue a file, edit it in Excel and read the result.

The instructions include downloadable templates, code and expected results. The same example files are used in both languages.

## Reference

- [API](./api.md) — function arguments, results and errors.
- [Excel and limitations](./xlsx.md) — formatting, formulas and workbook features.
- [Development](./development.md) and [Architecture](./architecture.md) — for changes to the library itself.

Requires Node.js 22.13 or newer and ExcelJS 4.4.0. Licensed under MIT.

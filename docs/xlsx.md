# Excel and limitations

Sheetbind reads unencrypted `.xlsx` templates and returns `.xlsx` files. Edit the workbook's appearance in Excel before importing it. The sections below describe how that content follows the rendered data.

## Formatting and empty cells

Fonts, fills, borders, alignment, number formats, themes, row heights and column widths come from the template. A repeated region copies its formatting, including styled empty cells. Empty cells with notes, validation, conditional formatting or drawing anchors also participate in placement.

Content beside a repeat keeps the rows or columns it occupies. This can leave space between blocks of different sizes. A blank-looking cell or a gap does not by itself create a submitted record; form records follow the rules in [form reading](./forms.md).

Merged cells inside a repeat follow their copy. A merge outside it can stretch across inserted rows or columns. Keep structural tags outside merged cells and follow the [region layout rules](./templates.md).

## Input formats and value types

Form reading uses the format of the submitted cell before running validation. A numeric format converts numeric text to a number and rounds it to the declared decimal precision. Text format (`@`) preserves text and converts numeric values to strings.

| Input | Recommended format | Read result |
| --- | --- | --- |
| Quantity or amount | A numeric format such as `0.00` | A number: stored `10.075` becomes `10.08` |
| Identifier with leading zeroes | Text (`@`), with `string` validation | A string such as `"001"` |
| Date as text | Text (`@`), entered as `2026-01-15` | A string; add a custom calendar-date rule if needed |
| Native Excel date | A date format such as `dd.mm.yyyy` | An ISO string, for example `2026-01-15T00:00:00.000Z` |

For form fields with `string` validation or an active choice list, issuance changes `General` to text. An empty choice source with `emptySource=input` leaves the format unchanged. Explicit formats in the template take precedence.

`General` keeps numeric precision. It accepts numeric text with a decimal point or comma, while text with leading zeroes such as `"006"` remains text. Use `@` to make an identifier's type explicit. Fixed and optional decimal places, percent, scientific notation, scaling commas and numeric format sections determine rounding. Percent values stay fractional numbers: `0.123456` with `0.00%` reads as `0.1235`. Fraction and elapsed-time formats do not round the stored number. Reading does not change the workbook. Trailing zeroes are display information: the JSON number `12.30` is represented as `12.3`.

Native Excel dates become ISO strings; Sheetbind does not infer a timezone from the workbook. An ordinary numeric serial without a date format remains a number. Formulas entered into input fields produce a `formula` issue. Formula cells outside input fields are not returned as data.

List labels are matched using the original text: `"2026"` stays a list option, and selected objects or keys follow the `@choice` return setting. Numeric conversion applies to free input, including `emptySource=input`.

Use strings for long identifiers: Excel numeric precision can lose digits. Cell text is limited to 32,767 UTF-16 code units, and characters that XML cannot represent are rejected. Standard XLSX row and column limits also apply.

## Formulas and calculation

Write formulas in the template using ordinary Excel cells. Local cell and range references follow placement; references inside repeated formulas follow the corresponding copy. Named ranges follow the moved cells and keep their worksheet scope.

**Sheetbind does not calculate formulas.** It asks Excel to recalculate when the workbook opens. Another viewer may show blank or stale results until recalculation. If the server needs a total immediately, calculate it in the application and pass it as data.

When a repeat is empty and its whole referenced range disappears, a reference can become `#REF!`. For a total that should be zero, either include a fixed header row in the `SUM` range or explicitly handle the missing range:

```text
=IFERROR(SUM(Amounts),0)
```

Use ordinary A1 references for calculations that must expand with a repeat. Structured table references, external workbook links, 3D references and other Excel-specific expressions are not guaranteed to follow a repeat when its size changes. Their original content can survive in the file without every reference being adjusted. External links are not fetched.

For a formula that looks up an object-choice property, [workbookChoiceRange](./api.md) returns the generated dictionary range name.

## Native Excel tables

An Excel table does not define a Sheetbind repeat. Sheetbind relocates a retained table's range, but does not turn table rows into template records or interpret all structured table formulas.

For a data-driven repeat, convert its Excel table to an ordinary range, put Sheetbind markers around the body, and use formulas over that body. Other tables can remain in the workbook. Check their headers, totals, filter ranges and formulas in the rendered file.

## Printing

Paper size, orientation, margins, scaling and page settings are retained from the template. Print areas and repeated title rows follow the rendered coordinates, including changes introduced by forms.

Sheetbind returns an XLSX; Excel determines pagination and printing. Check print preview after rendering a different number of records, particularly when merged headings or large row heights affect page breaks.

## Notes, pictures and ranges

| Content | How it follows the output |
| --- | --- |
| Notes and hyperlinks | Their cell addresses move or repeat with the owning region |
| Pictures | Drawing anchors move or repeat with their region; both bounds move together |
| Native validation and conditional formatting | Their ranges follow the actual target cells, including otherwise empty cells |
| Filters and sort settings | Their enclosing range follows placement; a setting is removed when its whole target disappears |
| Defined names | Local references move; a fully removed reference becomes `#REF!` |

A rule applying to separate groups of cells can become several ranges. This keeps validation or a fill off the gaps between copies and off hidden form boundaries. Formulas and filters that need a single rectangle use an enclosing range, which can include cells between its ends.

A picture's position and a chart's data references are different concerns. Retaining a chart, drawing or other XLSX part does not guarantee that every formula or extension inside it is adjusted for repetition. Check those features in the application that will open the result.

## Dropdowns and existing validation

Sheetbind creates Excel dropdowns for [lists and choices](./fields.md). If a source cell has only an input prompt, that prompt is retained when adding the dropdown.

If the same output cell already has a full Excel validation rule, rendering rejects the conflict with `RangeError`. Remove the native rule or the Sheetbind list/choice for that field. A cell's native rule is not silently combined with a second validation rule.

Excel's validation UI does not replace form reading. Pasted values can still need checking; `readWorkbookForm` resolves the choice and runs the declared field rules.

## Check a workbook you depend on

Source XLSX parts outside the adapter's scope are retained where possible. This is not a guarantee for arbitrary Excel extensions, embedded content or every spreadsheet application.

For your template, render zero, one and several records. Open the files in the intended spreadsheet application and check values, formulas after recalculation, metadata placement and print preview. For forms, also edit and save a separate completed file, then read it with the original template. A successful report render does not verify that editing and reading path.

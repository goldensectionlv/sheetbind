# Validation and lists

Write rules in the same cell as the [binding](./templates.md):

```text
{quantity}
{@validate:required|number|min:0}
```

Tags for one field can start on separate lines within the same cell. In Excel for Windows, use Alt+Enter. Line breaks between tags do not change the binding or rules; writing them on one line also works.

Here `quantity` is the data field, `required` rejects an empty value, `number` requires a number, and `min:0` rejects negative numbers. Rules run when reading a completed form. Report rendering and form issuance do not execute validators or require their handlers.

## Global validation registration

Register application rules once at startup:

```ts
import { registerValidationRule } from 'sheetbind'

registerValidationRule('twoLetters', {
  validate: value => typeof value === 'string' && /^[A-Z]{2}$/.test(value),
  message: 'Enter two uppercase letters',
})
```

The cell `{code}{@validate:twoLetters}` uses this rule in `readWorkbookForm(template, bytes)` without per-call handlers. Per-call `validationRules` override global registrations. Built-in names cannot be replaced; duplicate registration of an application name is an error. Registration belongs to the loaded package instance; register separately in another process or worker.

## Built-in rules

The library has eight built-in rules:

| Rule | Check | Example |
| --- | --- | --- |
| `required` | The value is not empty | `required` |
| `string` | A string | `string` |
| `number` | A finite number | `number` |
| `boolean` | `true` or `false` | `boolean` |
| `object` | An object, not an array | `object` |
| `min` | A number at least as large as the limit | `min:0` |
| `max` | A number no larger than the limit | `max:24` |
| `maxLength` | A string no longer than the limit | `maxLength:100` |

`min` and `max` take one finite number. `maxLength` takes a nonnegative integer and counts UTF-16 code units.

Rules run from left to right. The first failure stops validation of that field; other fields are still checked. Multiple `@validate` directives in one cell append rules in the order written.

`null`, `""`, and whitespace-only strings count as empty. All built-in rules except `required` skip empty values. `0` and `false` are not empty. Rules do not convert values: the string `"12"` fails `number`, and whitespace is not trimmed.

For an optional field, omit `required`. The `?` binding prefix permits a missing field in report input; it does not cancel `required` when reading a form. An omitted property and a property containing `undefined` both mean absence, including in nested objects. You do not need to replace them with `null` before issuing a form.

## A list of strings

Use `@list` when the selected text is the value you need:

```text
{status}{@list:Statuses}
```

Create an `Input` sheet in `list-template.xlsx`: A1 = `Status`, B1 = this tag. Or download the [template](/examples/tutorials/list-template.xlsx). After [installation](./getting-started.md), save [list-report.ts](/examples/tutorials/list-report.ts) next to it:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('list-template.xlsx'))
const data = { status: 'Draft' }
const dictionaries = { Statuses: ['Draft', 'Ready'] }
await writeFile('list-report.xlsx', await renderWorkbookReport(template, data, { dictionaries }))
```

Run from that directory:

```sh
pnpm exec tsx list-report.ts
```

Cell B1 in `list-report.xlsx` contains `Draft` and offers the `Draft`, `Ready` list. Add `{@validate:required}` to require a selection.

This example creates a report. For a file the user will return to your application, replace `renderWorkbookReport` with `renderWorkbookForm` in both the import and the call. Reading that [form](./forms.md) returns the selected string. A dropdown alone does not make a report readable by `readWorkbookForm`.

The list must contain at least one string. The issued form stores this list; no dictionary argument is needed when reading. A value outside the list produces a `list` issue.

## Selecting an object or a key

Use `@choice` when each option has a separate identifier and display label:

```text
{product}{@choice:Products; key=id; label=name}
```

Create an `Input` sheet in `choice-template.xlsx`: A1 = `Product`, B1 = this tag. Or download the [template](/examples/tutorials/choice-template.xlsx). Save [choice-report.ts](/examples/tutorials/choice-report.ts) next to it:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'

const template = await importWorkbookXlsx(await readFile('choice-template.xlsx'))
const dictionaries = {
  Products: [
    { id: '001', name: 'Paper' },
    { id: '002', name: 'Pen' },
  ],
}
const data = { product: dictionaries.Products[0] }
await writeFile('choice-report.xlsx', await renderWorkbookReport(template, data, { dictionaries }))
```

```sh
pnpm exec tsx choice-report.ts
```

Cell B1 in `choice-report.xlsx` contains `Paper` and offers a list of products. `key=id` names the identifier field; `label=name` names the label shown in Excel. By default the data field holds the whole object.

To read the file back, create it with `renderWorkbookForm`, as in the [form tutorial](./forms.md). Reading a selection of `Paper` from that form returns:

```json
{ "product": { "id": "001", "name": "Paper" } }
```

To use just the identifier, replace the tag in B1 with this `return=key` version:

```text
{product}{@choice:Products; key=id; label=name; return=key}
```

Replace the `data` line in `choice-report.ts` and run the script again:

```ts
const data = { product: '001' }
```

The dictionary stays the same. Excel shows `Paper`; reading the form returns `{ "product": "001" }`. The key keeps its type: the string `'001'` does not become a number.

Keys must be unique nonempty strings or finite numbers. Labels must be nonempty strings. Duplicate labels appear in Excel as, for example, `Paper [001]` and `Paper [002]`. Rendering rejects the dictionary if adding keys still leaves ambiguous display labels.

You cannot combine `@list` and `@choice` in one field. Set requiredness and other checks separately with `@validate`.

## Where options come from

A `@choice` source can be a named dictionary or a path to an array of objects in the data:

| Source | Where the options are stored |
| --- | --- |
| `Products` | `options.dictionaries.Products` |
| `.availableProducts` | `availableProducts` on the current record |
| `$root.products` | `products` at the root of the data |

For example, inside an `items` repeat:

```text
{.product}{@choice:$root.products; key=id; label=name; return=key}
```

```ts
const data = {
  products: [{ id: '001', name: 'Paper' }],
  items: [{ product: '001' }],
}
```

Both reports and forms can use a separate `.availableProducts` array from each record. All list and choice sources are saved in the issued form. `return=key` changes only the returned value; it does not require a second dictionary input. The application separately checks whether the returned data is still acceptable.

## Different options for different rows

Put each question's answers directly on its record:

| Cell | Content |
| --- | --- |
| A1 | `{#questions}` |
| A2 | `{.question}` |
| B2 | `{.answer}{@choice:.answers; key=id; label=label; return=key}` |
| B3 | `{/questions}` |

```ts
const data = {
  questions: [
    { question: 'Colour', answers: [{ id: 'red', label: 'Red' }, { id: 'blue', label: 'Blue' }] },
    { question: 'Size', answers: [{ id: 'small', label: 'Small' }, { id: 'large', label: 'Large' }] },
  ],
}
const issued = await renderWorkbookForm(template, data)
const result = await readWorkbookForm(template, completed)
```

The first row offers `Red / Blue`, and the second offers `Small / Large`. No extra identifier or flattened dictionary is needed. Source arrays do not become submitted fields.

Local options are saved for the issued fields. Fill or clear values while keeping the same records and row order. To add, remove or reorder records, or change their options, issue a new form. Local sources are not automatically transferred between rows.

Omitted, `undefined` and `null` data sources such as `.answers` or `$root.catalog.options` behave like an empty array. You do not need to add `answers: []` to every record.

For questions that accept text when `answers` is empty or absent, add `emptySource=input` together with `return=key`. Populated sources still require a listed value, and `@validate` rules still apply. An empty source allows ordinary input without a dropdown.

If a named dictionary is absent, each resolve, render or read call emits one `console.warn` per missing name. The field keeps ordinary input without a dropdown, dictionary lookup or list validation; `@validate` still applies. Rendering keeps scalar input as supplied and uses the declared `label` property for an object value. Reading returns the cell value without reconstructing a key or object. Supplied dictionaries remain validated; an explicitly empty string list is invalid. Supply named dictionaries through `options.dictionaries` when issuing the file; available sources are saved in the form.

## Custom messages

Set a rule's message in the cell:

```text
{quantity}{@validate:required|number|min:0}{@validationMessage:required:"Enter a quantity"}
```

Or provide messages for the whole execution through `options.validationMessages`:

```ts
const options = { validationMessages: { required: 'Enter a value' } }
```

A message in the cell takes precedence over an execution message. To change the text for just one use of a rule, set `message` in that directive:

```text
{quantity}{@validate:number}{@validate:max:100; message="Enter at most 100"}
```

This message has the highest priority. With no overrides, the handler's message is used. Quote text that contains spaces or punctuation.

## A custom rule

A name in a tag does not create a new built-in rule: the application supplies its handler. Add `multipleOf` to require a whole number divisible by a positive integer argument.

Create an `Input` sheet in `rules-template.xlsx`, or download the [template](/examples/tutorials/rules-template.xlsx):

| Cell | Content |
| --- | --- |
| A1 | `Quantity` |
| B1 | `{quantity}{@validate:required\|number\|multipleOf:2}` |

After [installation](./getting-started.md), save [rules-form.ts](/examples/tutorials/rules-form.ts) next to it:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookForm, readWorkbookForm } from 'sheetbind'
import type { ValidationOptions } from 'sheetbind'

const options: ValidationOptions = {
  validationRules: {
    multipleOf: {
      validateArgs: args => args.length === 1 && typeof args[0] === 'number'
        && Number.isSafeInteger(args[0]) && args[0] > 0,
      validate(value, [step]) {
        return typeof value === 'number' && Number.isSafeInteger(value)
          && typeof step === 'number' && value % step === 0
      },
      message: ({ args }) => `Enter a whole number divisible by ${args[0]}`,
    },
  },
}
const template = await importWorkbookXlsx(await readFile('rules-template.xlsx'))
const data = { quantity: 4 }
await writeFile('rules-form.xlsx', await renderWorkbookForm(template, data))
const result = await readWorkbookForm(template, await readFile('rules-form.xlsx'), options)
console.log(JSON.stringify(result, null, 2))
```

Run from that directory:

```sh
pnpm exec tsx rules-form.ts
```

The script issues `rules-form.xlsx` and reads the saved file. Reading succeeds for `quantity: 4`. Changing the value to `3` still creates the file, but reading returns `success: false` with the `multipleOf` rule and the message `Enter a whole number divisible by 2`. `multipleOf:0` causes a configuration error during reading, before checking the value.

`validateArgs` checks the tag's arguments; a rule without it accepts no arguments. `validate` must synchronously return `true` or `false`. Empty values are skipped by default; set `skipEmpty: false` on the handler to check them. Do not mutate the data or perform asynchronous requests inside the handler.

The XLSX stores the rule's name and arguments. Its implementation stays in the application: supply `validationRules` only when reading a form. An unknown rule stops reading with a template error; provide its handler before reading. Built-in rules cannot be replaced. Handler types and execution errors are covered in the [API](./api.md).

Next: [issue a form, fill it in Excel, and read the result](./forms.md).

## Formatting values

Formatters transform values when rendering reports, issuing forms and calling `resolveWorkbook`. They do not mutate input data. Declare formatting in the cell:

```text
{deliveryDate | format_date:DD.MM.YYYY}
{enabled | bool_replace:"Yes","No"}
{price | float}
```

The equivalent multiline cell is:

```text
{deliveryDate}
{@format:format_date:DD.MM.YYYY}
```

| Formatter | Result |
| --- | --- |
| `float` | Converts numeric text to a finite number; preserves blanks and values that cannot be converted |
| `bool_replace:"Yes","No"` | First label for truthy values, second for falsy values |
| `format_date:DD.MM.YYYY` | Formats a date string or Unix timestamp in milliseconds as text; defaults to `DD.MM.YYYY` and uses the process's local date components |

Date masks support `YYYY`, `YY`, `MM`, `M`, `DD`, `D`, `HH`, `H`, `mm`, `m`, `ss`, `s`. Quote masks containing time separators: `format_date:"DD.MM.YYYY HH:mm"`. Unrecognized dates retain their input value. Use Excel number formats such as `0.00` for decimal places on cells that already contain numbers.

Register application formatters once at startup:

```ts
import { registerFormatter } from 'sheetbind'

registerFormatter('uppercase', value => value === null ? null : String(value).toUpperCase())
registerFormatter('suffix', (value, args) => String(value ?? '') + String(args[0] ?? ''))
```

```text
{code | uppercase | suffix:"!"}
```

Built-in formatters preserve `null`, empty strings and whitespace-only text. A blank answer does not become `0` or the negative label; actual `0` and `false` still get formatted.

Pipelines run left to right. Each handler receives a value and JSON argument array and synchronously returns a string, finite number, boolean or `null`. Unknown formatters, handler failures and invalid results stop rendering. Built-in names are reserved; duplicate registration of an application name is an error. Register separately for each loaded package instance, process or worker.

Formatters **never run on read** and are not inverted: a cell formatted with `bool_replace` returns the text “Yes”, not a boolean. Use `@choice` to return a key or object from a display label. Formatting cannot share a cell with `@choice` or `@list`. Handler registrations are not stored in XLSX.

The [complete runnable example](/examples/tutorials/formatting.ts) issues a form with built-in and custom formatters and reads it using a global rule. After installing dependencies, run `pnpm exec tsx formatting.ts`.

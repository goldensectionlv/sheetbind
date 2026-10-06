# Validation and lists

Write rules in the same cell as the [binding](./templates.md):

```text
{quantity}{@validate:required|number|min:0}
```

Here `quantity` is the data field, `required` rejects an empty value, `number` requires a number, and `min:0` rejects negative numbers. Rules run when building a report and reading a completed form. Issuing a form defers value validation until reading.

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

For an optional field, omit `required`. The `?` binding prefix permits a missing field in report input; it does not cancel `required` when reading a form. If a value is missing, omit the property from the input object: explicit `undefined` is not valid JSON input.

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

The list must contain at least one string. Pass the same list through `dictionaries` when reading the form. A value outside the list produces a `list` issue.

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

A report can use a separate `.availableProducts` array from each record. **Options inside form repeats must be shared:** use a named dictionary or a `$root` array. Per-record sources are rejected. The application checks whether a selected option is allowed for a particular record.

The source required when reading a form depends on the mode:

| Setting | What to supply when reading |
| --- | --- |
| `@list:Statuses` | The original list through `dictionaries` |
| `@choice:Products; …; return=key` | The original dictionary through `dictionaries` |
| `@choice:$root.products; …; return=key` | The original array through `context: { products }` |
| `@choice` without `return=key` | The source is saved in the issued form; no need to supply it again |

For `return=key`, retain a snapshot of the source at issuance. The reader matches the label against the supplied source: if a current dictionary gives that label a different ID, the result can change. `context` contains choice sources and is not added to submitted data. In object mode, the application separately checks whether the returned object is still acceptable.

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

After [installation](./getting-started.md), save [rules-report.ts](/examples/tutorials/rules-report.ts) next to it:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { importWorkbookXlsx, renderWorkbookReport } from 'sheetbind'
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
await writeFile('rules-report.xlsx', await renderWorkbookReport(template, data, options))
```

Run from that directory:

```sh
pnpm exec tsx rules-report.ts
```

Cell B1 in `rules-report.xlsx` contains `4`. Changing the value to `3` makes rendering throw `TemplateError` with the `multipleOf` rule and the message `Enter a whole number divisible by 2`. Using `multipleOf:0` produces a configuration error before checking the value.

`validateArgs` checks the tag's arguments; a rule without it accepts no arguments. `validate` must synchronously return `true` or `false`. Empty values are skipped by default; set `skipEmpty: false` on the handler to check them. Do not mutate the data or perform asynchronous requests inside the handler.

The XLSX stores the rule's name and arguments. Its implementation stays in the application: supply the same `validationRules` when issuing and reading a form. Built-in rules cannot be replaced. Handler types and execution errors are covered in the [API](./api.md).

Next: [issue a form, fill it in Excel, and read the result](./forms.md).

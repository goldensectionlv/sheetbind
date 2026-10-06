# API

Import functions and types from `sheetbind`. Node.js `Buffer` values can be passed wherever a function accepts `Uint8Array`. See [templates](./templates.md) for XLSX tags and [fields](./fields.md) for validation and choices.

## Data and options

Execution data must be a plain JSON object. Nested objects, arrays, strings, finite numbers, booleans and `null` are accepted. Omit absent properties or use `null`; `undefined`, `Date`, functions, accessors, cyclic objects and sparse arrays are rejected. The exported `JsonValue` type describes JSON values.

The signatures below use two abbreviations. **`RunOptions` and `ReadOptions` are not exported types**; they describe these parameter shapes:

```ts
type RunOptions = ValidationOptions & {
  readonly dictionaries?: Dictionaries
}

type ReadOptions = RunOptions & {
  readonly context?: Readonly<Record<string, unknown>>
}
```

| Option | Used by | Purpose |
| --- | --- | --- |
| `dictionaries` | Resolve, render, read | Named string or object lists declared by fields |
| `validationRules` | Resolve, render, read | Custom synchronous field rules |
| `validationMessages` | Resolve, render, read | Message overrides by rule name |
| `context` | Read only | Shared choice collections referenced through `$root`; these values do not become submitted data |

`Dictionaries` is a readonly object whose values are string arrays or arrays of JSON objects. `ValidationOptions` contains `validationRules` and `validationMessages`. A custom handler uses `ValidationRule`; its predicate receives `ValidationContext` (`root`, `current`, `path`). `ValidationMessage` is a string or a function receiving `ValidationMessageContext`, which also includes `value`, `rule`, `args` and `index`. See the [complete handler example](./fields.md).

## Import a template

```ts
declare function importWorkbookXlsx(bytes: Uint8Array): Promise<WorkbookTemplate>
```

Loads a tagged XLSX and checks its tags and repeat geometry. It does not resolve application data or run field predicates.

`WorkbookTemplate` is a type for the imported result. It has no public constructor or editing methods. Reuse the result for multiple executions; retain the original XLSX for storage and import it again after editing. Use `import type { WorkbookTemplate } from 'sheetbind'` for annotations.

## Render a report

```ts
declare function renderWorkbookReport(
  template: WorkbookTemplate,
  data: unknown,
  options?: RunOptions,
): Promise<Buffer>
```

Resolves bindings, repeats and field rules, then returns the report bytes. Invalid field values fail rendering. Save the returned buffer to an `.xlsx` file. Formula calculation and retained Excel features are described in [Excel and limitations](./xlsx.md).

## Issue a form

```ts
declare function renderWorkbookForm(
  template: WorkbookTemplate,
  data: unknown,
  options?: RunOptions,
): Promise<Buffer>
```

Returns an editable XLSX with the structure needed for reading it later. It checks the template and rule configuration; required fields may be blank at issuance. Field values are validated when the completed form is read.

An empty array for a repeat containing one row of input fields produces one blank input row. Pass an array of 20 empty objects to issue 20 rows. Forms support repeats down the sheet; see [form editing and reading](./forms.md) for supported edits and multirow records.

## Read a completed form

```ts
declare function readWorkbookForm(
  template: WorkbookTemplate,
  bytes: Uint8Array,
  options?: ReadOptions,
): Promise<WorkbookFormResult>
```

Use the original template and the completed file. Reading returns fields declared by bindings, after resolving choices and omitting empty single-row records. It does not merge the result with issuance data or match records by identity.

```ts
type WorkbookFormResult =
  | { readonly success: true, readonly data: Readonly<Record<string, unknown>> }
  | { readonly success: false, readonly issues: readonly WorkbookFormIssue[] }
```

There is no partial `data` on failure. Check `result.success` and also handle exceptions from template or application configuration.

For string lists and choices returning a key, supply the issuance dictionaries again. Root collections used by key choices go in `context`. Object choices use the source embedded in the issued file. The [form guide](./forms.md) shows what to retain for reading.

## Inspect values and placement

```ts
declare function resolveWorkbook(
  template: WorkbookTemplate,
  data: unknown,
  options?: RunOptions,
): WorkbookLayout
```

Synchronously resolves and validates report data without creating a file. This is the report layout; issued forms add hidden rows for their structure.

`WorkbookLayout.sheets` contains sheets with `id`, `name`, optional `state`, `rows`, `columns`, `print`, and a `cells` array. `WorkbookCellInstance` describes each cell:

| Field | Meaning |
| --- | --- |
| `id`, `definitionId` | Execution cell ID and the imported cell's ID |
| `at` | `GridAddress`: `{ row, column }`, starting at 1 |
| `size` | `GridOffset`: `{ rows, columns }`, including a merged cell's extent |
| `value` | `{ literal: TemplateValue }` or `{ formula: string }`; formulas have no leading `=` |
| `rules` | Optional `FieldRules`: validation, messages, string list or choice |
| `origin` | `{ nodeId, dataPath, iterations: [{ nodeId, index }] }`: source node, data path and repeat indexes |
| `contextPath` | Data path of the binding context |
| `choice` | Optional `{ text, items: [{ key, label, value, text }] }`: display text and options |

`TemplateValue` is `string | number | boolean | null`. `FieldValue` also permits an object choice value. `ChoiceOption` describes an option without its Excel display text; `ResolvedChoice` describes a selected key and those options. `ValueExpression` describes a literal or data binding; the latter uses `DataReference` (`path`, optional `from: 'current' | 'root'`). These types do not provide a template construction API.

The result is readonly in TypeScript. Cells and sheet settings are copied: JavaScript changes to the result do not affect sibling cells, the template, input data or later runs. IDs describe the imported template and execution; they are not record identifiers. Original XLSX parts and authored tag expressions are not exposed in the layout.

## Find dictionary dependencies

```ts
declare function workbookDictionarySources(template: WorkbookTemplate): string[]
```

Returns sorted, unique names declared by string lists and named choices. It does not resolve data or include contextual collections such as `$root.products`. The application loads these dictionaries and passes them through `options.dictionaries`.

## Parse and format validation

```ts
declare function parseValidation(value: unknown): readonly ValidationRuleUse[]
declare function formatValidation(value: Validation): string
```

`Validation` is a pipeline string or an ordered array of `ValidationRuleUse`. Each use has a `rule` name, optional JSON `args` and an optional string `message`.

```ts
import { parseValidation, formatValidation } from 'sheetbind'

const rules = parseValidation('required|number|min:0')
const text = formatValidation(rules) // 'required|number|min:0'
```

Parsing checks syntax and normalizes the list; it does not execute rules or check whether a custom handler exists. Handler names and arguments are checked during execution preparation, including empty repeats. Formatting omits occurrence messages, so use the structured array when those messages must be retained. `ValidationIssue` describes a rule failure: `code`, `rule`, `args`, `index`, `message`.

## Get a choice range name

```ts
declare function workbookChoiceRange(rule: ChoiceRule, field?: string): string
```

Returns an Excel defined name for a **named object choice**. Omit `field` to address the displayed choices; supply a top-level object property to address its values. The name depends on the dictionary name and the choice's `key` and `label` mappings.

```ts
import { workbookChoiceRange } from 'sheetbind'
import type { ChoiceRule } from 'sheetbind'

const rule: ChoiceRule = {
  source: { dictionary: 'Products' },
  key: 'id',
  label: 'name',
}
const labels = workbookChoiceRange(rule)
const prices = workbookChoiceRange(rule, 'price')
const formula = `INDEX(${prices},MATCH(B3,${labels},0))`
```

Use the resulting formula in the source workbook. Rendering creates these ranges when an object choice using that dictionary is emitted (`return` omitted or `'object'`). The helper only computes a name; it does not create a range. A property range exists only if that property occurs in the source objects. A contextual source throws `RangeError`; key-returning choices do not create these object-property ranges.

`ChoiceRule` contains `source`, `key`, `label` and optional `return: 'object' | 'key'`. Its source is `{ dictionary: string }` or a `DataReference`.

## Errors

| Situation | Outcome |
| --- | --- |
| Invalid values or structure in a returned form | `WorkbookFormResult` with `success: false` |
| Returned file cannot be read as XLSX | The same result with `code: 'invalid-workbook'` |
| Unreadable template or invalid XLSX tag | `TaggedXlsxError` |
| Missing data or dictionaries, invalid field rules | `TemplateError`, often its `TaggedXlsxError` subclass |
| Custom predicate or message throws or returns the wrong type | `ValidationExecutionError` |
| Malformed arguments, unsupported geometry or conflicting native validation | `SyntaxError`, `TypeError` or `RangeError`, depending on the check |

`TemplateError.issues` contains `TemplateIssue` objects with `phase: 'template' | 'data'`, `code`, `path`, `nodeId` and `message`. Rule failures may add `rule`, `args` and `index`.

`TaggedXlsxError` extends `TemplateError`. Its `TaggedXlsxIssue` adds optional `sheetName` and `address` of the authored tag. Unreadable template errors retain the original failure in `cause`.

`ValidationExecutionError` exposes `rule`, `path` and `cause`. Predicates must return a synchronous boolean; message functions must return a string. A Promise is invalid in either case.

`WorkbookFormIssue` contains `phase: 'structure' | 'value' | 'xlsx'`, `code`, `path`, `message` and optional `sheetName`, `address`, `nodeId`, `rule`, `args`, `index`. Its address belongs to the returned workbook; its data path uses indexes after empty rows are omitted. Some structural failures cannot identify a cell. The [form reader](./forms.md) shows both `try/catch` and `result.success` handling.

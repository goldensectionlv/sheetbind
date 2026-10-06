# API

Import functions and types from `sheetbind`. Node.js `Buffer` values can be passed wherever a function accepts `Uint8Array`. See [templates](./templates.md) for XLSX tags and [fields](./fields.md) for validation and choices.

## Data and options

Execution data is a plain JavaScript object. Nested objects, arrays, strings, finite numbers, booleans and `null` are accepted. Object properties with `undefined` mean absence at any depth; there is no need to replace them with `null` or remove them before rendering. Forms leave missing fields blank, while required report bindings still report `missing-source`. `Date`, functions, accessors, cyclic objects, sparse arrays and `undefined` array items are rejected. The exported `JsonValue` type describes serializable JSON values, not every accepted input object.

The signatures below use two abbreviations. **`RunOptions` and `ReadOptions` are not exported types**; they describe these parameter shapes:

```ts
type RunOptions = {
  readonly dictionaries?: Dictionaries
}

type ReadOptions = ValidationOptions
```

| Option | Used by | Purpose |
| --- | --- | --- |
| `dictionaries` | Resolve, render | Named string or object lists declared by fields |
| `validationRules` | Read | Custom synchronous field rules |
| `validationMessages` | Read | Message overrides by rule name |

`Dictionaries` is a readonly object whose values are string arrays or arrays of data objects. Undefined object properties are omitted from the dictionary snapshots saved in forms; input objects are not changed. `ValidationOptions` contains `validationRules` and `validationMessages`. A custom handler uses `ValidationRule`; its predicate receives `ValidationContext` (`root`, `current`, `path`). `ValidationMessage` is a string or a function receiving `ValidationMessageContext`, which also includes `value`, `rule`, `args` and `index`. See the [complete handler example](./fields.md).

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

Resolves bindings and repeats, then returns the report bytes. Rendering does not execute `@validate` or require its handlers. It still checks template structure, data bindings and dictionary sources. Save the returned buffer to an `.xlsx` file. Formula calculation and retained Excel features are described in [Excel and limitations](./xlsx.md).

## Issue a form

```ts
declare function renderWorkbookForm(
  template: WorkbookTemplate,
  data: unknown,
  options?: RunOptions,
): Promise<Buffer>
```

Returns an editable XLSX with the structure needed for reading it later. It checks the template and dictionary sources; required fields may be blank at issuance. Validation handlers are only needed when reading the completed form.

An empty, omitted or `null` array for a repeat containing one row of input fields produces one blank input row. Pass an array of 20 empty objects to issue 20 rows. Forms support repeats down the sheet; see [form editing and reading](./forms.md) for supported edits and multirow records.

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
  | { readonly success: false, readonly issues: readonly WorkbookFormIssue[], readonly data?: Readonly<Record<string, unknown>> }
```

Value failures contain both `data` and `issues`, so the application can show the parsed fields for correction. An invalid choice remains the entered string; formulas and unsupported values produce `null` and an issue. These data have not passed validation. Unreadable XLSX or invalid structure produces no `data`. Check `result.success` and also handle template or application configuration exceptions. Types and precision follow the [submitted cells' formats](./xlsx.md#input-formats-and-value-types).

All list and choice sources are embedded in the issued file, including row-local arrays. Reading only uses `validationRules` and `validationMessages` from the options; dictionaries and context are not read inputs. Keep the original template.

An unknown validation rule is skipped with one `console.warn` per rule name per read. Known rules still run in their original order. Invalid arguments for a known rule and errors inside a supplied handler remain errors.

## Inspect a template without data

```ts
const findings = inspectWorkbookTemplate(template, {
  dictionaries: ['Products', 'Statuses'],
})
```

Checks formatter availability, built-in formatter arguments and dictionary names. Each `WorkbookTemplateFinding` has `severity` (`error` or `warning`), `code`, `message`, `nodeId`, `path` and a `sheetName`/`address` location. An unknown formatter is an error; an unknown dictionary is a warning. Omitting `dictionaries` skips dictionary-name checks. Custom formatters and validation rules are not invoked. Syntax and layout are checked earlier by `importWorkbookXlsx`.

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

`ChoiceRule` contains `source`, `key`, `label` and optional `return: 'object' | 'key'`. `emptySource: 'input'` permits free input for an empty source and requires `return: 'key'`. Its source is `{ dictionary: string }` or a `DataReference`; `.answers` can select an array from each record in both reports and forms. Omitted, `undefined` and `null` data references behave like empty arrays; named dictionaries must still be supplied. See [validation and lists](./fields.md).

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

## Registering handlers

```ts
declare function registerFormatter(name: string, formatter: Formatter): void
declare function registerValidationRule(name: string, rule: ValidationRule): void
```

Registrations belong to the loaded package instance; register once at application startup. Formatters run during rendering and layout; validators run only on read. Per-call `validationRules` override global rules. Built-in names are reserved; registering an application name again replaces its handler. `Formatter` receives `TemplateValue` and `readonly JsonValue[]` and synchronously returns `TemplateValue`. `FieldRules.format` accepts a pipeline string or `FormatterUse` array (`formatter`, optional `args`); the configuration type is `Formatting`. See [examples](./fields.md#formatting-values).

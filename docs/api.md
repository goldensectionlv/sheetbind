# API

Import functions and types from `sheetbind`. Node.js `Buffer` values can be passed wherever a function accepts `Uint8Array`. See [templates](./templates.md) for XLSX tags and [fields](./fields.md) for validation and choices.

## Data and options

Execution data is a plain JavaScript object. Nested objects, arrays, strings, finite numbers, booleans and `null` are accepted. Object properties with `undefined` mean absence at any depth. Rendering leaves missing fields blank; report bindings without `?` emit a warning. Forms allow blank fields at issuance and validate them on read. `Date`, functions, accessors, cyclic objects, sparse arrays and `undefined` array items are rejected. `JsonValue` describes serializable JSON values.

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

`Dictionaries` is a readonly object of string arrays or arrays of data objects. Only names declared by the template are read and copied; unused dictionaries cannot block rendering. Duplicate strings are removed with a warning. Missing, empty or unusable sources leave ordinary input with a warning. Undefined object properties are omitted from sources stored in forms without changing caller data. `ValidationOptions` contains `validationRules` and `validationMessages`. A `ValidationRule` receives `ValidationContext` (`root`, `current`, `path`). `ValidationMessage` is a string or a function receiving `ValidationMessageContext`, which also includes `value`, `rule`, `args` and `index`. See the [handler example](./fields.md).

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

Value failures contain both `data` and `issues`, so the application can show the parsed fields for correction. An invalid choice remains the entered string; formulas without a usable saved result and unsupported values produce `null` and an issue. These data have not passed validation. Unreadable XLSX, a damaged source payload or invalid structure produces no `data`. An absent named dictionary produces a warning and leaves the cell value unchanged; it does not stop reading. Check `result.success` and also handle template or application configuration exceptions. Types and precision follow the [submitted cells' formats](./xlsx.md#input-formats-and-value-types).

Available list and choice sources are embedded in the issued file, including row-local arrays. Reading only uses `validationRules` and `validationMessages` from the options; dictionaries and context are not read inputs. Keep the original template.

An unknown validation rule stops reading with a template error, including fields in empty repeats. Register its handler or supply it through `validationRules` for the read call. Rules run in their original order. Invalid arguments for a known rule and errors inside a supplied handler remain errors.

## Inspect a template without data

```ts
const findings = inspectWorkbookTemplate(template, {
  dictionaries: ['Products', 'Statuses'],
})
```

Checks formatter availability, built-in formatter arguments and dictionary names. Each `WorkbookTemplateFinding` has `severity`, `code`, `message`, `nodeId`, `path` and a `sheetName`/`address` location. Unavailable formatting and unknown dictionaries produce warnings. Omitting `dictionaries` skips dictionary-name checks. Handlers are not invoked. Syntax and layout are checked earlier by `importWorkbookXlsx`.

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

`TemplateValue` is `string | number | boolean | null`. `FieldValue` also permits an object choice value. `ChoiceOption` describes an option without its Excel display text. `ValueExpression` describes a literal or data binding; the latter uses `DataReference` (`path`, optional `from: 'current' | 'root'`). These types do not provide a template construction API.

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

Use the resulting formula in the source workbook and declare a choice using that dictionary in the template. Rendering creates the referenced ranges from the supplied dictionary, regardless of the choice's `return` mode or the number of repeated records. Missing properties produce blank cells; an empty dictionary produces a blank range. The helper only computes a name. Unreferenced properties stay in the saved JSON and are not written into cells. Excel precision applies to formula cells; reading the chosen object preserves the saved JSON values. A contextual source or a conflict with an existing range of the same name throws `RangeError`.

`ChoiceRule` contains `source`, `key`, `label` and optional `return: 'object' | 'key'`. Its source is `{ dictionary: string }` or a `DataReference`; `.answers` selects an array from each record. Missing, `undefined`, `null` and empty sources allow ordinary input with a warning. This applies to named, root and local sources, including object choices; without a source, reading returns the cell value rather than reconstructing an object. `emptySource: 'input'` is also accepted for key choices but is not required. Populated lists are checked on form read, not on render. See [validation and lists](./fields.md).

## Errors

Rendering resolves values, formatting and placement to produce a valid XLSX. It does not validate field values against lists or `@validate`. Reading reconstructs the returned records, resolves saved choices and validates completed fields. Recoverable rendering problems produce warnings; unreadable files, ambiguous mappings and invalid structure remain errors.

| Situation | Outcome |
| --- | --- |
| Invalid values or structure in a returned form | `WorkbookFormResult` with `success: false` |
| Returned file cannot be read as XLSX | The same result with `code: 'invalid-workbook'` |
| Unreadable template or invalid XLSX tag | `TaggedXlsxError` |
| Missing, empty or unusable render source | `console.warn`; preserve the scalar value or the declared object label and skip the list |
| Missing bound value during rendering | `console.warn` and a blank cell; `?` suppresses the warning |
| Unknown or failed formatter | `console.warn`; retain the value before the formatting pipeline |
| Unknown validator or invalid validator arguments | Template/configuration error when preparing form reading |
| Unmatched input equals a valid choice label | Form issuance stops with `ambiguous-choice`; report rendering remains available |
| Custom predicate throws or returns the wrong type | `ValidationExecutionError` |
| Validation message throws or returns the wrong type | `console.warn`; keep the validation issue with a default message |
| Conflicting native validation | `console.warn`; preserve the authored validation and omit the overlapping generated dropdown |
| Unsupported data, geometry or XLSX limits | `SyntaxError`, `TypeError`, `TemplateError` or `RangeError`, depending on the check |

`TemplateError.issues` contains `TemplateIssue` objects with `phase: 'template' | 'data'`, `code`, `path`, `nodeId` and `message`. Rule failures may add `rule`, `args` and `index`.

`TaggedXlsxError` extends `TemplateError`. Its `TaggedXlsxIssue` adds optional `sheetName` and `address` of the authored tag. Unreadable template errors retain the original failure in `cause`.

`ValidationExecutionError` exposes `rule`, `path` and `cause`. Predicates must return a synchronous boolean. Message functions must return a string; a failed message uses `failed validation: <rule>` and does not change the validation result. Unused per-read handlers and messages are ignored.

`WorkbookFormIssue` contains `phase: 'structure' | 'value' | 'xlsx'`, `code`, `path`, `message` and optional `sheetName`, `address`, `nodeId`, `rule`, `args`, `index`. Its address belongs to the returned workbook; its data path uses indexes after empty rows are omitted. Some structural failures cannot identify a cell. The [form reader](./forms.md) shows both `try/catch` and `result.success` handling.

## Registering handlers

ESM imports and CommonJS requires of the same installed package share templates, error classes and registered handlers. Both entry points use one runtime and one set of type declarations.

```ts
declare function registerFormatter(name: string, formatter: Formatter): void
declare function registerValidationRule(name: string, rule: ValidationRule): void
```

Registrations belong to the loaded package instance; register once at application startup. Formatters run during rendering and layout; validators run only on read. Per-call `validationRules` override global rules. Built-in names are reserved; duplicate registration of an application name is an error. `Formatter` receives `TemplateValue` and `readonly JsonValue[]` and synchronously returns `TemplateValue`. `FieldRules.format` accepts a pipeline string or `FormatterUse` array (`formatter`, optional `args`); the configuration type is `Formatting`. See [examples](./fields.md#formatting-values).

# Architecture

This page describes Sheetbind's implementation. For user-facing behavior, see [templates](./templates.md), [forms](./forms.md) and [Excel and limitations](./xlsx.md).

Sheetbind is one package for XLSX reports and editable forms. Its public entry point is `src/index.ts`; templates enter through `importWorkbookXlsx`.

## Modules

| Directory | Responsibility |
| --- | --- |
| `src/core/` | Data paths, values, field rules, formatters, validation and choice labels |
| `src/grid/` | Workbook definitions, XLSX coordinates and limits, repeats, placement and formula references |
| `src/xlsx/` | Template import, reports, form issuance and reading, native content preservation |

ESLint keeps value rules and geometry independent of file I/O. The main entry files are `xlsx/workbook-template.ts`, `grid/workbook-data.ts`, `grid/workbook-layout.ts`, `xlsx/workbook-form.ts` and `xlsx/workbook-source.ts`.

Form field preparation lives in `xlsx/form-definition.ts`, record operations in `xlsx/form-records.ts`, and structural boundaries in `xlsx/workbook-form-markers.ts`. `xlsx/workbook-lists.ts` writes dropdowns, dictionary ranges and saved form sources, and reads those sources back. Files for axes, formulas and native XLSX parts each own their algorithm. Internal exports are private unless exposed by `src/index.ts`.

## Data and ownership

| Representation | Contents and consumers |
| --- | --- |
| Source XLSX inside `WorkbookTemplate` | Original formatting and native Excel features; the writer preserves unaffected package content |
| `WorkbookDefinition` | Bindings, regions, geometry and required source references; used by execution, forms and diagnostics |
| `WorkbookData` | Expanded values and repeat instances with concrete data paths; shared by axes, placement and form boundaries during one call |
| `WorkbookPlan` | Sheet plans with their definition, data, placed cells, axes, regions and coordinate mapping; used by formulas, markers and writing |
| Public `WorkbookLayout` | An independent readonly view returned by `resolveWorkbook` |

The sheet definition owns its source package path. Cells retain authored addresses and expressions. Source XML may omit a cell synthesized during import, so writing handles an absent physical cell.

Each `WorkbookSheetPlan` carries a sheet and its geometry together. Definitions and axes use logical coordinates; placed cells receive final positions and formulas. XLSX mapping accounts for removed tag rows and inserted form control rows. Native ranges, formulas and form boundaries reuse these mappings.

Dictionary input is copied and validated at the operation boundary. Choice projections borrow validated records. Public layouts copy option values; form reading copies selected objects into independent fields. Missing named dictionaries warn once per name per call and leave ordinary input. Supplied sources retain their checks.

## Reports and form issuance

```text
Tagged XLSX -> import -> WorkbookTemplate
                            |
Data + dictionaries --------+
                            |
                  resolve values and repeats
                            |
                   place rows and columns
                            |
                  map formula references
                            |
              write into the source XLSX -> file
```

Import uses ExcelJS to decode the workbook, compiles tags and checks region ownership. Execution resolves values, expands repeats and prepares choice labels. Formatting runs during execution; validation handlers run when reading completed forms.

The writer opens the source ZIP and applies the plan to cells, dimensions, merges, names and form structure. The shared cell writer owns text limits, the string table and derived styles. Existing style IDs remain valid. Native validation conflicts and inherited prompts are handled at final worksheet coordinates.

The hidden list sheet stores dropdowns, dictionary columns referenced by formulas and form sources. Formula references determine the dictionary columns to write, including when repeats are empty. Identical dropdown lists share a range.

Form preparation expands object scopes into field and collection paths, checks binding ownership and collects input fields. Issuance permits blank required values. Hidden record boundaries and a control column beyond all authored content, including empty repeats, support later reading. These boundaries participate in coordinate mapping before formulas and native references are written.

## Reading a form

```text
Original template + completed XLSX
                  |
         check workbook structure
                  |
      collect fields and their locations
                  |
       decode choices and build records
                  |
        omit empty single-row records
                  |
         validate the completed data
                  |
          data and value issues
```

ExcelJS decodes the returned workbook. Markers describe record counts and boundaries; shared placement checks field positions and merges. The supplied template defines bindings and validation rules. Applications own document identity and business comparisons.

The submitted cell format determines numeric types and precision; native dates become ISO strings. Choices use sources saved in the file. Shared dictionaries and root collections are stored once; local sources belong to the issued field paths. Their snapshot is serialized before asynchronous writing. Local sources support filling and clearing issued fields; changing those records or their order requires reissuing the form.

Source checks and label decoding precede validation. A missing named dictionary leaves the entered value and emits a warning. A malformed saved payload is a file failure. Empty single-row records are omitted; invalid nonempty input keeps its row. Values, validation contexts and issues use the resulting data paths.

## Invariants when changing the code

- **Inputs belong to the caller.** Reusing a template must not change its definition, source bytes, data or dictionaries. Public layouts own independent values, rules, choices and geometry.
- **Placement has one source of coordinates.** Cells, native metadata, formulas and print references use shared axis plans plus tag and form row mapping.
- **Bounds and exact references differ.** `SourceCoordinates.range` gives enclosing bounds; `references` can split ranges. Validation and conditional formatting must not spread across unrelated gaps or hidden form boundaries. Consumers handle removed targets.
- **Preserve authored content.** Merge generated features with existing settings and retain unaffected package parts.
- **Reading follows returned records.** Keep complete record boundaries; applications decide which changes are permitted.
- **Locations belong to the XLSX boundary.** Core issues carry data paths and node IDs; the adapter adds template or submitted-cell addresses.

Keep regression cases in tests and verify the affected boundary using the [development instructions](./development.md).

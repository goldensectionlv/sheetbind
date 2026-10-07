# Architecture

This page is for changes to Sheetbind itself. For template behavior and file support, see [templates](./templates.md), [forms](./forms.md) and [Excel and limitations](./xlsx.md).

Sheetbind targets XLSX. DOCX and other formats are outside its architectural requirements. The directories below group current library tasks; they do not define a universal document core with pluggable format implementations.

## Modules

| Directory | Responsibility | Does not own |
| --- | --- | --- |
| `src/core/` | Data paths, values, field rules and choice labels | Worksheet traversal and file I/O |
| `src/grid/` | Workbook model, XLSX coordinates and limits, repeats, placement and formula references | Reading and writing package parts |
| `src/xlsx/` | Template import, reports, form issuance and reading, native content preservation | Application data loading or business workflows |

`src/index.ts` defines the public package API. An internal export does not become public automatically. ESLint separates value rules and geometry from file reading and writing. Forms belong to XLSX and have no separate format-independent layer.

Template import, storage and inspection live in `xlsx/workbook-template.ts`; form issuance and reading live in `xlsx/workbook-form.ts`. Field preparation is in `xlsx/form-definition.ts`, and record operations are in `xlsx/form-records.ts`. Collecting, saving and reading choice sources share `xlsx/workbook-choice-sources.ts`. Separate files for axes, formulas, markers and native XLSX parts correspond to distinct algorithms.

## Models and ownership

The library has three operations: rendering a report, issuing a form and reading a completed form. They share bindings and placement but differ in where values come from and which checks apply. `core` contains value rules; workbook region execution belongs to `grid`.

| Representation | Contents | Consumers |
| --- | --- | --- |
| Source XLSX inside `WorkbookTemplate` | Formatting and native Excel features | The writer retains source package parts and changes affected fragments |
| `WorkbookDefinition` | Template bindings, regions and geometry with source cell references | Data resolution, placement, form preparation and diagnostics |
| `WorkbookData` | Values and repeat instances with concrete data paths | Axis planning, placement and form boundaries; exists only for the current call |
| `WorkbookPlan` | Sheet plans containing the definition, shared repeat instances, placed cells, axes and coordinate mapping | Formulas, form control rows and the XLSX writer |

`WorkbookLayout`, returned by `resolveWorkbook`, is an independent public projection of placed cells. It is a consumer result rather than another editable template. The public entry into the current model is tagged XLSX import.

Within `WorkbookPlan`, each `WorkbookSheetPlan` retains one sheet and its geometry. Stages pass these plans together; matching a sheet with its axes and regions does not depend on indexes in parallel arrays. The original definition and axes retain logical coordinates, while placed cells receive final positions and formulas in sequence. The XLSX adapter also accounts for removed tag rows and inserted form control rows.

Placement retains the existing `WorkbookData` instances instead of constructing another region tree. Form markers traverse those instances and obtain their boundaries through the same axis mapping used for placement. Reports do not build a separate form-region layout.

The expanded body owns the iteration context shared by its cells. Placement creates public cell origins from that context directly. Source worksheet metadata remains in the definition; the placed sheet contains only output settings and cells.

The XLSX writer consumes these placed cells directly. Form input formatting is applied while writing source cell styles; it does not require another copy of the workbook or a second output-cell model. Explicit number formats from the template remain authoritative.

When simplifying, first examine representations and the transitions between them. Shared geometry serves all three operations, retained XLSX parts preserve native Excel features, and the returned form structure supports reading changed records. Removing one of these contracts requires checking every consumer; moving functions between directories does not itself reduce the number of contracts.

## Report and form issuance

```text
Tagged XLSX -> import -> WorkbookTemplate
                            |
Data + dictionaries ---------+
                            |
                  resolve values and repeats
                            |
                   place rows and columns
                            |
                  map formula references
                            |
              write into the source XLSX -> file
```

Import compiles tags into a `WorkbookDefinition`, validates region ownership and retains the source file in `WorkbookTemplate`. Formatting and native Excel features stay in the XLSX package instead of being reconstructed as a second style model.

Writing opens the source package once and applies `WorkbookPlan` directly to its XML: cells, row and column sizes, merges, names and form structure. All written cells share one string table; source style IDs stay valid, with only the required format variants appended. `workbook-lists.ts` writes the shared hidden sheet for dictionaries and saved form sources. The package is then compressed into the final file. There is no intermediate ExcelJS workbook, its serialization or a merge of two packages; ExcelJS is used for template import and form reading.

Execution resolves values and expands regions directly from the workbook definition, retaining their definitions and concrete data paths. There is no separate generic operation tree or identity translation. Field declarations and dictionary dependencies are checked before expanding repeats, including empty ones. Execution does not look up or execute validation handlers. The grid builds row and column plans; values, formulas, sheet settings and native content use those plans to find their output positions. The XLSX writer combines placed values with the retained source content.

Field rules are normalized at import, and dictionaries are validated and copied once at the operation boundary. Internal steps reuse those values without reparsing; dictionary dependencies are checked even for empty repeats.

The dropdown writer receives ready-to-write strings regardless of source kind, deduplicates identical lists and writes Excel ranges. The shared cell serializer checks text limits. Native validation conflicts and inherited prompts are handled when writing the source worksheet, at final coordinates.

Dictionary ranges for formulas come directly from declared choices and the supplied dictionaries. Formula references determine their columns; repeated cell instances and the properties present in individual records do not determine whether a range exists. Dropdowns and formula ranges share the hidden sheet but are written independently.

Form issuance uses the same placement and writing path. It prepares editable fields, permits blank required values and adds the structure needed to recognize records when reading. Its hidden rows participate in coordinate mapping before formulas and other references are written.

Form preparation expands object scopes into explicit field and collection paths. This small normalization keeps markers, saved choice sources and validation contexts tied to repeat records. It changes neither placement rules nor the source template. Scope behavior is checked through rendered and returned XLSX files, rather than by requiring a particular intermediate tree shape.

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

Reading checks the submitted structure and extracts field values with data paths and cell addresses. The submitted cell's format determines numeric types and precision, and native dates become ISO strings. Extracted fields are private state within one read operation, not a contract between a form engine and a format adapter. The result is assembled in an object owned by that call without another copy. Current records come from the returned file; they are not matched to original application records.

Markers build expanded region instances directly, together with the empty result structure. Reading does not resolve those empty values as application input or run issuance rules, dictionaries or formula rewriting. Shared placement then checks the expected boundaries and field positions, so copying only part of a record block is still rejected by the same geometry used for issuance.

All list and choice sources travel with the issued workbook. Shared dictionaries and root collections are stored once; local sources are saved for the issued field paths. No additional columns carry context between rows. Local lists support filling the issued records; changing their membership or order requires issuing a new form. Reading does not reload application dictionaries, and the return mode only selects an object or its key.

Form preparation contains the definition and fields; dictionaries belong to a particular issuance or returned file. Choice resolution receives the source array directly. Issuance selects it from data or a dictionary; reading selects it from the field's saved source. Reading checks the source even for blank or invalid input, then decodes the field value immediately.

Template import and form reading use one XLSX decoder. Writers account for authored resources before allocating generated names. Choice objects are stored independently of formula ranges: those ranges are created only for references found in the template's formulas. A damaged saved source is a file failure, not an invalid user value.

Choice decoding, empty-row handling and field validation share the final data paths. An invalid nonempty input keeps its row. Issues retain their field location while their array indexes follow the same remapping as the result.

## Invariants when changing the code

- **Inputs belong to the caller.** Reusing a template must not change its definition, source bytes, supplied data or dictionaries. Public layouts contain independent cell values, rules, choices and geometry.
- **Placement has one source of coordinates.** Cells, native metadata, formulas and print references use the shared row and column plans plus tag/form row mapping. A writer must not calculate a competing set of repeat offsets.
- **A bounding range and an exact set of cells are different results.** `SourceCoordinates.range` gives enclosing bounds; `references` can split into separate ranges. Validation and conditional formatting must not spread across unrelated gaps or hidden form boundaries. Removed targets must be handled by their consumer.
- **Native content has an owner even without a cell value.** Notes, validation, conditional formatting and drawing anchors can occupy otherwise empty cells. Their region determines whether they move, repeat or disappear.
- **Form structure and values are separate.** Uploaded values cannot redefine a region. Validation runs on the completed submission; comparison with stored business records belongs to the application.
- **Locations are added at the format boundary.** Core issues contain data paths and node IDs. The XLSX adapter adds authored addresses to template errors and returned-file addresses to form issues.

The relevant entry points are `xlsx/workbook-template.ts`, `grid/workbook-data.ts`, `grid/workbook-layout.ts`, `xlsx/workbook-form.ts` and `xlsx/workbook-source.ts`. Keep concrete regression cases in tests; use the [development checks](./development.md) to verify the affected boundary.

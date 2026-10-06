# Architecture

This page is for changes to Sheetbind itself. For template behavior and file support, see [templates](./templates.md), [forms](./forms.md) and [Excel and limitations](./xlsx.md).

## Modules

| Directory | Responsibility | Does not own |
| --- | --- | --- |
| `src/core/` | Data references, contexts, repeats, field rules and choices | Cells, XLSX and file I/O |
| `src/grid/` | Workbook definitions, rectangular regions, placement and formula references | XLSX serialization and application state |
| `src/form/` | Input fields, submitted records, choice decoding and validation | Reading or writing Excel cells |
| `src/xlsx/` | Tag import, source XLSX content, file writing and form structure in the workbook | Application data loading or business workflows |

`src/index.ts` defines the public package API. An internal export does not become public automatically. ESLint enforces the main dependency boundaries: core is independent of grid, form and XLSX; grid and form do not import the XLSX adapter.

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

Execution checks the field declarations and dictionary dependencies before expanding repeats, including empty ones. It does not look up or execute validation handlers. The grid builds row and column plans; values, formulas, sheet settings and native content use those plans to find their output positions. The XLSX writer combines placed values with the retained source content.

Form issuance uses the same placement and writing path. It prepares editable fields, permits blank required values and adds the structure needed to recognize records when reading. Its hidden rows participate in coordinate mapping before formulas and other references are written.

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

The XLSX adapter checks the submitted structure and extracts field values with data paths and cell addresses. The submitted cell's format also determines numeric types and precision here, and native dates become ISO strings. Form processing works with this submission, without inspecting ExcelJS cells. It derives the current records from the returned file; it does not match them to original application records.

All list and choice sources travel with the issued workbook. Shared dictionaries and root collections are stored once; row-local sources use hidden references that move when whole rows or record blocks are copied or sorted. The XLSX adapter resolves those references before passing choice contexts to form processing. Reading does not reload application dictionaries, and the return mode only selects an object or its key.

Choice decoding, empty-row handling and field validation share the final data paths. An invalid nonempty input keeps its row. Issues retain their field location while their array indexes follow the same remapping as the result.

## Invariants when changing the code

- **Inputs belong to the caller.** Reusing a template must not change its definition, source bytes, supplied data or dictionaries. Public layouts contain independent cell values, rules, choices and geometry.
- **Placement has one source of coordinates.** Cells, native metadata, formulas and print references use the shared row and column plans plus tag/form row mapping. A writer must not calculate a competing set of repeat offsets.
- **A bounding range and an exact set of cells are different results.** `SourceCoordinates.range` gives enclosing bounds; `references` can split into separate ranges. Validation and conditional formatting must not spread across unrelated gaps or hidden form boundaries. Removed targets must be handled by their consumer.
- **Native content has an owner even without a cell value.** Notes, validation, conditional formatting and drawing anchors can occupy otherwise empty cells. Their region determines whether they move, repeat or disappear.
- **Form structure and values are separate.** Uploaded values cannot redefine a region. Validation runs on the completed submission; comparison with stored business records belongs to the application.
- **Locations are added at the format boundary.** Core issues contain data paths and node IDs. The XLSX adapter adds authored addresses to template errors and returned-file addresses to form issues.

The relevant entry points are `xlsx/workbook-template.ts`, `grid/workbook-layout.ts`, `xlsx/workbook-form.ts`, `form/workbook-read.ts` and `xlsx/workbook-package.ts`. Keep concrete regression cases in tests; use the [development checks](./development.md) to verify the affected boundary.

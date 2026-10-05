# Changelog

## Unreleased (0.2.0)

### Fixed

- Named ranges and native formula references become `#REF!` when their entire range disappears with a repeat, including whole-row and whole-column references. Excel can no longer normalize reversed output bounds onto unrelated cells.
- Filters and sorts are removed when their entire range disappears with a repeat. Vanished sort keys are omitted, and a sort state with no remaining keys is removed without dropping a surviving filter.
- Worksheet selections and viewport anchors remain valid when selected template rows or columns repeat or disappear.
- Malformed form control tuples are rejected as `invalid-marker` before interpreting their structure; valid `sheetbind.form/3` files retain their encoding and signature.

### Changed

- Tagged XLSX is the template storage format, carrying layout, styles, bindings and repeats.
- The `sheetbind` entry exports template import, layout resolution, reports and forms.
- `WorkbookTemplate` is an opaque imported XLSX handle; application data and dictionaries are execution inputs.
- Field validation uses ordered rules and tagged directives; earlier field options are unsupported.
- The standalone validation-message directive is `@validationMessage`; `@message` is no longer supported. Inline rule options still use `message`.
- Documentation and examples describe the current model directly.

- Form reading derives the current records from region and block boundaries. One-row repeats support insertion, deletion and sorting without record policies; empty collections provide one blank input line.
- Source Excel formatting and page settings pass through without a portable style projection.
- `resolveWorkbook` returns explicitly defined public layout fields without source XLSX metadata. Returned cells and sheet settings are independent of other cells, inputs and subsequent executions.

### Removed

- Form record policies, issued-key ledgers, new-record defaults and change tracking. Applications own comparison with their stored data.
- Writable workbook models, `exportWorkbookXlsx`, template tag writers and style-editing projections.
- Alpha template compilation, ad hoc JSON layout, `DocumentConfig` forms and the `sheetbind/v2` entry points.
- Old request and invoice demos, parallel editors and their duplicate tests.
- Standalone JSON templates/projects, their version wrappers and public import/parsing APIs. No project-format migration or compatibility aliases remain.
- `WorkbookSession`, editing commands, selection and undo/redo, together with the `sheetbind/authoring` entry, editor examples and browser verification dependencies.
- Pipe-option and URL-encoded choice tags, field-rule migration and `InsertPreset`.
- The general-purpose `writeXlsxBuffer` export, native-table serialization repairs and their dedicated tests. Native tables are outside the workbook template and form contracts.

The removed code and demos remain available in Git history. This is a breaking minor release line; it has not been published by this change.

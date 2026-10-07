# Changelog

## Unreleased

### Added

- Template formatting pipelines with built-in `float`, `bool_replace` and `format_date`, plus application formatters registered through `registerFormatter`.
- Global validation registration through `registerValidationRule`. Per-read `validationRules` override global rules; validation still runs only on read.
- Row-local choice sources such as `.answers` work when filling the issued rows. Changing their membership or order requires issuing a new form.

### Changed

- Removed the unused `ResolvedChoice` type export; public layouts expose their actual cell choice shape.
- ESM and CommonJS entry points share one runtime and one type declaration, so templates and registered handlers work across both loaders.
- Missing, empty and unavailable list/choice sources leave ordinary input with a warning, consistently across named, root and local sources. No `emptySource` option is needed. Unused named dictionaries are ignored and repeated strings are deduplicated.
- Rendering preserves values outside lists; list membership and field validation run only when reading a completed form. Missing report values leave blank cells with a warning, suppressed by optional bindings. Unavailable or failed formatting retains the original value before the pipeline.
- Existing Excel validation is preserved when it conflicts with a generated dropdown; only the overlapping generated validation is skipped, with a warning.
- Unused per-read validators and messages are ignored. Failed message formatting preserves the validation issue with a default message and a warning; predicate execution failures remain errors.
- Runtime data and dictionaries accept undefined object properties at any depth as absent values. Saved choice-source snapshots omit those properties without changing caller data. Stored JSON definitions remain strict; undefined array items remain invalid.
- Omitted and `null` repeat sources behave like empty arrays, including nested and column repeats. Forms retain the same blank input row as for an explicit empty array. Other non-array values and non-object items remain errors.
- Forms no longer insert a first row or store a template hash. Reading checks actual record boundaries and merged fields against the supplied template; application code owns template identity/version checks. Reissue forms produced by earlier versions: the form carrier is now version 4.
- Rendering and layout resolution no longer execute `@validate` or require its handlers. Supply `validationRules` and `validationMessages` only when reading a form.
- Unknown validation rules stop reading with a template error, including fields in empty repeats. Rules retain their order and original indexes; invalid known arguments and handler failures remain errors.
- Form reading uses the dictionaries embedded at issuance for lists, key choices and object choices. `readWorkbookForm` no longer accepts `dictionaries` or `context`; validation handlers and messages remain supported. Reissue older forms whose list or key sources were not embedded.

## 0.2.0 — 2026-10-06

### Added

- XLSX templates with data bindings, nested regions and repeats down rows or across columns.
- Report generation and editable forms with reading against the original template.
- Field validation, custom rules, messages, string lists and object choices.
- Form reading after inserting, deleting or reordering records. Single-row repeats with an empty collection provide one blank input row.
- Layout inspection through `resolveWorkbook`, with independent copies of output cells and sheet settings.
- Russian and English guides with downloadable templates, scripts and completed examples.

### Fixed

- Named ranges and native formula references become `#REF!` when their entire range disappears with a repeat, including whole-row and whole-column references. Excel can no longer normalize reversed output bounds onto unrelated cells.
- Filters and sorts are removed when their entire range disappears with a repeat. Vanished sort keys are omitted, and a sort state with no remaining keys is removed without dropping a surviving filter.
- Worksheet selections and viewport anchors remain valid when selected template rows or columns repeat or disappear.
- Malformed form control markers are rejected as `invalid-marker` when reading a completed form.
- `resolveWorkbook` rejects non-JSON input data consistently with report and form generation.

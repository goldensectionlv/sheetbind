# Changelog

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

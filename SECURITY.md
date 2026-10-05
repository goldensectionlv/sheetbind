# Security policy

[English](./SECURITY.md) | [Русский](./SECURITY.ru.md)

## Supported versions

Before the first stable release, security fixes target the latest published
`0.x` version and the current `main` branch. Older `0.x` minor lines are not
maintained separately.

## Reporting a vulnerability

Do not publish technical details, an exploit, a malicious workbook, or personal
data in a normal issue.

Use the repository host's private vulnerability-reporting channel. If it is not
available, open a public issue asking for a private contact and include no
technical details.

Include privately:

- affected versions;
- a minimal reproduction with synthetic data;
- expected impact;
- known mitigations.

Acknowledgement, remediation, and disclosure timing will be coordinated after
initial review. Please allow maintainers to prepare a fix before disclosure.

## Trust boundary

Sheetbind processes ZIP/XML content through ExcelJS and JSZip. It is not a
sandbox for untrusted workbooks. The consuming application must limit input
size, processing time, and process access to resources.

Sheetbind does not execute VBA macros. A workbook containing macros, external
links, formulas, or embedded objects is not automatically safe for its recipient.

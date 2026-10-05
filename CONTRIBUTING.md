# Contributing

[English](./CONTRIBUTING.md) | [Русский](./CONTRIBUTING.ru.md)

Use Node.js and pnpm versions from `package.json`. Install with `pnpm install --frozen-lockfile`; run `pnpm check:release` before submitting changes to package exports or XLSX behavior.

The public entry is `sheetbind`. Keep bindings and validation in `src/core`, the workbook model and layout in `src/grid`, form records in `src/form`, and file transport in `src/xlsx`.

Use neutral synthetic examples. Check XLSX after export and after an independent consumer opens it; report rendering alone does not verify form reading. Keep [English](./docs/development.md) and [Russian](./docs/ru/development.md) documentation synchronized when behavior changes. Add breaking API changes to `CHANGELOG.md`.

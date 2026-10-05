# Development

Run these commands from the repository root. Use the Node.js and pnpm versions declared in `package.json` and keep the existing lockfile.

## Prepare the checkout

```sh
pnpm install --frozen-lockfile
pnpm test
```

The code is organized as described in [architecture](./architecture.md). Public exports are defined in `src/index.ts`; examples and regression tests normally use that entry point.

## Commands

| Command | Purpose |
| --- | --- |
| `pnpm test` | Run the test suite |
| `pnpm test:watch` | Run tests while editing |
| `pnpm lint` / `pnpm lint:fix` | Check code style / apply automatic fixes |
| `pnpm typecheck` | Check library, tests, examples and scripts with TypeScript |
| `pnpm check:package` | Build, pack and verify the installed archive in independent ESM and CommonJS consumers |
| `pnpm check:library` | Run lint, type checking, tests and package verification |
| `pnpm check:release` | Run library checks plus documentation checks and build |
| `pnpm docs:dev` | Start the documentation site locally |
| `pnpm docs:check` | Check links and Russian/English page structure |
| `pnpm docs:build` | Build the documentation site |
| `pnpm docs:preview` | Preview the built site |
| `pnpm example:report` | Render the package example into `temp/package/` |
| `pnpm example:tutorials` | Create and check the tutorial workbooks in `temp/tutorials/` |

The docs commands `dev`, `check` and `build` run `docs:prepare` first. It prepares downloadable examples in `docs/public/examples/`; edit their sources in `examples/`, not the generated copies.

## Examples and fixtures

| Location | Purpose |
| --- | --- |
| `examples/tutorials/` | Runnable scripts and workbook generation for the guide |
| `examples/package/` | Report rendering and form reading through the installed package |
| `examples/regions/`, `examples/comparison/` | Nested regions and column repeats |
| `examples/choices/`, `examples/validation/` | Object choices and custom field rules |
| `test/` | Regression scenarios and focused algorithm tests |
| `scripts/verify-package.ts` | Archive installation, runtime exports, public types and consumer examples |

`example:tutorials` produces reports with zero, one and several records, a form with 20 blank rows, completed and invalid forms, and multirow contact records. Generated local output belongs in ignored `temp/`.

XLSX fixtures are inputs to regression tests. Regenerate an affected fixture only when its authored definition is intentionally changing. A documentation edit does not call for rebuilding every fixture.

## Code conventions

TypeScript runs in strict mode, including unused-local and unused-parameter checks. Library source disallows explicit `any`: receive external values as `unknown` and validate them. Declare public results and contracts between modules explicitly.

ESLint defines the formatting rules. The common style uses two spaces, single quotes, no semicolons and multiline blocks. Use `pnpm lint:fix` for mechanical formatting; split functions according to responsibility and data flow.

Keep Russian and English pages at matching paths with matching section structure. Code, JSON and workbook examples are shared, use neutral English data and have one source.

## Verify a change

Start with the affected tests and a complete example through the public API. Internal imports belong in focused tests of an algorithm. Before a release, run `pnpm check:release`.

| Changed behavior | Required evidence |
| --- | --- |
| Data resolution or validation | Correct values and errors, missing data, empty repeats and repeated template use |
| Placement or metadata coordinates | Zero, one and several records; affected axes, nested or neighbouring regions; references in the saved XLSX |
| Form issuance or reading | Issue a file, edit and save it separately, then read with the original template; check data and cell addresses in errors |
| Public API or types | Install the actual archive; check ESM/CommonJS runtime usage and accepted/rejected TypeScript usage |
| Documentation | Follow the affected instructions, check downloads, run `docs:check` and `docs:build` |

Opening a file with the same library that wrote it is not enough evidence for native Excel behavior. When a change affects formulas, validation, drawings, print settings or editable blocks, also open and save the relevant files in the intended spreadsheet application. Record what was checked; package verification does not perform that external step.

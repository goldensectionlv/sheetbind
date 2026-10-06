import assert from 'node:assert/strict'
import { execFileSync } from 'node:child_process'
import { copyFile, mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

interface PackResult {
  filename: string
  files: { path: string }[]
  name: string
  version: string
}

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const npmCli = path.join(path.dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js')

function run(command: string, args: string[], cwd: string, capture = false): string {
  const result = execFileSync(command, args, {
    cwd, encoding: 'utf8', env: { ...process.env, NO_COLOR: '1' },
    stdio: capture ? ['ignore', 'pipe', 'inherit'] : 'inherit',
  })
  return typeof result === 'string' ? result : ''
}

function npm(args: string[], cwd: string, capture = false): string {
  return process.platform === 'win32'
    ? run(process.execPath, [npmCli, ...args], cwd, capture)
    : run('npm', args, cwd, capture)
}

async function main(): Promise<void> {
  const workspace = await mkdtemp(path.join(os.tmpdir(), 'sheetbind-package-'))
  console.log(`Package verification workspace: ${workspace}`)
  try {
    const output = npm(['pack', '--ignore-scripts', '--json', '--pack-destination', workspace], root, true)
    const packs = JSON.parse(output.slice(Math.max(0, output.lastIndexOf('\n[') + 1))) as PackResult[]
    assert.equal(packs.length, 1)
    const pack = packs[0]
    const manifest = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8')) as { version: string }
    assert.equal(pack.name, 'sheetbind')
    assert.equal(pack.version, manifest.version)
    const files = pack.files.map(file => file.path)
    for (const required of ['package.json', 'LICENSE', 'README.md', 'README.ru.md', 'dist/index.js', 'dist/index.cjs', 'dist/index.d.ts', 'dist/index.d.cts']) {
      assert(files.includes(required), `Missing package file: ${required}`)
    }
    for (const file of files) {
      assert(/^(?:dist\/|package\.json$|LICENSE$|README(?:\.ru)?\.md$)/.test(file), `Unexpected package file: ${file}`)
    }

    const consumer = path.join(workspace, 'consumer')
    await mkdir(consumer)
    await writeFile(path.join(consumer, 'package.json'), JSON.stringify({ name: 'sheetbind-consumer', private: true, type: 'module' }))
    npm(['install', '--ignore-scripts', '--no-audit', '--no-fund', '--package-lock=false', path.join(workspace, pack.filename), 'exceljs@4.4.0', 'typescript@5.9.3', '@types/node@22.20.1'], consumer)
    const installed = JSON.parse(await readFile(path.join(consumer, 'node_modules', 'sheetbind', 'package.json'), 'utf8')) as { version: string, exports: Record<string, unknown> }
    assert.equal(installed.version, pack.version)
    assert.deepEqual(Object.keys(installed.exports).sort(), ['.', './package.json'])

    const examples = {
      package: ['verify.ts', 'validation.ts', 'template.xlsx', 'template.data.json'],
      validation: ['definition.ts', 'rules.ts', 'template.xlsx', 'template.data.json'],
      tutorials: (await readdir(path.join(root, 'examples', 'tutorials'))).filter(name => name.endsWith('.ts')),
      walkthroughs: (await readdir(path.join(root, 'examples', 'walkthroughs'))).filter(name => /\.(ts|json)$/.test(name)),
    }
    for (const [directory, names] of Object.entries(examples)) {
      await mkdir(path.join(consumer, directory))
      for (const name of names) {
        await copyFile(path.join(root, 'examples', directory, name), path.join(consumer, directory, name))
      }
    }
    await writeFile(path.join(consumer, 'types.ts'), `
  import { renderWorkbookReport, resolveWorkbook } from 'sheetbind'
  import type { WorkbookTemplate, WorkbookFormResult, WorkbookLayout, WorkbookCellInstance } from 'sheetbind'
  declare const template: WorkbookTemplate
  const layout = resolveWorkbook(template, {}) satisfies WorkbookLayout
  const cell = layout.sheets[0].cells[0] satisfies WorkbookCellInstance
  const position: number = cell.at.row
  void position
  // @ts-expect-error source package details are private to the imported template
  void layout.sheets[0].xlsx
  // @ts-expect-error public cells do not expose source addresses or tag expressions
  void cell.xlsx
  // @ts-expect-error the resolved layout is a readonly view
  cell.at.row = 5
  void renderWorkbookReport(template, {})
  declare const result: WorkbookFormResult
  if (result.success) {
    void result.data
    // @ts-expect-error no issues on success
    void result.issues
  }
  // @ts-expect-error a template must be imported from XLSX
  resolveWorkbook({ sheets: [] }, {})
  // @ts-expect-error imported templates do not expose a mutable workbook model
  void template.sheets
  // @ts-expect-error report data is required
  void renderWorkbookReport(template)
  `)
    const opaqueType = `
  import { WorkbookTemplate } from 'sheetbind'
  declare const template: WorkbookTemplate
  void template
  // @ts-expect-error WorkbookTemplate is an opaque type, not a runtime constructor
  new WorkbookTemplate({}, new Uint8Array())
  // @ts-expect-error internal content is not a public runtime API
  WorkbookTemplate.content(template)
  `
    await writeFile(path.join(consumer, 'opaque-type.ts'), opaqueType)
    await writeFile(path.join(consumer, 'opaque-type.cts'), opaqueType)
    await writeFile(path.join(consumer, 'tsconfig.json'), JSON.stringify({
      compilerOptions: { module: 'NodeNext', moduleResolution: 'NodeNext', target: 'ES2023', strict: true, esModuleInterop: true, resolveJsonModule: true, rootDir: '.', outDir: 'build' },
      include: ['package/verify.ts', 'tutorials/*.ts', 'walkthroughs/*.ts', 'types.ts', 'opaque-type.ts', 'opaque-type.cts'],
    }))
    run(process.execPath, ['node_modules/typescript/bin/tsc'], consumer)
    for (const [directory, names] of Object.entries(examples)) {
      for (const name of names.filter(name => name.endsWith('.xlsx'))) {
        await copyFile(path.join(consumer, directory, name), path.join(consumer, 'build', directory, name))
      }
    }
    run(process.execPath, ['build/package/verify.js', 'artifacts'], consumer)
    run(process.execPath, ['build/tutorials/run.js', 'artifacts/tutorials'], consumer)
    run(process.execPath, ['build/tutorials/read-form.js', 'artifacts/tutorials'], consumer)
    // Run the exact downloadable scripts from the directory used by the reader.
    const tutorialDirectory = path.join(consumer, 'artifacts', 'tutorials')
    function tutorial(name: string, capture = false): string {
      return run(process.execPath, [path.join(consumer, 'build', 'tutorials', `${name}.js`)], tutorialDirectory, capture)
    }
    tutorial('first-report')
    tutorial('report')
    tutorial('rules-form')
    assert.deepEqual(JSON.parse(tutorial('formatting', true)), { success: true, data: { date: '01.10.2026', enabled: 'No', amount: 12.5, code: 'AB' } })
    tutorial('list-report')
    tutorial('choice-report')
    tutorial('first-form')
    await copyFile(path.join(tutorialDirectory, 'completed.xlsx'), path.join(tutorialDirectory, 'table-completed.xlsx'))
    await copyFile(path.join(tutorialDirectory, 'first-form-completed.xlsx'), path.join(tutorialDirectory, 'completed.xlsx'))
    assert.deepEqual(JSON.parse(tutorial('read-first-form', true)), { customer: { name: 'Alex' } })
    tutorial('issue-form')
    await copyFile(path.join(tutorialDirectory, 'table-completed.xlsx'), path.join(tutorialDirectory, 'completed.xlsx'))
    assert.deepEqual(JSON.parse(tutorial('read-completed-form', true)), JSON.parse(await readFile(path.join(tutorialDirectory, 'completed.json'), 'utf8')))
    await copyFile(path.join(tutorialDirectory, 'contacts-issued.xlsx'), path.join(tutorialDirectory, 'contacts-completed.xlsx'))
    const contacts = JSON.parse(tutorial('read-contacts', true)) as { contacts: unknown[] }
    assert.equal(contacts.contacts.length, 2)
    run(process.execPath, ['build/walkthroughs/generate.js', 'artifacts/walkthroughs'], consumer)
    const walkthroughDirectory = path.join(consumer, 'artifacts', 'walkthroughs')
    for (const name of ['budget', 'study-plan', 'registration-issue']) {
      run(process.execPath, [path.join(consumer, 'build', 'walkthroughs', `${name}.js`)], walkthroughDirectory)
    }
    const registration = run(process.execPath, [path.join(consumer, 'build', 'walkthroughs', 'registration-read.js')], walkthroughDirectory, true)
    assert.deepEqual(JSON.parse(registration), JSON.parse(await readFile(path.join(walkthroughDirectory, 'registration-completed.json'), 'utf8')))
    await writeFile(path.join(consumer, 'smoke.cjs'), `
  const assert = require('node:assert/strict')
  const api = require('sheetbind')
  assert.equal(Object.hasOwn(api, 'WorkbookTemplate'), false)
  for (const name of ['importWorkbookXlsx', 'renderWorkbookReport', 'renderWorkbookForm', 'readWorkbookForm', 'resolveWorkbook', 'registerFormatter', 'registerValidationRule']) {
    assert.equal(typeof api[name], 'function')
  }
  async function main() {
    const ExcelJS = require('exceljs')
    const book = new ExcelJS.Workbook()
    book.addWorksheet('Input').getCell('A1').value = '{v | twice}{@validate:four}'
    api.registerFormatter('twice', value => Number(value) * 2)
    api.registerValidationRule('four', { validate: value => value === 4 })
    const template = await api.importWorkbookXlsx(await book.xlsx.writeBuffer())
    const file = await api.renderWorkbookForm(template, { v: 2 })
    assert.deepEqual(await api.readWorkbookForm(template, file), { success: true, data: { v: 4 } })
    console.log('Installed CommonJS: registered formatters and validation: ok')
  }
  main().catch(error => { console.error(error); process.exitCode = 1 })
  `)
    run(process.execPath, ['smoke.cjs'], consumer)
    await writeFile(path.join(consumer, 'smoke.mjs'), `
  import assert from 'node:assert/strict'
  import * as api from 'sheetbind'
  assert.equal(Object.hasOwn(api, 'WorkbookTemplate'), false)
  `)
    run(process.execPath, ['smoke.mjs'], consumer)
    console.log(`Package verification: ${files.length} files, installed ESM/CJS/types: ok`)
  }
  finally {
    const directory = path.resolve(workspace)
    assert.equal(path.dirname(directory), path.resolve(os.tmpdir()), 'Unsafe package cleanup directory')
    assert(path.basename(directory).startsWith('sheetbind-package-'), 'Unsafe package cleanup name')
    if (process.env.SHEETBIND_KEEP_PACKAGE_TEMP === '1') {
      console.log(`Kept package verification workspace: ${directory}`)
    }
    else {
      await rm(directory, { force: true, recursive: true })
    }
  }
}

await main()

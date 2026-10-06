import { expect, it, vi } from 'vitest'
import { inspectWorkbookTemplate, registerFormatter, registerValidationRule } from '../src/index'
import { importAuthoredWorkbook } from './xlsx'

it('locates missing render dependencies inside empty repeats without executing handlers', async () => {
  const formatter = vi.fn(value => value)
  const validate = vi.fn(() => false)
  registerFormatter('inspectOnly', formatter)
  registerValidationRule('inspectOnly', { validate })
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').addRows([
    ['{#items}'],
    ['{.value | inspectOnly}{@validate:inspectOnly}', '{.bad | unknownInspectionFormatter}', '{.choice}{@list:Options}', '{.valid}{@choice:Known; key=id; label=name}'],
    [null, null, null, '{/items}'],
  ]))
  expect(inspectWorkbookTemplate(template, { dictionaries: ['Known'] })).toMatchObject([
    { code: 'invalid-format', severity: 'error', sheetName: 'Input', address: 'B2' },
    { code: 'unknown-dict', severity: 'warning', sheetName: 'Input', address: 'C2' },
  ])
  expect(formatter).not.toHaveBeenCalled()
  expect(validate).not.toHaveBeenCalled()
})

it('checks builtin arguments before rendering and only checks dictionaries when names are supplied', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').addRows([
    ['{value | bool_replace:"Yes"}', '{status}{@list:External}'],
  ]))
  expect(inspectWorkbookTemplate(template)).toMatchObject([{ severity: 'error', address: 'A1', message: 'bool_replace requires two text labels' }])
})

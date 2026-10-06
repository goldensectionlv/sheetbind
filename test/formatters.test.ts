import { expect, it, vi } from 'vitest'
import { readWorkbookForm, registerFormatter, registerValidationRule, renderWorkbookForm, renderWorkbookReport } from '../src/index'
import type { Formatter } from '../src/index'
import { importAuthoredWorkbook, openWorkbook } from './xlsx'

it('formats date, boolean and numeric text in reports and forms without changing the payload', async () => {
  const template = await importAuthoredWorkbook(book => {
    const sheet = book.addWorksheet('Data')
    sheet.addRows([
      ['{date | format_date:DD.MM.YYYY}', '{enabled}\n{@format:bool_replace:"Yes","No"}', '{price | float}', '{code}', '{zero | float}', '{blank | float}'],
    ])
    sheet.getCell('C1').numFmt = '#,##0.00'
  })
  const input = { date: '2026-10-01', enabled: false, price: '12.50', code: '006', zero: '0', blank: '' }
  const snapshot = structuredClone(input)
  for (const render of [renderWorkbookReport, renderWorkbookForm]) {
    const bytes = await render(template, input)
    const book = await openWorkbook(bytes)
    expect(['A1', 'B1', 'C1', 'D1', 'E1', 'F1'].map(address => book.worksheets[0].getCell(address).value)).toEqual(['01.10.2026', 'No', 12.5, '006', 0, ''])
    expect(book.worksheets[0].getCell('C1').numFmt).toBe('#,##0.00')
    if (render === renderWorkbookForm) {
      expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data: { date: '01.10.2026', enabled: 'No', price: 12.5, code: '006', zero: 0, blank: null } })
    }
  }
  expect(input).toEqual(snapshot)
})

it('registers custom formatter pipelines once, retains quoted arguments and never runs them on read', async () => {
  const custom = vi.fn<Formatter>((value, args) => String(value) + String(args[0]))
  registerFormatter('testSuffix', custom)
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Data').getCell('A1').value = '{v | float | testSuffix:" €,;:|{}"}')
  const bytes = await renderWorkbookForm(template, { v: '12.5' })
  expect(custom).toHaveBeenCalledWith(12.5, [' €,;:|{}'])
  custom.mockClear()
  expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data: { v: '12.5 €,;:|{}' } })
  expect(custom).not.toHaveBeenCalled()
})

it('rejects unknown formatters even in empty repeats and rejects invalid formatter results', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Data').addRows([
    ['{#items}'], ['{.v | absentFormatter}'], ['{/items}'],
  ]))
  await expect(renderWorkbookReport(template, { items: [] })).rejects.toThrow('Unknown formatter')
  registerFormatter('testBadResult', (() => ({ invalid: true })) as unknown as Formatter)
  const invalid = await importAuthoredWorkbook(book => book.addWorksheet('Data').getCell('A1').value = '{v | testBadResult}')
  await expect(renderWorkbookForm(invalid, { v: 1 })).rejects.toThrow('synchronous finite scalar')
  expect(() => registerFormatter('float', value => value)).toThrow('reserved')
  expect(() => registerFormatter('__proto__', value => value)).toThrow('reserved')
})

it('preserves null and invalid date text and accepts numeric timestamps', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Data').addRows([
    ['{#items}'], ['{.value | format_date:YYYY-MM-DD}'], ['{/items}'],
  ]))
  const output = await openWorkbook(await renderWorkbookReport(template, { items: [{ value: null }, { value: 'invalid' }, { value: Date.UTC(2026, 9, 1, 12) }] }))
  expect([1, 2, 3].map(row => output.worksheets[0].getCell(row, 1).value)).toEqual([null, 'invalid', '2026-10-01'])
})

it('registers validation globally but runs it only on read, with per-call overrides', async () => {
  const validate = vi.fn(value => value === 'accepted')
  registerValidationRule('testRegistered', { validate, message: 'Expected accepted' })
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Data').getCell('A1').value = '{v}{@validate:testRegistered}')
  const bytes = await renderWorkbookForm(template, { v: 'rejected' })
  await renderWorkbookReport(template, { v: 'rejected' })
  expect(validate).not.toHaveBeenCalled()
  expect(await readWorkbookForm(template, bytes)).toMatchObject({ success: false, issues: [{ phase: 'value', rule: 'testRegistered', message: 'Expected accepted' }] })
  expect(await readWorkbookForm(template, bytes, { validationRules: { testRegistered: { validate: () => true } } })).toEqual({ success: true, data: { v: 'rejected' } })
  expect(() => registerValidationRule('required', { validate: () => true })).toThrow('reserved')
})

it('reads an issued form with updated field paths and validation instead of requiring a matching hash', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Data').getCell('A1').value = '{old}{@validate:min:1}')
  const bytes = await renderWorkbookForm(template, { old: 2 })
  const changed = await openWorkbook(bytes)
  expect(changed.worksheets[0].getCell('A1').value).toBe(2)
  const source = await importAuthoredWorkbook(book => book.addWorksheet('Data').getCell('A1').value = '{current}{@validate:min:3}')
  expect(await readWorkbookForm(source, bytes)).toMatchObject({ success: false, issues: [{ phase: 'value', path: '$data.current', code: 'min' }] })
})

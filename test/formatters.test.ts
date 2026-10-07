import { expect, it, vi } from 'vitest'
import { readWorkbookForm, registerFormatter, renderWorkbookForm, renderWorkbookReport } from '../src/index'
import type { Formatter } from '../src/index'
import { importAuthoredWorkbook, openWorkbook } from './xlsx'

it.each(['bool_replace:"Yes","No"', 'float', 'format_date:YYYY-MM-DD'])('preserves blank input rows with %s', async format => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').addRows([
    ['{#items}'], ['{.value | ' + format + '}{@validate:required}'], ['{/items}'],
  ]))
  for (const input of [{}, { items: [{}] }, { items: [{ value: null }, { value: '' }, { value: ' ' }] }]) {
    expect(await readWorkbookForm(template, await renderWorkbookForm(template, input))).toEqual({ success: true, data: { items: [] } })
  }
})

it('does not hide a missing required boolean answer in a partially filled row', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').addRows([
    ['{#items}'], ['{.name}', '{.enabled | bool_replace:"Yes","No"}{@validate:required}'], [null, '{/items}'],
  ]))
  const data = { items: [{ name: 'Missing' }, { name: 'False', enabled: false }, { name: 'Zero', enabled: 0 }] }
  expect(await readWorkbookForm(template, await renderWorkbookForm(template, data))).toMatchObject({ success: false,
    data: { items: [{ name: 'Missing', enabled: null }, { name: 'False', enabled: 'No' }, { name: 'Zero', enabled: 'No' }] },
    issues: [{ code: 'required', address: 'B2' }],
  })
})

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
  expect(() => registerFormatter('testSuffix', () => 'Replaced')).toThrow('already registered')
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

import { expect, it } from 'vitest'
import type { CellValue } from 'exceljs'
import { readFile } from 'node:fs/promises'
import { importWorkbookXlsx, readWorkbookForm, renderWorkbookForm } from '../src/index'
import { importAuthoredWorkbook, openWorkbook, saveWorkbook } from './xlsx'

it('reads a workbook saved by native Excel, including blank formatters and local choices', async () => {
  const template = await importWorkbookXlsx(await readFile(new URL('fixtures/form-values-template.xlsx', import.meta.url)))
  const completed = await readFile(new URL('fixtures/form-values-excel.xlsx', import.meta.url))
  expect(await readWorkbookForm(template, completed)).toEqual({ success: true, data: {
    quantity: 2, price: 10.08, code: '006', numericText: 12.35, date: '2026-10-02T00:00:00.000Z',
    ratio: 0.1235, scientific: 1230000, negative: -1.01, scaled: 1230000, duration: 1.5,
    numberAsText: '123', items: [], selected: 'a',
  } })
})

it.each<[CellValue, string, unknown]>([
  ['12.5', '0.00', 12.5],
  ['12,345', '0.00', 12.35],
  [' 12.5 ', 'General', 12.5],
  ['006', '@', '006'],
  ['006', 'General', '006'],
  ['006', '0', 6],
  [123, '@', '123'],
  [10.075, '#,##0.00', 10.08],
  [-1.005, '0.00;[Red](0.00)', -1.01],
  [12.34567, '0.000', 12.346],
  [12.34567, '0.###', 12.346],
  [12.34567, 'General', 12.34567],
  [12.5, '0', 13],
  [0.123456, '0.00%', 0.1235],
  [1234567, '0.00,,', 1230000],
  [1234567, '0.00E+00', 1230000],
  [0.00001234567, '0.00E+00', 0.0000123],
  [12345.67, '##0.0E+0', 12300],
  [12345.67, '##0.00E+00', 12350],
  [0.01234567, '##0.0E+0', 0.0123],
  [12.34567, '00.0E+0', 12.3],
  [12345.67, '00.0E+0', 12000],
  [999.96, '##0.0E+0', 1000],
  [-12345.67, '##0.0E+0', -12300],
  [{ formula: '5/2', result: 2.5 }, '0', 3],
  [{ formula: '10.075', result: 10.075 }, '0.00', 10.08],
  [{ formula: '"006"', result: '006' }, '@', '006'],
  [{ formula: '1=2', result: false }, 'General', false],
  [{ formula: '0', result: 0 }, 'General', 0],
  [{ formula: '""', result: '' }, 'General', null],
  [12.345, '"Revision 2.0; "0.00" units"', 12.35],
  [12.345, '[>=100]0.0;0.00', 12.35],
  [123.456, '[>=100]0.0;0.00', 123.5],
  [12.34567, '# ?/?', 12.34567],
  [1.5, '[h]:mm', 1.5],
  [false, 'General', false],
  [0, '0.00', 0],
  ['', '0.00', null],
  ['not a number', '0.00', 'not a number'],
  [new Date('2026-10-02T12:34:56Z'), 'dd.mm.yyyy hh:mm:ss', '2026-10-02T12:34:56.000Z'],
])('reads submitted value %j using Excel format %s', async (value, format, expected) => {
  const template = await importAuthoredWorkbook(book => {
    const sheet = book.addWorksheet('Input')
    sheet.getCell('A1').value = '{value}'
    sheet.getCell('A1').numFmt = '@'
  })
  const book = await openWorkbook(await renderWorkbookForm(template, {}))
  const cell = book.worksheets[0].getCell('A1')
  cell.value = value
  cell.numFmt = format
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toEqual({ success: true, data: { value: expected } })
})

it('keeps blank list fields null and uses saved formula labels without losing zero', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').addRows([
    ['{#items}'], ['{.name}', '{.list}{@list:Codes}', '{.choice}{@choice:Options; key=id; label=name; return=key}', '{.plain}'],
    [null, null, null, '{/items}'],
  ]))
  const book = await openWorkbook(await renderWorkbookForm(template, { items: [{ name: 'Blank' }, { name: 'Zero' }] }, {
    dictionaries: { Codes: ['0', '006'], Options: [{ id: 'zero', name: '0' }] },
  }))
  book.worksheets[0].getCell('B3').value = { formula: '0', result: 0 }
  book.worksheets[0].getCell('C3').value = { formula: '0', result: 0 }
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toEqual({ success: true, data: { items: [
    { name: 'Blank', list: null, choice: null, plain: null }, { name: 'Zero', list: '0', choice: 'zero', plain: null },
  ] } })
})

it.each([{ formula: '1+1' }, { formula: '1/0', result: { error: '#DIV/0!' as const } }])('reports a missing or failed formula result without discarding other values (%j)', async value => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').addRow(['{name}', '{amount}']))
  const book = await openWorkbook(await renderWorkbookForm(template, { name: 'Keep' }))
  book.worksheets[0].getCell('B1').value = value
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false,
    data: { name: 'Keep', amount: null }, issues: [{ phase: 'value', code: 'formula', address: 'B1' }],
  })
})

it('validates converted values and retains usable data alongside input errors', async () => {
  const template = await importAuthoredWorkbook(book => {
    const sheet = book.addWorksheet('Input')
    sheet.addRows([['{name}', '{amount}{@validate:required|number}', '{code}{@choice:Options; key=id; label=name; return=key}']])
    sheet.getCell('B1').numFmt = '0.00'
  })
  const book = await openWorkbook(await renderWorkbookForm(template, { name: 'Keep me' }, { dictionaries: { Options: [{ id: 'a', name: 'Allowed' }] } }))
  book.worksheets[0].getCell('C1').value = 'Unknown'
  const invalid = await readWorkbookForm(template, await saveWorkbook(book))
  expect(invalid).toMatchObject({ success: false, data: { name: 'Keep me', amount: null, code: 'Unknown' }, issues: [{ code: 'required' }, { code: 'choice' }] })
  book.worksheets[0].getCell('B1').value = '12.345'
  book.worksheets[0].getCell('C1').value = 'Allowed'
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toEqual({ success: true, data: { name: 'Keep me', amount: 12.35, code: 'a' } })
  expect(await readWorkbookForm(template, Buffer.from('broken file'))).not.toHaveProperty('data')
})

it('decodes numeric-looking list labels before applying numeric input formats', async () => {
  const template = await importAuthoredWorkbook(book => {
    const sheet = book.addWorksheet('Input')
    sheet.addRows([[
      '{year}{@list:Years}', '{selected}{@choice:Options; key=id; label=name; return=key}',
      '{free}{@choice:.options; key=id; label=name; return=key}',
    ]])
    sheet.getCell('B1').numFmt = '0.00'
    sheet.getCell('C1').numFmt = '0.00'
  })
  const issued = await renderWorkbookForm(template, { year: '2026', selected: 'a', free: '10.075' }, {
    dictionaries: { Years: ['2026', '2027'], Options: [{ id: 'a', name: '10.075' }] },
  })
  expect(await readWorkbookForm(template, issued)).toEqual({ success: true, data: { year: '2026', selected: 'a', free: 10.08 } })
})

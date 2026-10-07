import { expect, it } from 'vitest'
import { readWorkbookForm, renderWorkbookForm, renderWorkbookReport, resolveWorkbook } from '../src/index'
import { importAuthoredWorkbook, openWorkbook, saveWorkbook } from './xlsx'
import { data as comparisonData, definition as comparison } from '../examples/comparison/definition'

const template = await importAuthoredWorkbook(book => {
  const sheet = book.addWorksheet('Items')
  sheet.addRows([
    ['Header'], ['{#order.items}'],
    ['{.name}', '{.quantity}', { formula: 'IF(B3="","",B3*2)' }],
    [null, null, '{/order.items}'], ['Footer'],
  ])
  sheet.getRow(3).height = 31
  sheet.getCell('B3').numFmt = '0.00'
})

it.each([{}, { order: undefined }, { order: null }, { order: {} }, { order: { items: undefined } }, { order: { items: null } }, { order: { items: [] } }])('treats absent, undefined and null collections like empty arrays through report and form XLSX %#', async data => {
  const before = structuredClone(data)
  expect(resolveWorkbook(template, data)).toEqual(resolveWorkbook(template, { order: { items: [] } }))
  const report = (await openWorkbook(await renderWorkbookReport(template, data))).worksheets[0]
  expect(report.getCell('A1').value).toBe('Header')
  expect(report.getCell('A2').value).toBe('Footer')

  const bytes = await renderWorkbookForm(template, data)
  expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data: { order: { items: [] } } })
  const book = await openWorkbook(bytes)
  const sheet = book.worksheets[0]
  expect(sheet.getRow(3).hidden).toBeFalsy()
  expect(sheet.getRow(3).height).toBe(31)
  expect(sheet.getCell('B3').numFmt).toBe('0.00')
  expect(sheet.getCell('C3').formula).toBe('IF(B3="","",B3*2)')
  expect(sheet.getCell('A5').value).toBe('Footer')
  sheet.getCell('A3').value = 'Added'
  sheet.getCell('B3').value = 0
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toEqual({ success: true, data: { order: { items: [{ name: 'Added', quantity: 0 }] } } })
  expect(data).toStrictEqual(before)
})

it('handles missing and null nested repeats independently in each record', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Groups').addRows([
    ['{#groups}'], ['{.name}'], ['{#.items}'], ['{.name}'], ['{/.items}'], ['{/groups}'], ['End'],
  ]))
  const groups = [{ name: 'Missing' }, { name: 'Null', items: null }, { name: 'Empty', items: [] }, { name: 'Full', items: [{ name: 'Existing' }] }]
  const data = { groups }
  const before = structuredClone(data)
  const report = (await openWorkbook(await renderWorkbookReport(template, data))).worksheets[0]
  expect([1, 2, 3, 4, 5, 6].map(row => report.getCell(row, 1).value)).toEqual(['Missing', 'Null', 'Empty', 'Full', 'Existing', 'End'])
  expect(await readWorkbookForm(template, await renderWorkbookForm(template, data))).toEqual({ success: true, data: {
    groups: [{ name: 'Missing', items: [] }, { name: 'Null', items: [] }, { name: 'Empty', items: [] }, { name: 'Full', items: [{ name: 'Existing' }] }],
  } })
  expect(data).toEqual(before)
})

it('treats missing column repeats and nested row repeats as empty without changing adjacent bands', async () => {
  const variants = [
    [{ items: comparisonData.items }, { items: comparisonData.items, offers: [] }],
    [{ items: null, offers: null }, { items: [], offers: [] }],
    [{ items: comparisonData.items, offers: [{ name: 'First' }, { name: 'Second', lines: null }] },
      { items: comparisonData.items, offers: [{ name: 'First', lines: [] }, { name: 'Second', lines: [] }] }],
  ]
  for (const [input, explicit] of variants) {
    const before = structuredClone(input)
    expect(resolveWorkbook(comparison, input)).toEqual(resolveWorkbook(comparison, explicit))
    const actual = await openWorkbook(await renderWorkbookReport(comparison, input))
    const expected = await openWorkbook(await renderWorkbookReport(comparison, explicit))
    for (const [index, sheet] of actual.worksheets.entries()) {
      expect(sheet.getSheetValues()).toEqual(expected.worksheets[index].getSheetValues())
      expect(sheet.model.merges).toEqual(expected.worksheets[index].model.merges)
    }
    expect(input).toEqual(before)
  }
})

it.each([{}, 1, '', false, [null], [1]])('still rejects malformed collections through every public execution entry point %#', async items => {
  for (const run of [resolveWorkbook, renderWorkbookReport, renderWorkbookForm]) {
    await expect(Promise.resolve().then<unknown>(() => run(template, { order: { items } }))).rejects.toMatchObject({ issues: [{
      code: Array.isArray(items) ? 'invalid-item' : 'invalid-collection',
      path: '$data.order.items' + (Array.isArray(items) ? '[0]' : ''),
    }] })
  }
})

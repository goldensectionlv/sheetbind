import { expect, it, vi } from 'vitest'
import { readWorkbookForm, renderWorkbookForm, renderWorkbookReport, resolveWorkbook } from '../src/index'
import { assertJson } from '../src/core/json'
import { importAuthoredWorkbook, openWorkbook, saveWorkbook } from './xlsx'

it('accepts undefined properties at every depth without changing the caller data', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').addRows([
    ['{?supplier.name}', '{?supplier.inn}', '{?supplier.kpp}', '{?optional.detail}'],
    ['{#items}'], ['{.name}', '{?.note}', '{.zero}', '{.enabled}'], [null, null, null, '{/items}'],
  ]))
  const data = {
    supplier: { name: undefined, inn: '006', kpp: undefined }, optional: undefined,
    unused: { nested: [{ value: undefined, other: { missing: undefined } }] },
    items: [{ name: 'Item', note: undefined, zero: 0, enabled: false }],
  }
  const before = structuredClone(data)
  expect(resolveWorkbook(template, data)).toEqual(resolveWorkbook(template, {
    supplier: { inn: '006' }, items: [{ name: 'Item', zero: 0, enabled: false }],
  }))
  const report = (await openWorkbook(await renderWorkbookReport(template, data))).worksheets[0]
  expect(['A1', 'B1', 'C1', 'D1'].map(address => report.getCell(address).value)).toEqual([null, '006', null, null])
  expect(['A2', 'B2', 'C2', 'D2'].map(address => report.getCell(address).value)).toEqual(['Item', null, 0, false])
  expect(await readWorkbookForm(template, await renderWorkbookForm(template, data))).toEqual({ success: true, data: {
    supplier: { name: null, inn: '006', kpp: null }, optional: { detail: null },
    items: [{ name: 'Item', note: null, zero: 0, enabled: false }],
  } })
  expect(data).toStrictEqual(before)
})

it.each([undefined, { inn: undefined, extra: undefined }])('issues ordinary form fields when their container or value is undefined %#', async supplier => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = '{supplier.inn}{@validate:required}')
  const book = await openWorkbook(await renderWorkbookForm(template, { supplier }))
  expect(book.worksheets[0].getCell('A1').value).toBeNull()
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [{ rule: 'required', path: '$data.supplier.inn' }] })
  book.worksheets[0].getCell('A1').value = '006'
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toEqual({ success: true, data: { supplier: { inn: '006' } } })
})

it.each([{ order: undefined }, { order: { items: undefined } }])('treats undefined repeat properties like missing collections %#', async data => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').addRows([
    ['{#order.items}'], ['{.name}'], ['{/order.items}'], ['End'],
  ]))
  const before = structuredClone(data)
  expect((await openWorkbook(await renderWorkbookReport(template, data))).worksheets[0].getCell('A1').value).toBe('End')
  expect(await readWorkbookForm(template, await renderWorkbookForm(template, data))).toEqual({ success: true, data: { order: { items: [] } } })
  expect(data).toStrictEqual(before)
})

it('still diagnoses undefined required report bindings as missing fields', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = '{supplier.inn}')
  await expect(renderWorkbookReport(template, { supplier: { inn: undefined } })).rejects.toMatchObject({ issues: [{ code: 'missing-source', path: '$data.supplier.inn' }] })
})

it('omits undefined dictionary properties from saved named, root and local choice sources', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').addRows([
    ['{named}{@choice:Options; key=id; label=label}'],
    ['{root}{@choice:$root.catalog; key=id; label=label}'],
    ['{#items}'], ['{.local}{@choice:.options; key=id; label=label}'], ['{/items}'],
  ]))
  const object = { id: '006', label: 'Option', optional: undefined, nested: { missing: undefined, zero: 0, enabled: false, blank: null } }
  const normalized = { id: '006', label: 'Option', nested: { zero: 0, enabled: false, blank: null } }
  const data = { named: object, root: object, catalog: [object], items: [{ local: object, options: [object] }] }
  const dictionaries = { Options: [object] }
  const before = structuredClone({ data, dictionaries })
  await renderWorkbookReport(template, data, { dictionaries })
  const issued = await renderWorkbookForm(template, data, { dictionaries })
  const book = await openWorkbook(issued)
  const range = book.definedNames.getRanges('_sb_object_sources').ranges[0]
  const [start, end] = range.split('!')[1].replaceAll('$', '').split(':')
  const sheet = book.worksheets.find(sheet => sheet.state === 'veryHidden')!
  const first = sheet.getCell(start)
  const last = sheet.getCell(end)
  let text = ''
  for (let row = Number(first.row) + 1; row <= Number(last.row); row++) {
    text += sheet.getCell(row, Number(first.col)).value
  }
  const stored = JSON.parse(text)
  expect(stored.dictionaries.Options).toStrictEqual([normalized])
  expect(stored.context.catalog).toStrictEqual([normalized])
  expect(Object.values(stored.local)).toEqual([expect.objectContaining({ values: [[normalized]] })])
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toStrictEqual({ success: true, data: {
    named: normalized, root: normalized, items: [{ local: normalized }],
  } })
  expect({ data, dictionaries }).toStrictEqual(before)
})

it('keeps persisted JSON strict and rejects accessors without executing them', async () => {
  expect(() => assertJson({ optional: undefined })).toThrow(SyntaxError)
  const getter = vi.fn(() => 'Option')
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = '{value}{@choice:Options; key=id; label=label}')
  const item = { id: '006', label: 'Option', optional: undefined }
  Object.defineProperty(item, 'label', { get: getter, enumerable: true })
  await expect(renderWorkbookForm(template, {}, { dictionaries: { Options: [item] } })).rejects.toThrow(SyntaxError)
  expect(getter).not.toHaveBeenCalled()
})

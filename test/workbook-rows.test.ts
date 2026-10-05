import { expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import { readWorkbookForm, renderWorkbookForm } from '../src/xlsx/workbook-form'
import { importWorkbookXlsx } from '../src/xlsx/workbook-template'
import { FORM_MARKER_PREFIX } from '../src/xlsx/workbook-form-markers'
import { exampleFile, importAuthoredWorkbook, openWorkbook, saveWorkbook } from './xlsx'
import { data, dictionaries } from '../examples/records/definition'

async function templateBytes() {
  const book = new ExcelJS.Workbook()
  book.addWorksheet('Lines').addRows([
    ['{contact}{@validate:required}'],
    ['{#items}'],
    ['{.name}{@validate:required}', '{.quantity}{@validate:number|min:0}', '{.approved}{@validate:boolean}'],
    [null, null, '{/items}'], ['End'],
  ])
  return saveWorkbook(book)
}

it('preserves Unicode field names through import, issuance and reading', async () => {
  const template = await importAuthoredWorkbook(book => {
    book.addWorksheet('Fields').addRows([
      ['{document.сode}'],
      ['{#éléments}'],
      ['{.名称}', '{.café}'],
      [null, '{/éléments}'],
    ])
  })
  const input = { document: { сode: '001' }, éléments: [{ 名称: 'Paper', café: '002' }] }
  expect(await readWorkbookForm(template, await renderWorkbookForm(template, input))).toEqual({ success: true, data: input })
})

it('issues twenty explicitly empty records with styles, formulas and lists, then reads filled lines', async () => {
  const template = await importAuthoredWorkbook(book => {
    const sheet = book.addWorksheet('Lines')
    sheet.addRows([
      ['{#items}'],
      ['{.name}{@validate:required}', '{.quantity}{@validate:number|min:0}', '{.status}{@list:Statuses}', { formula: 'IF(B2="","",B2*2)' }],
      [null, null, null, '{/items}'], ['End'],
    ])
    sheet.getRow(2).height = 28
    sheet.getCell('B2').numFmt = '0.000'
    sheet.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFCC' } }
  })
  const dictionaries = { Statuses: ['Open', 'Closed'] }
  const input = { items: Array.from({ length: 20 }, () => ({})) }
  const bytes = await renderWorkbookForm(template, input, { dictionaries })
  const book = await openWorkbook(bytes)
  const sheet = book.worksheets[0]
  const start = Number(find(sheet, FORM_MARKER_PREFIX + '["repeat",1]').row) + 1
  expect(Number(find(sheet, 'End').row) - start).toBe(21)
  for (let index = 0; index < 20; index++) {
    const row = sheet.getRow(start + index)
    expect(row.height).toBe(28)
    expect(row.getCell(1).value).toBeNull()
    expect(row.getCell(1).fill).toMatchObject({ fgColor: { argb: 'FFFFFFCC' } })
    expect(row.getCell(2).numFmt).toBe('0.000')
    expect(row.getCell(3).dataValidation).toMatchObject({ type: 'list' })
    expect(row.getCell(4).formula).toBe(`IF(B${row.number}="","",B${row.number}*2)`)
  }
  expect(await readWorkbookForm(template, bytes, { dictionaries })).toEqual({ success: true, data: { items: [] } })
  for (const index of [0, 9, 19]) {
    const row = sheet.getRow(start + index)
    row.getCell(1).value = `Line ${index + 1}`
    row.getCell(2).value = index
    row.getCell(3).value = 'Open'
  }
  expect(await readWorkbookForm(template, await saveWorkbook(book), { dictionaries })).toEqual({ success: true, data: {
    items: [0, 9, 19].map(index => ({ name: `Line ${index + 1}`, quantity: index, status: 'Open' })),
  } })
  expect(input.items).toEqual(Array.from({ length: 20 }, () => ({})))
})

function find(sheet: ExcelJS.Worksheet, value: unknown): ExcelJS.Cell {
  let found: ExcelJS.Cell | undefined
  sheet.eachRow(row => row.eachCell(cell => {
    if (cell.value === value && (!cell.isMerged || cell.master === cell)) {
      found = cell
    }
  }))
  if (!found) {
    throw new Error(`Missing ${value}`)
  }
  return found
}

it.each([0, 1, 30])('reads %s lines with a freshly imported template and no issuance data', async count => {
  const source = await templateBytes()
  const input = { contact: 'Jordan', items: Array.from({ length: count }, (_item, index) => ({ name: `Line ${index}`, quantity: index, approved: false })) }
  const snapshot = structuredClone(input)
  const rendered = await renderWorkbookForm(await importWorkbookXlsx(source), input)
  const result = await readWorkbookForm(await importWorkbookXlsx(source), rendered)
  expect(result).toEqual({ success: true, data: input })
  expect(input).toEqual(snapshot)
  if (!count) {
    const book = await openWorkbook(rendered)
    const sheet = book.worksheets[0]
    const start = Number(find(sheet, FORM_MARKER_PREFIX + '["repeat",1]').row) + 1
    sheet.getCell(start, 1).value = 'First line'
    sheet.getCell(start, 2).value = 0
    sheet.getCell(start, 3).value = false
    expect(await readWorkbookForm(await importWorkbookXlsx(source), await saveWorkbook(book))).toEqual({ success: true, data: {
      contact: input.contact, items: [{ name: 'First line', quantity: 0, approved: false }],
    } })
  }
})

it('reads inserted, removed, reordered and cleared rows without record identities', async () => {
  const source = await templateBytes()
  const input = { contact: 'Jordan', items: Array.from({ length: 30 }, (_item, index) => ({ name: `Line ${index}`, quantity: index, approved: false })) }
  const book = await openWorkbook(await renderWorkbookForm(await importWorkbookXlsx(source), input))
  const sheet = book.worksheets[0]
  const start = Number(find(sheet, 'Line 0').row)
  sheet.spliceRows(start + 1, 1)
  sheet.spliceRows(start, 0, ['Added', 0, false], [])
  const first = sheet.getRow(start + 2).values
  sheet.getRow(start + 2).values = sheet.getRow(start + 3).values
  sheet.getRow(start + 3).values = first
  expect(await readWorkbookForm(await importWorkbookXlsx(source), await saveWorkbook(book))).toEqual({ success: true, data: {
    contact: input.contact, items: [{ name: 'Added', quantity: 0, approved: false }, input.items[2], input.items[0], ...input.items.slice(3)],
  } })
  sheet.spliceRows(start, 31)
  expect(await readWorkbookForm(await importWorkbookXlsx(source), await saveWorkbook(book))).toEqual({ success: true, data: { contact: input.contact, items: [] } })
})

it('validates partially filled rows at their resulting paths after skipping empty lines', async () => {
  const template = await importWorkbookXlsx(await templateBytes())
  const book = await openWorkbook(await renderWorkbookForm(template, { contact: 'Jordan', items: [{}, { quantity: -1 }] }))
  const address = find(book.worksheets[0], -1).address
  const result = await readWorkbookForm(template, await saveWorkbook(book), { context: { contact: 'Other', items: [{ name: 'Not submitted' }] } })
  expect(result).toMatchObject({ success: false, issues: [
    { code: 'required', path: '$data.items[0].name' }, { code: 'min', path: '$data.items[0].quantity', address },
  ] })
})

it.each(['choice', 'choice-source'])('keeps a nonempty %s failure and reports its compacted nested path', async code => {
  const template = await importAuthoredWorkbook(book => {
    book.addWorksheet('Nested').addRows([
      ['{#departments}'], ['{.name}'], ['{#.groups}'], ['{.name}'], ['{#.items}'],
      ['{.product}{@validate:required|string}{@choice:$root.products; key=id; label=name; return=key}'],
      ['{/.items}'], ['{/.groups}'], ['{/departments}'],
    ])
  })
  const products = [{ id: '001', name: 'Paper' }]
  const input = { products, departments: [{ name: 'Office', groups: [{ name: 'Supplies', items: [{ product: '001' }, {}, {}] }] }] }
  const book = await openWorkbook(await renderWorkbookForm(template, input))
  const sheet = book.worksheets[0]
  const first = find(sheet, 'Paper')
  const edited = sheet.getCell(Number(first.row) + 2, 1)
  edited.value = 'Unknown'
  const result = await readWorkbookForm(template, await saveWorkbook(book), { context: code === 'choice' ? { products } : {} })
  expect(result.success).toBe(false)
  if (!result.success) {
    expect(result.issues.filter(issue => issue.address === edited.address)).toEqual([
      expect.objectContaining({ code, sheetName: 'Nested', address: edited.address, path: '$data.departments[0].groups[0].items[1].product' }),
    ])
    expect(result.issues.every(issue => issue.code === code)).toBe(true)
  }
})

it('retains a conflicting nonempty occurrence when the first occurrence is blank', async () => {
  const template = await importAuthoredWorkbook(book => {
    book.addWorksheet('Lines').addRows([['{#items}'], ['{.name}', '{.name}'], [null, '{/items}']])
  })
  const book = await openWorkbook(await renderWorkbookForm(template, { items: [{}, {}, {}] }))
  const sheet = book.worksheets[0]
  const start = Number(find(sheet, FORM_MARKER_PREFIX + '["repeat",1]').row) + 1
  const edited = sheet.getCell(start + 2, 2)
  edited.value = 'Entered'
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [
    { code: 'conflicting-field', address: edited.address, path: '$data.items[0].name' },
  ] })
})

it.each([{ formula: '1+1' }, new Date('2026-01-01T00:00:00Z')])('keeps an invalid native input and locates it after blank rows', async value => {
  const template = await importWorkbookXlsx(await templateBytes())
  const book = await openWorkbook(await renderWorkbookForm(template, { contact: 'Jordan', items: [{}, {}, {}] }))
  const sheet = book.worksheets[0]
  const start = Number(find(sheet, FORM_MARKER_PREFIX + '["repeat",1]').row) + 1
  const cell = sheet.getCell(start + 2, 1)
  cell.value = value
  if (value instanceof Date) {
    cell.numFmt = 'yyyy-mm-dd'
  }
  const result = await readWorkbookForm(template, await saveWorkbook(book))
  expect(result).toMatchObject({ success: false, issues: [
    { code: value instanceof Date ? 'non-scalar' : 'formula', address: cell.address, path: '$data.items[0].name' },
  ] })
})

it('reads child insertion under an unkeyed parent and treats its identifier as an ordinary field', async () => {
  const source = await exampleFile('records/template.xlsx')
  const template = await importWorkbookXlsx(source)
  const input = { ...data, sites: data.sites.map(site => ({ ...site, note: site.note?.replace('_x000a_', 'reviewed') ?? null })) }
  const book = await openWorkbook(await renderWorkbookForm(template, input, { dictionaries }))
  const sheet = book.worksheets[0]
  find(sheet, 'South site').value = 'Renamed site'
  const start = Number(find(sheet, '0007').row)
  sheet.spliceRows(start, 0, ['0007', 'Extra', 3, false])
  const merges = sheet.model.merges
  sheet.unMergeCells('A1:IW100')
  for (const ref of merges) {
    sheet.mergeCells(ref.replace(/(\d+)/g, value => String(Number(value) >= start ? Number(value) + 1 : Number(value))))
  }
  const expected = { ...input, sites: input.sites.map((site, index) => index === 2
    ? { ...site, name: 'Renamed site', work: [{ code: '0007', description: 'Extra', hours: 3, approved: false }, ...site.work] }
    : site) }
  expect(await readWorkbookForm(await importWorkbookXlsx(source), await saveWorkbook(book), { dictionaries })).toEqual({ success: true, data: expected })
})

it('reads copied and deleted multirow blocks by their current boundaries', async () => {
  const template = await importAuthoredWorkbook(book => {
    book.addWorksheet('Blocks').addRows([['{#items}'], ['Name', '{.name}'], ['Quantity', '{.quantity}'], [null, '{/items}']])
  })
  const book = await openWorkbook(await renderWorkbookForm(template, { items: [{ name: 'First', quantity: 1 }] }))
  const sheet = book.worksheets[0]
  const start = Number(find(sheet, FORM_MARKER_PREFIX + '["item",1]').row)
  const rows = Array.from({ length: 4 }, (_item, index) => sheet.getRow(start + index).values as ExcelJS.CellValue[])
  sheet.spliceRows(start + 4, 0, ...rows)
  sheet.getCell(start + 5, 2).value = 'Copied'
  sheet.spliceRows(start, 4)
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toEqual({ success: true, data: { items: [{ name: 'Copied', quantity: 1 }] } })
})

it('locates the first displaced boundary of an incomplete multirow block', async () => {
  const template = await importAuthoredWorkbook(book => {
    book.addWorksheet('Blocks').addRows([['{#items}'], ['Name', '{.name}'], ['Quantity', '{.quantity}'], [null, '{/items}']])
  })
  const book = await openWorkbook(await renderWorkbookForm(template, { items: [{ name: 'First', quantity: 1 }] }))
  book.worksheets[0].spliceRows(5, 0, ['An extra line'])
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [
    { code: 'form-layout', sheetName: 'Blocks', address: 'IW7', message: expect.stringContaining('expected at IW6, found IW7') },
  ] })
})

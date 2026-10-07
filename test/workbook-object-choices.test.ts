import { expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import { objectDefinition, objectData, objectResult, dictionaries } from '../examples/choices/definition'
import { resolveWorkbook, renderWorkbookReport } from '../src/xlsx/workbook-template'
import { readWorkbookForm, renderWorkbookForm } from '../src/xlsx/workbook-form'
import { workbookChoiceRange } from '../src/xlsx/workbook-choice-fields'
import { createChoiceResolver } from '../src/core/choices'
import { saveWorkbook, editExample, importAuthoredWorkbook } from './xlsx'
import { parseRange } from '../src/xlsx/addresses'

async function load(bytes: Buffer) {
  const book = new ExcelJS.Workbook()
  await book.xlsx.load(Uint8Array.from(bytes).buffer)
  return book
}
const bytes = saveWorkbook
function find(sheet: ExcelJS.Worksheet, value: unknown) {
  let found: ExcelJS.Cell | undefined
  sheet.eachRow(row => row.eachCell(cell => {
    if (cell.value === value) {
      found = cell
    }
  }))
  if (!found) {
    throw new Error(`Missing ${value}`)
  }
  return found
}
function rangeValues(book: ExcelJS.Workbook, name: string) {
  const range = book.definedNames.getRanges(name).ranges[0]
  const divider = range.lastIndexOf('!')
  const sheet = book.getWorksheet(range.slice(0, divider).replace(/^'|'$/g, ''))!
  const { start, end } = parseRange(range.slice(divider + 1).replace(/\$/g, ''))
  return Array.from({ length: end.row - start.row + 1 }, (_item, index) => sheet.getCell(start.row + index, start.column).value)
}
const codes = (result: Awaited<ReturnType<typeof readWorkbookForm>>) => {
  expect(result.success).toBe(false)
  return result.success ? [] : result.issues.map(issue => issue.code)
}

it('round-trips selected objects through tagged XLSX without application lookup data', async () => {

  const form = await renderWorkbookForm(objectDefinition, objectData, { dictionaries })
  expect(await readWorkbookForm(objectDefinition, form)).toEqual({ success: true, data: objectResult })
  const book = await load(form)
  find(book.worksheets[0], 'Service [0007]').value = 'Service [0008]'
  const result = await readWorkbookForm(objectDefinition, await bytes(book))
  expect(result).toEqual({ success: true, data: { ...objectResult, items: [{ ...objectResult.items[0], product: dictionaries.products[1] }, objectResult.items[1]] } })
  find(book.worksheets[0], 'Service [0008]').value = 'Service'
  expect(codes(await readWorkbookForm(objectDefinition, await bytes(book)))).toContain('choice')
})

it('reads objects from shared sources after sorting', async () => {
  const book = await load(await renderWorkbookForm(objectDefinition, objectData, { dictionaries }))
  const sheet = book.worksheets[0]
  const row = Number(find(sheet, 'line-a').row)
  const first = sheet.getRow(row).values
  sheet.getRow(row).values = sheet.getRow(row + 1).values
  sheet.getRow(row + 1).values = first
  const result = await readWorkbookForm(objectDefinition, await bytes(book))
  expect(result).toEqual({ success: true, data: { category: objectResult.category, items: [objectResult.items[1], objectResult.items[0]] } })
})

it('names Excel fields independently of key order and includes keys absent from the first option', async () => {
  const rule = { source: { dictionary: 'products' }, key: 'id', label: 'name', return: 'object' as const }
  const config = await editExample('choices/object-template.xlsx', book => {
    book.worksheets[0].getCell('B4').value = '{.product}{@validate:required|object}{@choice:products; key=id; label=name; return=object}'
    for (const [index, field] of ['rate', 'code', 'active', 'details'].entries()) {
      book.worksheets[0].getCell(1, index + 6).value = { formula: `INDEX(${workbookChoiceRange(rule, field)},MATCH(B4,${workbookChoiceRange(rule)},0))` }
    }
  })
  const products = dictionaries.products
  const sources = { ...dictionaries, products }
  const base = await load(await renderWorkbookReport(config, objectData, { dictionaries: sources }))
  const reordered = products.map(item => Object.fromEntries(Object.entries(item).reverse())).reverse()
  const other = await load(await renderWorkbookReport(config, objectData, { dictionaries: { ...sources, products: reordered } }))
  expect(rangeValues(base, workbookChoiceRange(rule))).toEqual(['Service [0007]', 'Service [0008]', 'Inspection'])
  expect(rangeValues(base, workbookChoiceRange(rule, 'rate'))).toEqual([null, 22, null])
  expect(rangeValues(base, workbookChoiceRange(rule, 'code'))).toEqual([null, '006', null])
  expect(rangeValues(base, workbookChoiceRange(rule, 'active'))).toEqual([null, false, null])
  expect(rangeValues(base, workbookChoiceRange(rule, 'details'))).toEqual([null, '{"unit":"hour"}', null])
  const names = (book: ExcelJS.Workbook) => book.definedNames.model.map(entry => entry.name).filter(name => name.startsWith('_sb_ref_')).sort()
  expect(names(other)).toEqual(names(base))
  expect(rangeValues(other, workbookChoiceRange(rule, 'rate'))).toEqual([null, 22, null])
})

it('reads object choices in new unkeyed rows from a shared source', async () => {
  const config = await importAuthoredWorkbook(book => {
    book.addWorksheet('Order').addRows([
      ['{#items}'], ['{.product}{@validate:required|object}{@choice:products; key=id; label=name}', '{.quantity}{@validate:number}'], [null, '{/items}'],
    ])
  })
  const products = dictionaries.products.slice(0, 2)
  const book = await load(await renderWorkbookForm(config, { items: [{ product: products[0], quantity: 1 }] }, { dictionaries: { products } }))
  const sheet = book.worksheets[0]
  const row = Number(find(sheet, 'Service [0007]').row)
  sheet.spliceRows(row + 1, 0, ['Service [0008]', 0])
  expect(await readWorkbookForm(config, await bytes(book))).toEqual({ success: true, data: { items: [
    { product: products[0], quantity: 1 }, { product: products[1], quantity: 0 },
  ] } })
})

it.each(['products', '$root.products'])('reads sorted, inserted and deleted records without identity fields using %s', async source => {
  const config = await importAuthoredWorkbook(book => {
    book.addWorksheet('Order').addRows([
      ['{#items}'], ['{.product}{@choice:' + source + '; key=id; label=name}', '{.quantity}{@validate:number}'], [null, '{/items}'],
    ])
  })
  const products = dictionaries.products
  const input = { products, items: [{ product: products[0], quantity: 2 }, { product: products[2], quantity: 3 }] }
  const book = await load(await renderWorkbookForm(config, input, { dictionaries: { products } }))
  const sheet = book.worksheets[0]
  const first = Number(find(sheet, 'Service [0007]').row)
  const values = sheet.getRow(first).values
  sheet.getRow(first).values = sheet.getRow(first + 1).values
  sheet.getRow(first + 1).values = values
  sheet.spliceRows(first + 1, 0, ['Service [0008]', 0])
  sheet.spliceRows(first + 2, 1)
  expect(await readWorkbookForm(config, await bytes(book))).toEqual({ success: true, data: { items: [
    { product: products[2], quantity: 3 }, { product: products[1], quantity: 0 },
  ] } })
})

it.each(['object', 'key'])('keeps per-record choice sources in reports and self-contained forms (%s)', async result => {
  const config = await importAuthoredWorkbook(book => {
    book.addWorksheet('Report').addRows([
      ['{#items}'], ['{.product}{@choice:.available; key=id; label=name; return=' + result + '}'], ['{/items}'],
    ])
  })
  const first = { id: 'same', name: 'First', rate: 10 }
  const second = { id: 'same', name: 'Second', rate: 20 }
  const input = { items: [first, second].map(product => ({ product: result === 'key' ? product.id : product, available: [product] })) }
  const report = await load(await renderWorkbookReport(config, input))
  expect(report.worksheets[0].getCell('A1').value).toBe('First')
  expect(report.worksheets[0].getCell('A2').value).toBe('Second')
  for (const data of [input, { items: [] }]) {
    expect(await readWorkbookForm(config, await renderWorkbookForm(config, data))).toEqual({ success: true,
      data: { items: data.items.map(item => ({ product: item.product })) } })
  }
})

it('requires valid embedded source data and does not fall back to guessing by label', async () => {
  const book = await load(await renderWorkbookForm(objectDefinition, objectData, { dictionaries }))
  const range = book.definedNames.getRanges('_sb_object_sources').ranges[0]
  const address = range.slice(range.lastIndexOf('!') + 1).split(':')[0].replace(/\$/g, '')
  book.worksheets[1].getCell(address).value = 'unsupported'
  expect(codes(await readWorkbookForm(objectDefinition, await bytes(book)))).toContain('choice-source')
})

it('keeps numeric transport precision out of core and rejects ambiguous or malformed object rules', async () => {
  const rule = { source: { dictionary: 'options' }, key: 'id', label: 'name' }
  const value = { id: 1234567890123456, name: 'Long key' }
  const config = await importAuthoredWorkbook(book => {
    book.addWorksheet('Choice').getCell('A1').value = '{category}{@validate:object}{@choice:options; key=id; label=name}'
  })
  const issued = await renderWorkbookForm(config, { category: value }, { dictionaries: { options: [value] } })
  expect(await readWorkbookForm(config, issued)).toEqual({ success: true, data: { category: value } })
  const noChoice = await importAuthoredWorkbook(book => {
    book.addWorksheet('Choice').getCell('A1').value = '{category}{@validate:object}'
  })
  expect(() => resolveWorkbook(noChoice, { category: value })).toThrow('declared object choice')
  expect(() => createChoiceResolver()(rule, {}, {}, { options: [{ id: '1', name: 'Same' }, { id: '1', name: 'Other' }] })).toThrow('unique')
})

it.each(['key', 'object'] as const)('keeps unused JSON properties out of formula ranges (%s)', async mode => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = `{selected}{@choice:Options; key=id; label=name; return=${mode}}`)
  const option = { id: 1234567890123456, name: 'Allowed', rate: 1 / 3, sum: 0.1 + 0.2, notes: 'x'.repeat(33000), ['x'.repeat(300)]: { nested: true } }
  const selected = mode === 'key' ? option.id : option
  const options = { dictionaries: { Options: [option] } }
  const issued = await renderWorkbookForm(template, { selected }, options)
  expect(await readWorkbookForm(template, issued)).toEqual({ success: true, data: { selected } })
  for (const content of [issued, await renderWorkbookReport(template, { selected }, options)]) {
    expect((await load(content)).definedNames.model.some(entry => entry.name.startsWith('_sb_ref_') || entry.name.startsWith('_sb_context_'))).toBe(false)
  }
})

it.each(['key', 'object'] as const)('creates only formula-referenced properties independently of return mode (%s)', async mode => {
  const rule = { source: { dictionary: 'Options' }, key: 'id', label: 'name', return: mode }
  const template = await importAuthoredWorkbook(book => {
    const sheet = book.addWorksheet('Input')
    sheet.getCell('A1').value = `{selected}{@choice:Options; key=id; label=name; return=${mode}}`
    sheet.getCell('B1').value = { formula: `INDEX(${workbookChoiceRange(rule, 'rate')},MATCH(A1,${workbookChoiceRange(rule)},0))` }
  })
  const option = { id: 'a', name: 'Allowed', rate: 1 / 3, notes: 'x'.repeat(33000) }
  const selected = mode === 'key' ? option.id : option
  const issued = await renderWorkbookForm(template, { selected }, { dictionaries: { Options: [option] } })
  const book = await load(issued)
  expect(rangeValues(book, workbookChoiceRange(rule, 'rate'))).toEqual([option.rate])
  expect(book.definedNames.getRanges(workbookChoiceRange(rule, 'notes')).ranges).toEqual([])
  expect(await readWorkbookForm(template, issued)).toEqual({ success: true, data: { selected } })
})

it('escapes field names without collisions in Excel case-insensitive identifiers', () => {
  const rule = { source: { dictionary: 'options' }, key: 'id', label: 'name' }
  const names = ['rate', 'Rate', 'a-b', 'a_b', 'a.b', 'a_u002d_b'].map(field => workbookChoiceRange(rule, field).toLowerCase())
  expect(new Set(names).size).toBe(names.length)
  const left = { source: { dictionary: 'a' }, key: 'u002d', label: 'b_c' }
  const right = { source: { dictionary: 'a-b' }, key: 'u005f', label: 'c' }
  expect(workbookChoiceRange(left, 'rate')).not.toBe(workbookChoiceRange(right, 'rate'))
})

it('reads shared choices in nested records after whole-row insertion', async () => {
  const config = await importAuthoredWorkbook(book => {
    book.addWorksheet('Order').addRows([
      ['{category}{@choice:$root.categories; key=id; label=name}'],
      ['{#sites}'], ['{.id}'],
      ['{#.items}'],
      ['{.id}', '{.product}{@choice:products; key=id; label=name}', '{.supplier}{@choice:suppliers; key=id; label=name}', '{.hours}{@validate:number}'],
      [null, null, null, '{/.items}'], [null, null, null, '{/sites}'],
    ])
  })
  const southOption = { id: '0010', name: 'Site service', rate: 10 }
  const input = { ...objectData, sites: [{ id: 'north', items: objectData.items }, { id: 'south', items: [{ ...objectData.items[0], product: southOption }] }] }
  const book = await load(await renderWorkbookForm(config, input, { dictionaries: { ...dictionaries, products: [...dictionaries.products, southOption] } }))
  const sheet = book.worksheets[0]
  const start = Number(find(sheet, 'Service [0007]').row)
  sheet.spliceRows(start, 0, ['line-a', 'Service [0008]', 'Remote supplier', 4])
  expect(await readWorkbookForm(config, await bytes(book))).toMatchObject({ success: true, data: { sites: [
    { id: 'north', items: [{ id: 'line-a', product: dictionaries.products[1], supplier: dictionaries.suppliers[1], hours: 4 }, ...objectResult.items] },
    { id: 'south', items: [{ ...objectResult.items[0], product: southOption }] },
  ] } })
})

it('reads a form whose only object choices are inside an empty repeat', async () => {
  const config = await editExample('choices/object-template.xlsx', book => {
    book.worksheets[0].getCell('B1').value = null
    book.worksheets[0].getCell('A3').value = '{#items}'
  })
  const form = await renderWorkbookForm(config, { items: [] }, { dictionaries })
  expect(await readWorkbookForm(config, form)).toEqual({ success: true, data: { items: [] } })
})

it('leaves repeated labels unchanged in core without imposing workbook display constraints', () => {
  const choice = { source: { dictionary: 'options' }, key: 'id', label: 'name' }
  const options = [{ id: 'a', name: 'One' }, { id: 'b', name: 'One' }, { id: 'c', name: 'One [a]' }]
  const result = createChoiceResolver()(choice, {}, {}, { options })
  expect(result).toMatchObject([{ label: 'One' }, { label: 'One' }, { label: 'One [a]' }])
})

it('shares sources when two sheets show different fields of the same records', async () => {
  const config = await importAuthoredWorkbook(book => {
    book.addWorksheet('First').addRows([
      ['{#items}'], ['{.id}', '{.product}{@choice:products; key=id; label=name}'], [null, '{/items}'],
    ])
    book.addWorksheet('Second').addRows([
      ['{#items}'], ['{.other}{@choice:$root.otherOptions; key=id; label=name}'], ['{/items}'],
    ])
  })
  const other = { id: 'extra', name: 'Extra', count: 2 }
  const input = { ...objectData, otherOptions: [other], items: objectData.items.map(item => ({ ...item, other })) }
  const form = await renderWorkbookForm(config, input, { dictionaries })
  expect(await readWorkbookForm(config, form)).toEqual({ success: true, data: { items: objectData.items.map(item => ({ id: item.id, product: item.product, other })) } })
})

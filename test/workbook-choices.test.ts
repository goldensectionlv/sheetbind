import { expect, it, vi } from 'vitest'
import ExcelJS from 'exceljs'
import { definition, data, declaredData, dictionaries } from '../examples/choices/definition'
import { workbookDictionarySources, resolveWorkbook, importWorkbookXlsx, renderWorkbookReport, renderWorkbookForm, readWorkbookForm } from '../src/index'
import { openWorkbook as load, editExample, importAuthoredWorkbook, saveWorkbook } from './xlsx'

const options = { dictionaries }
function find(sheet: ExcelJS.Worksheet, value: unknown) {
  let result: ExcelJS.Cell | undefined
  sheet.eachRow(row => row.eachCell(cell => {
    if (cell.value === value) {
      result = cell
    }
  }))
  if (!result) {
    throw new Error(`Missing ${value}`)
  }
  return result
}
const codes = (result: Awaited<ReturnType<typeof readWorkbookForm>>) => {
  expect(result.success).toBe(false)
  return result.success ? [] : result.issues.map(issue => issue.code)
}

it('imports readable choice tags and keeps source scope and value types through saved XLSX', async () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Choices')
  const tags = [
    '{?product}{@validate:required|object}{@choice:Products; key=identity.id; label=display.name}',
    '{productCode}{@validate:string}{@choice:.products; key=code; label=name; return=key}',
    '{supplierId}{@validate:number}{@choice:$root.suppliers; key=id; label=name; return=key}',
  ]
  tags.forEach((tag, index) => {
    sheet.getCell(1, index + 1).value = tag
  })
  sheet.getCell('B1').value = '{productCode}{@choice:.products; label=name; return=key; key=code}{@validate:string}'
  const config = await importWorkbookXlsx(await saveWorkbook(book))
  const product = { identity: { id: '006' }, display: { name: 'Paper' } }
  const input = { product, productCode: '007', supplierId: 0, products: [{ code: '007', name: 'Local' }], suppliers: [{ id: 0, name: 'Root' }] }
  const options = { dictionaries: { Products: [product] } }
  const rules = resolveWorkbook(config, input, options).sheets[0].cells.map(cell => cell.rules)
  expect(rules).toEqual([
    { validation: [{ rule: 'required' }, { rule: 'object' }], choice: { source: { dictionary: 'Products' }, key: 'identity.id', label: 'display.name' } },
    { validation: [{ rule: 'string' }], choice: { ...{ source: { path: 'products' }, key: 'code', label: 'name' }, return: 'key' as const } },
    { validation: [{ rule: 'number' }], choice: { ...{ source: { path: 'suppliers', from: 'root' }, key: 'id', label: 'name' }, return: 'key' as const } },
  ])
  expect(await readWorkbookForm(config, await renderWorkbookForm(config, input, options))).toEqual({ success: true, data: { product, productCode: '007', supplierId: 0 } })
  for (const invalid of ['key=id', 'label=name', 'key=__proto__.id; label=name', 'key=id; label=name; return=boolean']) {
    await expect(importAuthoredWorkbook(book => book.addWorksheet('Invalid').getCell('A1').value = `{x}{@choice:Products; ${invalid}}`)).rejects.toThrow()
  }
})

it('renders labels and validates keys without guessing the first duplicate label', async () => {
  expect(workbookDictionarySources(definition)).toEqual(['products', 'suppliers'])
  const layout = resolveWorkbook(definition, data, options)
  const choice = layout.sheets[0].cells.find(cell => cell.origin.dataPath === '$data.items[0].product')!
  expect(choice.value).toEqual({ literal: '0007' })
  expect(choice.choice?.text).toBe('Service [0007]')
  const book = await load(await renderWorkbookReport(definition, data, options))
  const sheet = book.worksheets[0]
  expect(book.definedNames.model.some(entry => entry.name.startsWith('_sb_ref_'))).toBe(false)
  expect(find(sheet, 'Service [0007]').dataValidation.type).toBe('list')
  const form = await load(await renderWorkbookForm(definition, data, options))
  const selected = find(form.worksheets[0], 'Service [0007]')
  selected.value = 'Service [0008]'
  expect(await readWorkbookForm(definition, await saveWorkbook(form))).toEqual({ success: true, data: { ...declaredData, items: [{ ...declaredData.items[0], product: '0008' }, declaredData.items[1]] } })
  selected.value = 'Service'
  expect(codes(await readWorkbookForm(definition, await saveWorkbook(form)))).toContain('choice')
  const report = await load(await renderWorkbookReport(definition, { ...data, category: 'unknown' }, options))
  expect(find(report.worksheets[0], 'unknown').value).toBe('unknown')
  expect(codes(await readWorkbookForm(definition, await renderWorkbookForm(definition, { ...data, category: 'unknown' }, options)))).toContain('choice')
})

it('renders unmatched input but rejects issuing a form when its text would decode to another choice', async () => {
  for (const mode of ['key', 'object'] as const) {
    const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = `{value}{@choice:Options; key=id; label=name; return=${mode}}`)
    const options = { dictionaries: { Options: [{ id: '001', name: 'Allowed' }] } }
    const data = { value: mode === 'key' ? 'Allowed' : { id: '002', name: 'Allowed' } }
    expect((await load(await renderWorkbookReport(template, data, options))).worksheets[0].getCell('A1').value).toBe('Allowed')
    await expect(renderWorkbookForm(template, data, options)).rejects.toMatchObject({ issues: [{ code: 'ambiguous-choice', path: '$data.value', sheetName: 'Input', address: 'A1' }] })
  }
})

it('reads choices from shared sources after sorting and insertion', async () => {
  const book = await load(await renderWorkbookForm(definition, data, options))
  const sheet = book.worksheets[0]
  const row = Number(find(sheet, 'line-a').row)
  const first = sheet.getRow(row).values
  sheet.getRow(row).values = sheet.getRow(row + 1).values
  sheet.getRow(row + 1).values = first
  sheet.spliceRows(row + 2, 0, [])
  sheet.getCell(row + 2, 1).value = 'line-c'
  sheet.getCell(row + 2, 2).value = 'Service [0008]'
  sheet.getCell(row + 2, 3).value = 'Local supplier'
  sheet.getCell(row + 2, 4).value = 4
  const result = await readWorkbookForm(definition, await saveWorkbook(book))
  expect(result).toEqual({ success: true, data: { category: 'maintenance', items: [declaredData.items[1], declaredData.items[0], { id: 'line-c', product: '0008', supplier: 7, hours: 4 }] } })
})

it('reads stored source payloads independently of dropdown labels or external context', async () => {
  const book = await load(await renderWorkbookForm(definition, data, options))
  const sheet = book.worksheets[0]
  book.worksheets[1].getCell('A1').value = 'Forged'
  expect(await readWorkbookForm(definition, await saveWorkbook(book))).toEqual({ success: true, data: declaredData })
  find(sheet, 'Maintenance').value = 'Forged'
  expect(codes(await readWorkbookForm(definition, await saveWorkbook(book)))).toContain('choice')
})

it('resolves independent mappings and retains original values when a render projection is unusable', async () => {
  const config = await editExample('choices/template.xlsx', book => {
    const sheet = book.worksheets[0]
    sheet.getCell('E4').value = sheet.getCell('C4').text.replace('label=name', 'label=code')
    sheet.getCell('E5').value = sheet.getCell('D5').value
    sheet.getCell('D5').value = null
  })
  const book = await load(await renderWorkbookForm(config, data, options))
  const sheet = book.worksheets[0]
  expect(find(sheet, 'Local supplier').dataValidation.formulae).not.toEqual(find(sheet, 'LOCAL').dataValidation.formulae)
  expect(await readWorkbookForm(config, await saveWorkbook(book))).toEqual({ success: true, data: declaredData })
  const duplicate = { ...dictionaries, suppliers: [{ id: 7, name: 'One' }, { id: 7, name: 'Two' }] }
  const collision = { ...dictionaries, suppliers: [{ id: 7, name: 'One' }, { id: 8, name: 'oNe' }, { id: 9, name: 'One [7]' }] }
  const warning = vi.spyOn(console, 'warn').mockImplementation(() => {})
  try {
    const partial = { ...dictionaries, suppliers: dictionaries.suppliers.map(item => ({ ...item, code: null })) }
    const bytes = await renderWorkbookForm(config, data, { dictionaries: partial })
    const output = (await load(bytes)).worksheets[0]
    expect(find(output, 'Local supplier').dataValidation.type).toBe('list')
    expect(find(output, 7).dataValidation).toBeUndefined()
    expect(await readWorkbookForm(config, bytes)).toEqual({ success: true, data: declaredData })
    for (const source of [duplicate, collision]) {
      for (const render of [renderWorkbookReport, renderWorkbookForm]) {
        const bytes = await render(definition, data, { dictionaries: source })
        const output = await load(bytes)
        const supplier = find(output.worksheets[0], 7)
        expect(supplier.dataValidation).toBeUndefined()
        if (render === renderWorkbookForm) {
          expect(await readWorkbookForm(definition, bytes)).toEqual({ success: true, data: declaredData })
        }
      }
    }
    expect(warning).toHaveBeenCalledWith(expect.stringContaining('Choice keys must be unique'))
    expect(warning).toHaveBeenCalledWith(expect.stringContaining('Choice display labels are ambiguous'))
  }
  finally {
    warning.mockRestore()
  }
})

it('supports empty optional choices', async () => {
  const config = await editExample('choices/template.xlsx', book => {
    book.worksheets[0].getCell('B1').value = book.worksheets[0].getCell('B1').text.replace('{category}', '{?category}')
  })
  const empty = { ...data, category: null, categories: [] }
  const book = await load(await renderWorkbookForm(config, empty, options))
  expect(book.worksheets[0].getCell('B1').dataValidation).toBeUndefined()
  expect(await readWorkbookForm(config, await saveWorkbook(book))).toEqual({ success: true, data: { ...declaredData, category: null } })
})

it('keeps issued choice sources separate from returned data even when their original paths overlap', async () => {
  const config = await importAuthoredWorkbook(book => book.addWorksheet('Input').addRows([
    ['{#items}'], ['{.name}', '{.selected}{@choice:$root.items; key=id; label=name; return=key}'], [null, '{/items}'],
  ]))
  const data = { items: [{ id: '006', name: 'First', selected: '006' }] }
  const output = await load(await renderWorkbookForm(config, data))
  output.worksheets[0].getCell('A2').value = 'Edited'
  expect(await readWorkbookForm(config, await saveWorkbook(output))).toEqual({ success: true, data: { items: [{ name: 'Edited', selected: '006' }] } })
  await expect(editExample('choices/template.xlsx', book => {
    book.worksheets[0].getCell('A3').value = '{#items | key=product}'
  })).rejects.toThrow('Unknown region option')
})

it('writes more than 256 contextual dropdown sources through the report API', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Order').addRows([
    ['{#items}'], ['{.code}{@choice:.options; key=id; label=name; return=key}'], ['{/items}'],
  ]))
  const items = Array.from({ length: 257 }, (_, index) => ({ code: index, options: [{ id: index, name: `Value ${index}` }] }))
  const book = await load(await renderWorkbookReport(template, { items }))
  expect(book.definedNames.model).toHaveLength(257)
  expect(book.getWorksheet('Order')!.getCell('A257').dataValidation.type).toBe('list')
  expect(book.getWorksheet('_sheetbind_lists')!.getCell(1, 257).value).toBe('Value 256')
})

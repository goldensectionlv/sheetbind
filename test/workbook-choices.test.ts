import { expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import { definition, data, declaredData, dictionaries } from '../examples/choices/definition'
import { WorkbookTemplate } from '../src/xlsx/template'
import { editExample } from './xlsx'
import { workbookCells } from '../src/grid/workbook'
import { workbookDictionarySources, resolveWorkbook, importWorkbookXlsx, renderWorkbookReport } from '../src/xlsx/workbook-template'
import { renderWorkbookForm, readWorkbookForm } from '../src/xlsx/workbook-form'
import { saveWorkbook } from './xlsx'
import { writeWorkbookDropdowns } from '../src/xlsx/workbook-dropdowns'

async function load(bytes: Buffer) {
  const book = new ExcelJS.Workbook()
  await book.xlsx.load(Uint8Array.from(bytes).buffer)
  return book
}
const bytes = saveWorkbook
const options = { dictionaries, context: data }
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
  const config = await importWorkbookXlsx(await bytes(book))
  const rules = workbookCells(WorkbookTemplate.content(config).definition.sheets[0]).map(cell => cell.rules)
  expect(rules).toEqual([
    { validation: [{ rule: 'required' }, { rule: 'object' }], choice: { source: { dictionary: 'Products' }, key: 'identity.id', label: 'display.name' } },
    { validation: [{ rule: 'string' }], choice: { ...{ source: { path: 'products' }, key: 'code', label: 'name' }, return: 'key' as const } },
    { validation: [{ rule: 'number' }], choice: { ...{ source: { path: 'suppliers', from: 'root' }, key: 'id', label: 'name' }, return: 'key' as const } },
  ])

})

it('preserves choice rules through tagged XLSX and submitted values', async () => {
  expect(workbookDictionarySources(definition)).toEqual(['products', 'suppliers'])
  const form = await renderWorkbookForm(definition, data, options)
  expect(await readWorkbookForm(definition, form, options)).toEqual({ success: true, data: declaredData })
})

it('renders labels and validates keys without guessing the first duplicate label', async () => {
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
  expect(await readWorkbookForm(definition, await bytes(form), options)).toEqual({ success: true, data: { ...declaredData, items: [{ ...declaredData.items[0], product: '0008' }, declaredData.items[1]] } })
  selected.value = 'Service'
  expect(codes(await readWorkbookForm(definition, await bytes(form), options))).toContain('choice')
  expect(() => resolveWorkbook(definition, { ...data, category: 'unknown' }, options)).toThrow('select a key')
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
  const result = await readWorkbookForm(definition, await bytes(book), { dictionaries, context: { categories: data.categories } })
  expect(result).toEqual({ success: true, data: { category: 'maintenance', items: [declaredData.items[1], declaredData.items[0], { id: 'line-c', product: '0008', supplier: 7, hours: 4 }] } })
})

it('does not trust uploaded helper dictionaries and requires application context for data sources', async () => {
  const book = await load(await renderWorkbookForm(definition, data, options))
  const sheet = book.worksheets[0]
  book.worksheets[1].getCell('A1').value = 'Forged'
  expect(await readWorkbookForm(definition, await bytes(book), options)).toEqual({ success: true, data: declaredData })
  expect(codes(await readWorkbookForm(definition, await bytes(book), { dictionaries }))).toContain('choice-source')
  find(sheet, 'Maintenance').value = 'Forged'
  expect(codes(await readWorkbookForm(definition, await bytes(book), options))).toContain('choice')
})

it('resolves independent mappings of one source and rejects key or display ambiguity', async () => {
  const config = await editExample('choices/template.xlsx', book => {
    const sheet = book.worksheets[0]
    sheet.getCell('E4').value = sheet.getCell('C4').text.replace('label=name', 'label=code')
    sheet.getCell('E5').value = sheet.getCell('D5').value
    sheet.getCell('D5').value = null
  })
  const book = await load(await renderWorkbookForm(config, data, options))
  const sheet = book.worksheets[0]
  expect(find(sheet, 'Local supplier').dataValidation.formulae).not.toEqual(find(sheet, 'LOCAL').dataValidation.formulae)
  expect(await readWorkbookForm(config, await bytes(book), options)).toEqual({ success: true, data: declaredData })
  const duplicate = { ...dictionaries, suppliers: [{ id: 7, name: 'One' }, { id: 7, name: 'Two' }] }
  await expect(renderWorkbookForm(definition, data, { dictionaries: duplicate })).rejects.toThrow('keys must be unique')
  const collision = { ...dictionaries, suppliers: [{ id: 7, name: 'One' }, { id: 8, name: 'One' }, { id: 9, name: 'One [7]' }] }
  await expect(renderWorkbookForm(definition, data, { dictionaries: collision })).rejects.toThrow('ambiguous')
})

it('supports empty optional choices', async () => {
  const config = await editExample('choices/template.xlsx', book => {
    book.worksheets[0].getCell('B1').value = book.worksheets[0].getCell('B1').text.replace('{category}', '{?category}')
  })
  const empty = { ...data, category: null, categories: [] }
  const book = await load(await renderWorkbookForm(config, empty, options))
  expect(book.worksheets[0].getCell('B2').dataValidation).toMatchObject({ type: 'custom', formulae: ['FALSE'], allowBlank: true })
  expect(await readWorkbookForm(config, await bytes(book), { dictionaries, context: empty })).toEqual({ success: true, data: { ...declaredData, category: null } })
})

it('rejects choice sources that overlap submitted fields and removed repeat keys', async () => {
  const config = await editExample('choices/template.xlsx', book => {
    book.worksheets[0].getCell('B4').value = '{.product}{@choice:$root.items; key=id; label=name; return=key}'
  })
  await expect(renderWorkbookForm(config, data, options)).rejects.toThrow('separate from submitted')
  await expect(editExample('choices/template.xlsx', book => {
    book.worksheets[0].getCell('A3').value = '{#items | key=product}'
  })).rejects.toThrow('Unknown region option')
})

it('writes more than 256 contextual dropdown sources', () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Order')
  const targets = Array.from({ length: 257 }, (_, index) => ({ sheet, address: `A${index + 1}`, rules: {}, items: [`Value ${index}`] }))
  expect(writeWorkbookDropdowns(book, targets)).toBeDefined()
  expect(book.definedNames.model).toHaveLength(257)
  expect(sheet.getCell('A257').dataValidation.type).toBe('list')
})

import { expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { importWorkbookXlsx, readWorkbookForm, renderWorkbookForm, renderWorkbookReport, resolveWorkbook, workbookChoiceRange } from '../src/index'
import { importAuthoredWorkbook, openWorkbook, saveWorkbook } from './xlsx'

it.each(['_sb_list_1', '_SB_LIST_1', '_sb_object_sources', '_SB_OBJECT_SOURCES'])('preserves authored names through the public report and form writers (%s)', async name => {
  const template = await importAuthoredWorkbook(book => {
    book.addWorksheet('Input').addRow(['{status}{@list:Statuses}', 42, { formula: `SUM(${name})` }])
    book.definedNames.add('Input!$B$1', name)
    book.addWorksheet('_SHEETBIND_LISTS').getCell('A1').value = 'Authored sheet'
  })
  const options = { dictionaries: { Statuses: ['Open'] } }
  for (const render of [renderWorkbookReport, renderWorkbookForm, renderWorkbookForm]) {
    const bytes = await render(template, { status: 'Open' }, options)
    const book = await openWorkbook(bytes)
    expect(book.definedNames.getRanges(name).ranges).toEqual(['Input!$B$1'])
    expect(book.getWorksheet('Input')!.getCell('C1').formula).toBe(`SUM(${name})`)
    expect(book.getWorksheet('_SHEETBIND_LISTS')!.getCell('A1').value).toBe('Authored sheet')
    const names = book.definedNames.model.map(entry => entry.name.toLowerCase())
    expect(new Set(names).size).toBe(names.length)
    if (render === renderWorkbookForm) {
      expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data: { status: 'Open' } })
    }
  }
})

it('does not overwrite an authored range with a formula projection using its reserved name', async () => {
  const rule = { source: { dictionary: 'Options' }, key: 'id', label: 'name' }
  const name = workbookChoiceRange(rule, 'rate')
  const template = await importAuthoredWorkbook(book => {
    book.addWorksheet('Input').addRow(['{selected}{@choice:Options; key=id; label=name}', 42, { formula: `SUM(${name})` }])
    book.definedNames.add('Input!$B$1', name.toUpperCase())
  })
  const selected = { id: 'a', name: 'Allowed', rate: 2 }
  for (const render of [renderWorkbookReport, renderWorkbookForm]) {
    await expect(render(template, { selected }, { dictionaries: { Options: [selected] } })).rejects.toThrow('conflicts with an authored name')
  }
})

it.each(['005F', '005f'])('uses one decoded sheet identity for template import, output and form reading (%s)', async hex => {
  const book = new ExcelJS.Workbook()
  book.addWorksheet('Input').getCell('A1').value = '{name}'
  const zip = await JSZip.loadAsync(await saveWorkbook(book))
  const xml = await zip.file('xl/workbook.xml')!.async('string')
  zip.file('xl/workbook.xml', xml.replace('name="Input"', `name="Input_x${hex}_x0041_"`))
  const bytes = await zip.generateAsync({ type: 'nodebuffer' })
  const template = await importWorkbookXlsx(bytes)
  expect(resolveWorkbook(template, { name: 'Example' }).sheets.map(sheet => sheet.name)).toEqual(['Input_x0041_'])
  for (const render of [renderWorkbookForm, renderWorkbookReport]) {
    const output = await render(template, { name: 'Example' })
    const content = await JSZip.loadAsync(output)
    expect((await content.file('xl/workbook.xml')!.async('string')).match(/<sheet\b/g)).toHaveLength(1)
    expect((await openWorkbook(output)).worksheets[0].getCell('A1').value).toBe('Example')
    if (render === renderWorkbookForm) {
      expect(await readWorkbookForm(await importWorkbookXlsx(bytes), output)).toEqual({ success: true, data: { name: 'Example' } })
    }
  }
})

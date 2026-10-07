import { expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { importWorkbookXlsx, readWorkbookForm, renderWorkbookForm, renderWorkbookReport, resolveWorkbook, workbookChoiceRange } from '../src/index'
import { importAuthoredWorkbook, openWorkbook, saveWorkbook } from './xlsx'

it('preserves existing relationships when adding strings and the hidden list sheet', async () => {
  const book = new ExcelJS.Workbook()
  book.addWorksheet('Input').getCell('A1').value = 7
  const zip = await JSZip.loadAsync(await saveWorkbook(book))
  const worksheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file('xl/worksheets/sheet1.xml', worksheet.replace(/<c\b[^>]*>[\s\S]*?<\/c>/, '<c r="A1" t="inlineStr"><is><t>{text}{@list:Texts}</t></is></c>'))
  let workbook = await zip.file('xl/workbook.xml')!.async('string')
  const relations = (await zip.file('xl/_rels/workbook.xml.rels')!.async('string')).replace(/<Relationship\b[^>]*\/>/g, node => {
    const id = /\bId="([^"]+)"/.exec(node)![1]
    const replacement = node.includes('/worksheet"') ? 'sheetbindStrings' : node.includes('/styles"') ? 'sheetbind2' : 'sheetbindStrings_'
    workbook = workbook.replace(`r:id="${id}"`, `r:id="${replacement}"`)
    return node.replace(`Id="${id}"`, `Id = '${replacement}'`)
  })
  zip.file('xl/workbook.xml', workbook)
  zip.file('xl/_rels/workbook.xml.rels', relations)
  const source = await zip.generateAsync({ type: 'nodebuffer' })
  const template = await importWorkbookXlsx(source)
  for (const render of [renderWorkbookReport, renderWorkbookForm]) {
    const bytes = await render(template, { text: '0007' }, { dictionaries: { Texts: ['0007'] } })
    const saved = await JSZip.loadAsync(bytes)
    const ids = [...(await saved.file('xl/_rels/workbook.xml.rels')!.async('string')).matchAll(/\bId\s*=\s*(["'])(.*?)\1/g)].map(match => match[2])
    expect(new Set(ids).size).toBe(ids.length)
    const output = await openWorkbook(bytes)
    expect(output.worksheets.map(sheet => sheet.name)).toEqual(['Input', '_sheetbind_lists'])
    expect(output.getWorksheet('Input')!.getCell('A1').value).toBe('0007')
    expect(output.getWorksheet('Input')!.getCell('A1').dataValidation.type).toBe('list')
    if (render === renderWorkbookForm) {
      expect(await readWorkbookForm(await importWorkbookXlsx(source), bytes)).toEqual({ success: true, data: { text: '0007' } })
    }
  }
})

it.each(['_sb_list_1', '_SB_LIST_1', '_sb_object_sources', '_SB_OBJECT_SOURCES'])('preserves authored names through the public report and form writers (%s)', async name => {
  const template = await importAuthoredWorkbook(book => {
    book.addWorksheet('Input').addRow(['{status}{@list:Statuses}', 42, { formula: `SUM(${name})` }])
    book.definedNames.add('Input!$B$1', name)
    book.addWorksheet('_SHEETBIND_LISTS').getCell('A1').value = 'Authored sheet'
    book.addWorksheet('Second').getCell('C4').value = '{other}{@list:Statuses}'
  })
  const options = { dictionaries: { Statuses: ['Open'] } }
  for (const render of [renderWorkbookReport, renderWorkbookForm, renderWorkbookForm]) {
    const bytes = await render(template, { status: 'Open', other: 'Open' }, options)
    const book = await openWorkbook(bytes)
    expect(book.definedNames.getRanges(name).ranges).toEqual(['Input!$B$1'])
    expect(book.getWorksheet('Input')!.getCell('C1').formula).toBe(`SUM(${name})`)
    expect(book.getWorksheet('_SHEETBIND_LISTS')!.getCell('A1').value).toBe('Authored sheet')
    const first = book.getWorksheet('Input')!.getCell('A1').dataValidation.formulae
    expect(book.getWorksheet('Second')!.getCell('C4').dataValidation.formulae).toEqual(first)
    const names = book.definedNames.model.map(entry => entry.name.toLowerCase())
    expect(new Set(names).size).toBe(names.length)
    if (render === renderWorkbookForm) {
      expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data: { status: 'Open', other: 'Open' } })
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

it.each(['rows', 'columns'])('writes declared formula sources independently of %s repeat instances', async axis => {
  const rule = { source: { dictionary: 'Options' }, key: 'id', label: 'name' }
  const rate = workbookChoiceRange(rule, 'Unit_Rate')
  const missing = workbookChoiceRange(rule, 'missing')
  const text = workbookChoiceRange(rule)
  const options = { dictionaries: { Options: [{ id: 'a', name: 'Option' }, { id: 'b', name: 'Option', Unit_Rate: 4 }] } }
  const template = await importAuthoredWorkbook(book => {
    book.addWorksheet('Input').addRows([
      [`{#items | axis=${axis}}`], ['{.selected}{@choice:Options; key=id; label=name; return=key}'], ['{/items}'],
    ])
    book.addWorksheet('Summary').addRow([{ formula: `SUM(${rate})` }, { formula: `COUNTA(${text})` }, { formula: `SUM(${missing})` }])
  })
  for (const count of [0, 1, 3]) {
    const data = { items: Array.from({ length: count }, () => ({ selected: 'a' })) }
    for (const render of axis === 'rows' ? [renderWorkbookReport, renderWorkbookForm] : [renderWorkbookReport]) {
      const bytes = await render(template, data, options)
      const output = await openWorkbook(bytes)
      const helper = output.getWorksheet('_sheetbind_lists')!
      const column = (header: string) => (helper.getRow(1).values as unknown[]).findIndex(value => value === header)
      expect(output.definedNames.getRanges(rate).ranges).toHaveLength(1)
      expect(output.definedNames.getRanges(missing).ranges).toHaveLength(1)
      expect(helper.getCell(2, column('Unit_Rate')).value).toBeNull()
      expect(helper.getCell(3, column('Unit_Rate')).value).toBe(4)
      expect(helper.getCell(2, column('missing')).value).toBeNull()
      expect(helper.getCell(2, column('Selection')).value).toBe('Option [a]')
      expect(helper.getCell(3, column('Selection')).value).toBe('Option [b]')
      const names = output.definedNames.model.map(entry => entry.name.toLowerCase())
      expect(new Set(names).size).toBe(names.length)
      if (render === renderWorkbookForm) {
        expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data })
      }
    }
  }
})

it('keeps requested dictionary columns when an empty source allows free input', async () => {
  const rule = { source: { dictionary: 'Options' }, key: 'id', label: 'name' }
  const rate = workbookChoiceRange(rule, 'rate')
  const text = workbookChoiceRange(rule)
  const template = await importAuthoredWorkbook(book => {
    book.addWorksheet('Input').addRow([
      '{selected}{@choice:Options; key=id; label=name; return=key; emptySource=input}',
      { formula: `SUM(${rate})+COUNTA(${text})` },
    ])
  })
  for (const render of [renderWorkbookReport, renderWorkbookForm]) {
    const bytes = await render(template, { selected: 'Custom' }, { dictionaries: { Options: [] } })
    const output = await openWorkbook(bytes)
    expect(output.definedNames.getRanges(rate).ranges).toEqual(["'_sheetbind_lists'!$B$2"])
    expect(output.definedNames.getRanges(text).ranges).toEqual(["'_sheetbind_lists'!$A$2"])
    expect(output.worksheets[0].getCell('A1').dataValidation).toBeUndefined()
    if (render === renderWorkbookForm) {
      expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data: { selected: 'Custom' } })
    }
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

it.each([
  { encoded: ['Input_x005F_x0041_', 'Input_x0041_'], names: ['Input_x0041_', 'InputA'] },
  { encoded: ['Input_x0041_', 'Input_x005F_x0041_'], names: ['InputA', 'Input_x0041_'] },
  { encoded: ['A'.repeat(24) + '_x005F_x0041_', 'A'.repeat(24) + '_x005F_x0042_'], names: ['A'.repeat(24) + '_x0041_', 'A'.repeat(24) + '_x0042_'] },
  { encoded: ['Input_x0009_A', 'Input_x0020_B'], names: ['Input\tA', 'Input B'] },
])('decodes all sheet names before checking their uniqueness and length ($names)', async ({ encoded, names }) => {
  const book = new ExcelJS.Workbook()
  book.addWorksheet('First').getCell('A1').value = '{first}'
  book.addWorksheet('Second').getCell('A1').value = '{second}'
  const zip = await JSZip.loadAsync(await saveWorkbook(book))
  const xml = await zip.file('xl/workbook.xml')!.async('string')
  zip.file('xl/workbook.xml', xml.replace('name="First"', `name="${encoded[0]}"`).replace('name="Second"', `name="${encoded[1]}"`))
  const bytes = await zip.generateAsync({ type: 'nodebuffer' })
  const original = Buffer.from(bytes)
  const data = { first: 'First value', second: 'Second value' }
  const template = await importWorkbookXlsx(bytes)
  expect(resolveWorkbook(template, data).sheets.map(sheet => sheet.name)).toEqual(names)
  for (const render of [renderWorkbookReport, renderWorkbookForm, renderWorkbookForm]) {
    const output = await render(template, data)
    const content = await JSZip.loadAsync(output)
    expect((await content.file('xl/workbook.xml')!.async('string')).match(/<sheet\b/g)).toHaveLength(2)
    if (render === renderWorkbookForm) {
      expect(await readWorkbookForm(await importWorkbookXlsx(bytes), output)).toEqual({ success: true, data })
      const sheet = await content.file('xl/worksheets/sheet1.xml')!.async('string')
      content.file('xl/worksheets/sheet1.xml', sheet.replace(/<c\b[^>]*\br="A1"[^>]*>[\s\S]*?<\/c>/, '<c r="A1" t="str"><f>""</f><v/></c>'))
      expect(await readWorkbookForm(template, await content.generateAsync({ type: 'nodebuffer' }))).toEqual({ success: true, data: { ...data, first: null } })
    }
  }
  expect(bytes).toEqual(original)
})

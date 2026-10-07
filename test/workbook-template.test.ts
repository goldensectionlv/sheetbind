import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { describe, expect, it, vi } from 'vitest'
import { importWorkbookXlsx, renderWorkbookReport, renderWorkbookForm, resolveWorkbook, TaggedXlsxError, TemplateError } from '../src/index'
import type { WorkbookTemplate } from '../src/index'
import { openWorkbook as load, saveWorkbook } from './xlsx'

describe('imported XLSX ownership', () => {
  it('rejects unsupported execution values before evaluating accessors across public entry points', async () => {
    const book = new ExcelJS.Workbook()
    book.addWorksheet('Data').getCell('A1').value = '{name}'
    const template = await importWorkbookXlsx(await saveWorkbook(book))
    const getter = vi.fn(() => 'Desk')
    const circular: Record<string, unknown> = { name: 'Desk' }
    circular.self = circular
    const invalid = [
      { name: 'Desk', unused: new Date() },
      { name: 'Desk', unused: [undefined] },
      { name: 'Desk', unused: () => 'unused' },
      { name: 'Desk', unused: Number.NaN },
      { name: 'Desk', unused: new Array(1) },
      circular,
      Object.defineProperty({}, 'name', { get: getter, enumerable: true }),
    ]
    for (const data of invalid) {
      for (const run of [resolveWorkbook, renderWorkbookReport, renderWorkbookForm]) {
        await expect(Promise.resolve().then<unknown>(() => run(template, data))).rejects.toBeInstanceOf(SyntaxError)
      }
    }
    expect(getter).not.toHaveBeenCalled()
  })

  it('reports invalid template bytes without exposing a ZIP library error as the contract', async () => {
    await expect(importWorkbookXlsx(new Uint8Array([1, 2, 3]))).rejects.toMatchObject({
      name: 'TaggedXlsxError', issues: [{ code: 'invalid-workbook', path: '$workbook' }], cause: expect.any(Error),
    })
  })

  it('locates invalid repeat sources and missing nested fields in the authored workbook', async () => {
    const book = new ExcelJS.Workbook()
    book.addWorksheet('Items').addRows([['{#items}'], ['{.name}'], ['{/items}']])
    const template = await importWorkbookXlsx(await saveWorkbook(book))
    for (const [data, code, path, address] of [[{ items: {} }, 'invalid-collection', '$data.items', 'A1'], [{ items: [{}] }, 'missing-source', '$data.items[0].name', 'A2']] as const) {
      for (const run of [() => resolveWorkbook(template, data), () => renderWorkbookReport(template, data)]) {
        try {
          await run()
          expect.fail('Expected a data error')
        }
        catch (error) {
          expect(error).toBeInstanceOf(TemplateError)
          expect(error).toBeInstanceOf(TaggedXlsxError)
          expect((error as TaggedXlsxError).issues).toMatchObject([{ code, path, sheetName: 'Items', address }])
        }
      }
    }
    await expect(renderWorkbookForm(template, { items: {} })).rejects.toMatchObject({ issues: [{ code: 'invalid-collection', sheetName: 'Items', address: 'A1' }] })
  })

  it('accepts only imported handles and owns its source bytes', async () => {
    const book = new ExcelJS.Workbook()
    book.addWorksheet('Data').getCell('A1').value = '{code}'
    const bytes = await saveWorkbook(book)
    const template = await importWorkbookXlsx(bytes)
    bytes.fill(0)
    expect((await load(await renderWorkbookReport(template, { code: '0007' }))).worksheets[0].getCell('A1').value).toBe('0007')
    expect(() => resolveWorkbook({ sheets: [] } as unknown as WorkbookTemplate, {})).toThrow('importWorkbookXlsx')
    const layout = resolveWorkbook(template, { code: '0007' })
    expect(layout.sheets[0]).not.toHaveProperty('xlsx')
    expect(layout.sheets[0].cells[0]).not.toHaveProperty('xlsx')
    Object.assign(layout.sheets[0].cells[0], { value: { literal: 'changed' } })
    expect(resolveWorkbook(template, { code: '0007' }).sheets[0].cells[0].value).toEqual({ literal: '0007' })
  })

  it('isolates resolved cell values, rules, choices and geometry from siblings and future executions', async () => {
    const book = new ExcelJS.Workbook()
    const sheet = book.addWorksheet('Items')
    sheet.addRows([
      ['{#items}'], ['{.status}{@validate:string}{@choice:statuses; key=id; label=name; return=key}'], ['{/items}'],
    ])
    sheet.getRow(2).height = 30
    sheet.getColumn(1).width = 24
    sheet.pageSetup.printArea = 'A1:B3'
    const template = await importWorkbookXlsx(await saveWorkbook(book))
    const data = { items: [{ status: 'planned' }, { status: 'planned' }] }
    const options = { dictionaries: { statuses: [{ id: 'planned', name: 'Planned', details: { rank: 0 } }] } }
    const layout = resolveWorkbook(template, data, options)
    const expected = structuredClone(layout)
    const [first, second] = layout.sheets[0].cells
    expect(first).toMatchObject({ at: { row: 1, column: 1 }, size: { rows: 1, columns: 1 }, value: { literal: 'planned' }, contextPath: '$data.items[0]' })
    expect(first.choice).toMatchObject({ text: 'Planned' })
    expect(first.definitionId).toBe(second.definitionId)
    expect(first.id).not.toBe(second.id)
    Object.assign(first.value, { literal: 'changed' })
    Object.assign(first.at, { row: 99 })
    Object.assign(first.size, { rows: 99 })
    Object.assign(first.rules!, { validation: 'number' })
    Object.assign(first.choice!.items[0], { text: 'Changed label' })
    Object.assign(first.choice!.items[0].value.details!, { rank: 99 })
    Object.assign(first.origin.iterations[0], { index: 99 })
    Object.assign(layout.sheets[0].rows![0], { height: 99 })
    Object.assign(layout.sheets[0].columns![0], { width: 99 })
    Object.assign(layout.sheets[0].print!.area!.end, { row: 99 })
    expect(second).toEqual(expected.sheets[0].cells[1])
    expect(resolveWorkbook(template, data, options)).toEqual(expected)
    expect(data.items).toEqual([{ status: 'planned' }, { status: 'planned' }])
    expect(options.dictionaries.statuses).toEqual([{ id: 'planned', name: 'Planned', details: { rank: 0 } }])
  })
})

describe('native XLSX preservation', () => {
  it('preserves merged member formatting and blank styled cells', async () => {
    const book = new ExcelJS.Workbook()
    const sheet = book.addWorksheet('Merged')
    sheet.getCell('A1').value = 'Title'
    sheet.getCell('A1').font = { bold: true }
    sheet.mergeCells('A1:C2')
    sheet.getCell('C2').style = { font: { italic: true }, border: { bottom: { style: 'thin', color: { argb: 'FF123456' } } } }
    sheet.getCell('E4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } }
    const template = await importWorkbookXlsx(await saveWorkbook(book))
    const output = await renderWorkbookReport(template, {})
    const result = new ExcelJS.Workbook()
    await result.xlsx.load(Uint8Array.from(output).buffer, { ignoreNodes: ['mergeCells'] })
    expect(result.worksheets[0].getCell('C2').style).toMatchObject(sheet.getCell('C2').style)
    expect(result.worksheets[0].getCell('E4').fill).toEqual(sheet.getCell('E4').fill)
    expect((await load(output)).worksheets[0].model.merges).toEqual(['A1:C2'])
  })
  it('retains the original default style table', async () => {
    const book = new ExcelJS.Workbook()
    book.addWorksheet('Default').getCell('A1').value = 'Default style'
    const zip = await JSZip.loadAsync(await saveWorkbook(book))
    const styles = await zip.file('xl/styles.xml')!.async('string')
    zip.file('xl/styles.xml', styles.replace(/(<cellXfs\b[^>]*>\s*<xf\b[^>]*fillId=")0/, '$11'))
    const template = await importWorkbookXlsx(await zip.generateAsync({ type: 'nodebuffer' }))
    const rendered = await JSZip.loadAsync(await renderWorkbookReport(template, {}))
    expect(await rendered.file('xl/styles.xml')!.async('string')).toMatch(/<cellXfs[^>]*>\s*<xf[^>]*fillId="1"/)
  })
  it.each(['array formula', 'validation', 'rich text', 'alignment', 'print', 'custom border', 'date format', 'sheet color'])('preserves native %s through template rendering', async feature => {
    const book = new ExcelJS.Workbook()
    const sheet = book.addWorksheet('Native')
    const cell = sheet.getCell('A1')
    cell.value = 'Text'
    if (feature === 'array formula') {
      const array = { formula: '1+1', shareType: 'array', ref: 'A1:B2' }
      cell.value = array
    }
    if (feature === 'validation') {
      cell.dataValidation = { type: 'list', formulae: ['"A,B"'] }
    }
    if (feature === 'rich text') {
      cell.value = { richText: [{ text: 'part', font: { bold: true } }] }
    }
    if (feature === 'alignment') {
      cell.alignment = { textRotation: 90 }
    }
    if (feature === 'print') {
      sheet.pageSetup.blackAndWhite = true
    }
    if (feature === 'custom border') {
      cell.border = { diagonal: { style: 'thin', up: true } }
    }
    if (feature === 'date format') {
      cell.value = 45_000
      cell.numFmt = 'yyyy-mm-dd'
    }
    if (feature === 'sheet color') {
      sheet.properties.tabColor = { argb: 'FF123456' }
    }
    const result = await load(await renderWorkbookReport(await importWorkbookXlsx(Buffer.from(await book.xlsx.writeBuffer())), {}))
    const rendered = result.worksheets[0].getCell('A1')
    if (feature === 'array formula') {
      expect(rendered.value).toMatchObject({ formula: '1+1', shareType: 'array', ref: 'A1:B2' })
    }
    if (feature === 'validation') {
      expect(rendered.dataValidation).toMatchObject(cell.dataValidation)
    }
    if (feature === 'rich text') {
      expect(rendered.value).toEqual(cell.value)
    }
    if (feature === 'alignment') {
      expect(rendered.alignment.textRotation).toBe(90)
    }
    if (feature === 'print') {
      expect(result.worksheets[0].pageSetup.blackAndWhite).toBe(true)
    }
    if (feature === 'custom border') {
      expect(rendered.border.diagonal).toMatchObject({ style: 'thin', up: true })
    }
    if (feature === 'date format') {
      expect(rendered.value).toBeInstanceOf(Date)
      expect(rendered.numFmt).toBe('yyyy-mm-dd')
    }
    if (feature === 'sheet color') {
      expect(result.worksheets[0].properties.tabColor).toEqual({ argb: 'FF123456' })
    }
  })

  it('keeps blank sheets, metadata-only rows and ordinary text containing braces', async () => {
    const book = new ExcelJS.Workbook()
    const sheet = book.addWorksheet('Blank')
    sheet.getRow(12).hidden = true
    sheet.getRow(12).getCell(1)
    sheet.getCell('A2').value = 'Before {code} after'
    const result = await load(await renderWorkbookReport(await importWorkbookXlsx(await saveWorkbook(book)), {}))
    expect(result.worksheets[0].getRow(12).hidden).toBe(true)
    expect(result.worksheets[0].getCell('A2').value).toBe('Before {code} after')
  })
})

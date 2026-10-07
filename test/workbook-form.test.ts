import { describe, expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import * as fieldsProject from '../examples/fields/template'
import { TemplateError, readWorkbookForm, renderWorkbookForm, renderWorkbookReport } from '../src/index'
import { FORM_MARKER_COLUMN, FORM_MARKER_PREFIX } from '../src/xlsx/workbook-form-markers'
import { openWorkbook as open, importAuthoredWorkbook } from './xlsx'

function authorForm(book: ExcelJS.Workbook, name = 'Form', contact = 'contact.name') {
  const sheet = book.addWorksheet(name)
  sheet.addRows([
    ['Service form'], ['{' + contact + '}{@validate:required|string}'],
    ['{#items}'], ['{.code}{@validate:string}{@list:codes}', '{.hours}{@validate:number|min:0|max:24}', '{.approved}{@validate:boolean}'],
    [null, null, '{/items}'], [], ['End'],
  ])
  sheet.mergeCells('A1:B1')
  sheet.mergeCells('A2:B2')
}
const template = await importAuthoredWorkbook(authorForm)
const options = { dictionaries: { codes: ['0007', '0008'] } }
const data = { contact: { name: 'Jordan Lee' }, items: [{ code: '0007', hours: 0, approved: false }, { code: '0008', hours: 2.5, approved: true }] }
const find = (sheet: ExcelJS.Worksheet, value: ExcelJS.CellValue) => {
  let found: ExcelJS.Cell | undefined
  sheet.eachRow(row => row.eachCell(cell => {
    if (!cell.isMerged || cell.master === cell) {
      if (cell.value === value) {
        found = cell
      }
    }
  }))
  if (!found) {
    throw new Error(`Missing cell ${JSON.stringify(value)}`)
  }
  return found
}
const mutate = async (change: (book: ExcelJS.Workbook) => void) => {
  const book = await open(await renderWorkbookForm(template, data, options))
  change(book)
  return readWorkbookForm(template, Buffer.from(await book.xlsx.writeBuffer()))
}
const codes = (result: Awaited<ReturnType<typeof readWorkbookForm>>) => {
  expect(result.success).toBe(false)
  return result.success ? [] : result.issues.map(issue => issue.code)
}

describe('shared workbook forms: definition + marked XLSX', () => {
  it('keeps form markers clear of authored columns beyond 256', async () => {
    const wide = await importAuthoredWorkbook(book => {
      book.addWorksheet('Wide').getCell(1, 300).value = '{value}'
    })
    const bytes = await renderWorkbookForm(wide, { value: 'ready' })
    const book = await open(bytes)
    const sheet = book.getWorksheet('Wide')!
    expect(sheet.getCell(1, 300).value).toBe('ready')
    expect(sheet.getRow(1).hidden).toBeFalsy()
    expect(sheet.getCell(2, 301).value).toBe(FORM_MARKER_PREFIX + '["/sheet"]')
    expect(await readWorkbookForm(wide, bytes)).toEqual({ success: true, data: { value: 'ready' } })
  })

  it('reads nested 0/1/many repeats from tagged XLSX', async () => {
    const original = fieldsProject.definition
    const settings = { dictionaries: fieldsProject.dictionaries }
    const expected = structuredClone(fieldsProject.data) as { sites: { work: { note?: string | null }[] }[] }
    expected.sites.forEach(site => site.work.forEach(item => {
      item.note ??= null
    }))
    const bytes = await renderWorkbookForm(original, fieldsProject.data, settings)
    expect(await readWorkbookForm(original, bytes)).toEqual({ success: true, data: expected })
  })

  it('omits empty lines and retains zero, false and only declared fields', async () => {
    const initial = { ...data, ignored: 'not returned', items: [{}, { code: '0007', hours: 0, approved: false, ignored: true }] }
    const bytes = await renderWorkbookForm(template, initial, options)
    expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data: {
      contact: data.contact, items: [{ code: '0007', hours: 0, approved: false }],
    } })
    expect(await readWorkbookForm(template, await renderWorkbookForm(template, { ...data, items: [] }, options))).toEqual({ success: true, data: { contact: data.contact, items: [] } })
  })

  it('issues blank required fields, restores DV at final addresses and reports concrete value paths on read', async () => {
    const bytes = await renderWorkbookForm(template, { items: [{ code: '0007', hours: -1 }] }, options)
    const book = await open(bytes)
    const sheet = book.getWorksheet('Form')!
    const code = find(sheet, '0007')
    expect(code.numFmt).toBe('@')
    expect(code.dataValidation.type).toBe('list')
    const result = await readWorkbookForm(template, bytes)
    expect(codes(result)).toEqual(['required', 'min'])
    if (!result.success) {
      expect(result.issues.map(issue => issue.path)).toEqual(['$data.contact.name', '$data.items[0].hours'])
    }
    find(sheet, -1).value = 3
    sheet.getCell('A2').value = 'Taylor Reed'
    expect(await readWorkbookForm(template, Buffer.from(await book.xlsx.writeBuffer()))).toEqual({ success: true, data: { contact: { name: 'Taylor Reed' }, items: [{ code: '0007', hours: 3, approved: null }] } })
  })

  it.each([
    ['number', (book: ExcelJS.Workbook) => {
      find(book.getWorksheet('Form')!, 0).value = 'four'
    }],
    ['max', (book: ExcelJS.Workbook) => {
      find(book.getWorksheet('Form')!, 0).value = 25
    }],
    ['list', (book: ExcelJS.Workbook) => {
      find(book.getWorksheet('Form')!, '0007').value = 'Unknown'
      book.worksheets.find(sheet => sheet.state === 'veryHidden')!.getCell('A1').value = 'Unknown'
    }],
    ['number', (book: ExcelJS.Workbook) => {
      const cell = find(book.getWorksheet('Form')!, 0)
      cell.value = new Date('2026-01-02')
      cell.numFmt = 'yyyy-mm-dd'
    }],
    ['merges', (book: ExcelJS.Workbook) => {
      book.getWorksheet('Form')!.unMergeCells('A2:B2')
    }],
    ['form-layout', (book: ExcelJS.Workbook) => {
      book.getWorksheet('Form')!.spliceRows(2, 0, [])
    }],
    ['invalid-marker', (book: ExcelJS.Workbook) => {
      book.getWorksheet('Form')!.getCell(1, FORM_MARKER_COLUMN).value = 'Broken'
    }],
    ['form-markers', (book: ExcelJS.Workbook) => {
      const sheet = book.getWorksheet('Form')!
      find(sheet, FORM_MARKER_PREFIX + '["repeat",1]').value = FORM_MARKER_PREFIX + '["repeat",999]'
    }],
  ] as const)('rejects %s and retains data only when the structure is valid', async (code, change) => {
    const result = await mutate(change)
    expect(codes(result)).toContain(code)
    if (!result.success && result.issues.every(issue => issue.phase === 'value')) {
      expect(result.data).toHaveProperty('contact.name', 'Jordan Lee')
    }
    else {
      expect(result).not.toHaveProperty('data')
    }
  })

  it.each([
    '{', 'null', '[]', '["unknown",1]',
    '["/sheet","extra"]', '["repeat"]', '["repeat","1"]', '["/repeat",0]', '["item",-1]', '["/item",1.5]', '["repeat",9007199254740992]',
  ])('rejects malformed marker %s before interpreting the form structure', async token => {
    const result = await mutate(book => {
      book.getWorksheet('Form')!.getCell(1, FORM_MARKER_COLUMN).value = FORM_MARKER_PREFIX + token
    })
    expect(result).toMatchObject({ success: false, issues: [{ phase: 'structure', code: 'invalid-marker', sheetName: 'Form', address: 'IW1' }] })
    expect(result).not.toHaveProperty('data')
  })

  it('rejects reports and truncated repeats while taking field paths from the supplied template', async () => {
    expect(codes(await readWorkbookForm(template, await renderWorkbookReport(template, data, options)))).toContain('form-markers')
    const changed = await importAuthoredWorkbook(book => authorForm(book, 'Form', 'another'))
    expect(await readWorkbookForm(changed, await renderWorkbookForm(template, data, options))).toEqual({ success: true, data: { another: data.contact.name, items: data.items } })
    expect(codes(await mutate(book => {
      const sheet = book.getWorksheet('Form')!
      sheet.spliceRows(Number(find(sheet, '0008').row) - 1, 3)
    }))).toContain('form-markers')
  })

  it('rejects invalid files and coordinates outside XLSX before layout', async () => {
    expect(codes(await readWorkbookForm(template, new Uint8Array([1, 2, 3])))).toEqual(['invalid-workbook'])
    const zip = await JSZip.loadAsync(await renderWorkbookForm(template, data, options))
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    zip.file('xl/worksheets/sheet1.xml', xml.replace(/(<row\b[^>]*\br=")1"/, '$11048577"'))
    expect(codes(await readWorkbookForm(template, await zip.generateAsync({ type: 'uint8array' })))).toEqual(['invalid-workbook'])
    zip.file('xl/worksheets/sheet1.xml', xml.replace('ref="A2:B2"', 'ref="A1:XFD1048577"'))
    expect(codes(await readWorkbookForm(template, await zip.generateAsync({ type: 'uint8array' })))).toEqual(['invalid-workbook'])
  })

  it('checks duplicate references and collection lengths across sheets without conflating record positions', async () => {
    const repeated = await importAuthoredWorkbook(book => {
      authorForm(book)
      authorForm(book, 'Copy')
    })
    const book = await open(await renderWorkbookForm(repeated, data, options))
    expect(await readWorkbookForm(repeated, Buffer.from(await book.xlsx.writeBuffer()))).toEqual({ success: true, data })
    find(book.getWorksheet('Copy')!, 'Jordan Lee').value = 'Taylor Reed'
    expect(codes(await readWorkbookForm(repeated, Buffer.from(await book.xlsx.writeBuffer())))).toContain('conflicting-field')
    const copy = book.getWorksheet('Copy')!
    copy.spliceRows(Number(find(copy, '0008').row), 1)
    expect(codes(await readWorkbookForm(repeated, Buffer.from(await book.xlsx.writeBuffer())))).toContain('conflicting-shape')
  })

  it('requires dictionaries for issuance and rejects incompatible template bindings', async () => {
    await expect(renderWorkbookForm(template, data)).rejects.toBeInstanceOf(TemplateError)
    for (const path of ['items', 'items.name']) {
      const collision = await importAuthoredWorkbook(book => authorForm(book, 'Form', path))
      for (const operation of [() => renderWorkbookForm(collision, data, options), () => readWorkbookForm(collision, new Uint8Array())]) {
        await expect(operation()).rejects.toMatchObject({ issues: [expect.objectContaining({ phase: 'template', code: 'conflicting-binding' })] })
      }
    }
    const captions = await importAuthoredWorkbook(book => book.addWorksheet('Notes').addRow(['Instructions', { formula: '1+1' }]))
    await expect(renderWorkbookForm(captions, {})).rejects.toMatchObject({ issues: [expect.objectContaining({ code: 'no-form-fields' })] })
    await expect(readWorkbookForm(captions, new Uint8Array())).rejects.toMatchObject({ issues: [expect.objectContaining({ code: 'no-form-fields' })] })
  })
})

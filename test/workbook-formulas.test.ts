import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { definition, dictionaries, data } from '../examples/formulas/definition'
import { copyWorkbookFormula, parseWorkbookFormula } from '../src/grid/workbook-formula'
import { resolveWorkbook, importWorkbookXlsx, renderWorkbookReport } from '../src/xlsx/workbook-template'
import { importAuthoredWorkbook } from './xlsx'
import { renderWorkbookForm, readWorkbookForm } from '../src/xlsx/workbook-form'

async function load(bytes: Uint8Array) {
  const book = new ExcelJS.Workbook()
  await book.xlsx.load(Uint8Array.from(bytes).buffer)
  return book
}
describe('workbook formulas', () => {
  it.each([0, 1, 3])('keeps formula meaning through tagged XLSX for %i rows, including cross-sheet totals', async count => {
    const values = { items: Array.from({ length: count }, (_, index) => ({ ...data.items[index % 2], id: `row-${index}` })) }
    const imported = definition
    const book = await load(await renderWorkbookReport(imported, values, { dictionaries }))
    const invoice = book.getWorksheet('Invoice')!
    for (let index = 0; index < count; index++) {
      expect(invoice.getCell(index + 3, 4).formula).toBe(`B${index + 3}*C${index + 3}`)
    }
    expect(invoice.getCell(count + 3, 4).formula).toBe(`IFERROR(SUM(${count ? `D3:D${count + 2}` : '#REF!'}),0)`)
    expect(book.getWorksheet('Summary')!.getCell('B1').formula).toBe(`'Invoice'!F${count + 3}`)
    expect(invoice.getCell(count + 3, 4).result).toBeUndefined()
    const bytes = await renderWorkbookForm(definition, values, { dictionaries })
    const form = await load(bytes)
    const sheet = form.getWorksheet('Invoice')!
    let total = 0
    sheet.eachRow(row => {
      if (row.getCell(1).value === 'Total') {
        total = row.number
      }
      else if (row.getCell(4).type === ExcelJS.ValueType.Formula) {
        expect(row.getCell(4).formula).toBe(`B${row.number}*C${row.number}`)
      }
    })
    expect(total).toBeGreaterThan(count + 3)
    expect(form.getWorksheet('Summary')!.getCell('B1').formula).toBe(`'Invoice'!F${total}`)
    expect(await readWorkbookForm(definition, bytes, { dictionaries })).toEqual({ success: true, data: values })
  })

  it('keeps unequal nested groups scoped and expands a grand-total range over all groups', async () => {
    const template = await importAuthoredWorkbook(book => {
      book.addWorksheet('Groups').addRows([
        [], ['{#groups}'], ['{.name}'], ['{#.lines}'],
        [null, '{.n}', null, { formula: 'B5*2' }], [null, null, null, '{/.lines}'],
        ['Total', null, null, { formula: 'IFERROR(SUM(D5:D6),0)' }], [null, null, null, '{/groups}'],
        [null, null, null, { formula: 'IFERROR(SUMIF(A3:A8,"Total",D3:D8),0)' }],
      ])
    })
    const values = { groups: [{ name: 'One', lines: [{ n: 1 }, { n: 2 }] }, { name: 'Empty', lines: [] }, { name: 'Three', lines: [{ n: 3 }] }] }
    const book = await load(await renderWorkbookReport(template, values))
    const sheet = book.worksheets[0]
    expect(sheet.getCell('D5').formula).toBe('IFERROR(SUM(D3:D4),0)')
    expect(sheet.getCell('D7').formula).toBe('IFERROR(SUM(#REF!),0)')
    expect(sheet.getCell('D9').formula).toBe('B9*2')
    expect(sheet.getCell('D10').formula).toBe('IFERROR(SUM(D9:D9),0)')
    expect(sheet.getCell('D11').formula).toBe('IFERROR(SUMIF(A2:A10,"Total",D2:D10),0)')
  })

  it('preserves strings, named fields, sheet quoting and absolute/copy semantics', () => {
    expect(copyWorkbookFormula('IF(A1="A1 and ""B2""",LOG10(B3),1E3)+$C$4+D$5+$E6', 2, 1)).toBe('IF(B3="A1 and ""B2""",LOG10(C5),1E3)+$C$4+E$5+$E8')
    expect(copyWorkbookFormula("'Sheet 2'!A1+ref_dict__rate+'O''Brien'!$B2", 2, 0)).toBe("'Sheet 2'!A3+ref_dict__rate+'O''Brien'!$B4")
    expect(() => parseWorkbookFormula('SUM(A:A)')).toThrow(/A1/)
    expect(() => parseWorkbookFormula('Table1[Amount]')).toThrow(/structured/)
    expect(() => parseWorkbookFormula('[other.xlsx]Sheet1!A1')).toThrow(/external/)
    expect(() => parseWorkbookFormula('SUM(Sheet1:Sheet3!A1)')).toThrow()
    expect(() => parseWorkbookFormula('SUM(A1#)')).toThrow(/spill/)
  })

  it('materializes shared formulas and retains array formulas with their extent', async () => {
    const book = new ExcelJS.Workbook()
    const sheet = book.addWorksheet('S')
    sheet.fillFormula('C1:C2', 'IF(A1="A1",B1,0)', [99, 99])
    const imported = await importWorkbookXlsx(Buffer.from(await book.xlsx.writeBuffer()))
    expect(resolveWorkbook(imported, {}).sheets[0].cells.map(cell => cell.value)).toEqual([{ formula: 'IF(A1="A1",B1,0)' }, { formula: 'IF(A2="A1",B2,0)' }])
    sheet.getCell('C2').value = null
    const array = { formula: 'A1:B2', shareType: 'array', ref: 'C1:D2' }
    sheet.getCell('C1').value = array
    const source = await importWorkbookXlsx(Buffer.from(await book.xlsx.writeBuffer()))
    const rendered = await load(await renderWorkbookReport(source, {}))
    expect(rendered.worksheets[0].getCell('C1').value).toMatchObject(array)
  })

  it('reads input fields independently of computed cells and their cached results', async () => {
    const book = await load(await renderWorkbookForm(definition, data, { dictionaries }))
    const sheet = book.getWorksheet('Invoice')!
    let start = 0
    sheet.eachRow(row => {
      if (row.getCell(1).value === 'service-a') {
        start = row.number
      }
    })
    sheet.getCell(start, 2).value = 5
    const formula = sheet.getCell(start, 4).formula
    expect(formula).toBe(`B${start}*C${start}`)
    sheet.getCell(start, 4).value = { formula, result: 999999 }
    const expected = { items: [{ ...data.items[0], quantity: 5 }, data.items[1]] }
    expect(await readWorkbookForm(definition, Buffer.from(await book.xlsx.writeBuffer()))).toMatchObject({ success: true, data: expected })
    sheet.getCell(start, 4).value = { formula: '1+1', result: 2 }
    expect(await readWorkbookForm(definition, Buffer.from(await book.xlsx.writeBuffer()))).toMatchObject({ success: true, data: expected })
    sheet.getCell(start, 4).value = { formula }
    sheet.getCell(start, 2).value = { formula: '1+1', result: 2 }
    expect(await readWorkbookForm(definition, Buffer.from(await book.xlsx.writeBuffer()))).toMatchObject({ success: false, issues: [expect.objectContaining({ code: 'formula' })] })
  })

  it('requests Excel recalculation instead of persisting stale formula caches', async () => {
    const bytes = await renderWorkbookReport(definition, data, { dictionaries })
    const zip = await JSZip.loadAsync(bytes)
    expect(await zip.file('xl/workbook.xml')!.async('string')).toContain('fullCalcOnLoad="1"')
  })
})

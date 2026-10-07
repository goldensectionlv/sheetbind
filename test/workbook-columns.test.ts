import { describe, expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import { data, definition } from '../examples/comparison/definition'
import { resolveWorkbook as resolveDefinition } from '../src/grid/workbook-layout'
import { WorkbookTemplate } from '../src/xlsx/workbook-template'
import { validateWorkbookDefinition } from '../src/grid/workbook-validate'
import type { WorkbookDefinition } from '../src/grid/workbook'
import { resolveWorkbook, renderWorkbookReport } from '../src/xlsx/workbook-template'
import { renderWorkbookForm } from '../src/xlsx/workbook-form'

describe('two axis report placement', () => {
  it('keeps prepared gaps, aligned totals, merged headings, widths and formulas through tagged XLSX', async () => {
    const saved = definition
    const layout = resolveWorkbook(saved, data).sheets[0]
    const at = (row: number, column: number) => layout.cells.find(cell => cell.at.row === row && cell.at.column === column)
    expect(at(4, 3)?.value).toEqual({ literal: 10 })
    expect(at(5, 5)?.value).toEqual({ literal: null })
    expect(at(6, 5)?.value).toEqual({ literal: 35 })
    expect(at(5, 6)?.value).toEqual({ formula: 'IF(E5="","",E5*$B5)' })
    expect(at(7, 8)?.value).toEqual({ formula: 'IFERROR(SUM(H4:H6),0)' })
    expect(at(8, 1)?.value).toEqual({ literal: 'Prepared comparison' })
    expect(layout.print?.area?.end).toEqual({ row: 8, column: 8 })
    const book = new ExcelJS.Workbook()
    await book.xlsx.load(Uint8Array.from(await renderWorkbookReport(saved, data)).buffer)
    expect(book.worksheets[0].getColumn(7).width).toBe(15)
    expect(book.worksheets[0].model.merges).toContain('G2:H2')
    expect(book.worksheets[1].getCell('B1').formula).toBe("IFERROR(SUM('Comparison'!D7:H7),0)")
  })

  it('pads geometry without inventing data records, including 0/1/3 instances on either axis', async () => {
    const template = definition
    for (const rows of [0, 1, 3]) {
      for (const columns of [0, 1, 3]) {
        const input = { items: data.items.slice(0, rows), offers: data.offers.slice(0, columns).map((offer, index) => ({ ...offer, lines: offer.lines.slice(0, index ? Math.max(0, rows - 1) : rows) })) }
        const layout = resolveWorkbook(template, input).sheets[0]
        const totals = layout.cells.filter(cell => 'formula' in cell.value && cell.value.formula.startsWith('IFERROR'))
        expect(totals).toHaveLength(columns)
        expect(totals.map(cell => cell.at.row)).toEqual(Array(columns).fill(4 + rows))
        expect(layout.cells.filter(cell => 'literal' in cell.value && cell.value.literal === 'Prepared comparison').map(cell => cell.at.row)).toEqual([5 + rows])
      }
    }
  })

  it('keeps two-dimensional reports outside the row-form carrier contract', async () => {
    await expect(renderWorkbookForm(definition, data)).rejects.toThrow('column repeats support reports')
  })

  it('aligns each nested row instance across columns, rather than only padding the final footer', async () => {
    const sheet = { ...WorkbookTemplate.content(definition).definition.sheets[0] }
    Reflect.deleteProperty(sheet, 'print')
    const offer = sheet.regions!.find(region => region.source.path === 'offers')!
    const lines = offer.regions![0]
    const template: WorkbookDefinition = { sheets: [{ ...sheet,
      cells: sheet.cells.filter(cell => !('literal' in cell.value && cell.value.literal === 'Prepared comparison')),
      regions: [{ ...sheet.regions!.find(region => region.source.path === 'items')!, height: 2 }, { ...offer, height: 5,
        cells: offer.cells.map(cell => cell.at.row === 4 ? { ...cell, at: { ...cell.at, row: 5 } } : cell),
        regions: [{ ...lines, height: 2, regions: [{ id: 'notes', type: 'repeat', source: { path: 'notes' }, row: 2, height: 1,
          cells: [{ id: 'note', at: { row: 1, column: 1 }, size: { rows: 1, columns: 2 }, value: { path: 'text' } }] }] }],
      }],
    }] }
    const counts = [[3, 0, 1], [0, 3, 0], [1, 1, 1]]
    const input = { ...data, offers: data.offers.map((offer, index) => ({ ...offer, lines: offer.lines.map((line, row) => ({ ...line, notes: Array.from({ length: counts[index][row] }, (_, n) => ({ text: `Note ${n}` })) })) })) }
    const cells = resolveDefinition(template, input).sheets[0].cells
    expect(cells.filter(cell => 'literal' in cell.value && ['Item A', 'Item B', 'Item C'].includes(String(cell.value.literal))).map(cell => cell.at.row)).toEqual([4, 8, 12])
    expect(cells.filter(cell => 'formula' in cell.value && cell.value.formula.startsWith('IFERROR')).map(cell => cell.at.row)).toEqual([14, 14, 14])
    expect(cells.find(cell => cell.at.row === 12 && cell.at.column === 8)?.value).toEqual({ formula: 'IF(G12="","",G12*$B12)' })
  })

  it('rejects overlapping growth bands and resolves empty cells in wide column repeats', () => {
    const region = { type: 'repeat' as const, source: { path: 'items' }, cells: [] }
    expect(() => validateWorkbookDefinition({ sheets: [{ id: 's', name: 'Sheet', cells: [], regions: [
      { ...region, id: 'a', row: 2, height: 1, column: 1, width: 2 },
      { ...region, id: 'b', row: 2, height: 2, column: 3, width: 2 },
    ] }] })).toThrow('reserve the same band')
    expect(resolveDefinition({ sheets: [{ id: 's', name: 'Sheet', cells: [], regions: [
      { ...region, id: 'a', row: 1, height: 1, column: 1, width: 129, axis: 'columns' },
    ] }] }, { items: [{}, {}] }).sheets[0].cells).toEqual([])
  })
})

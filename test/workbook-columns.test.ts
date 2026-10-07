import { describe, expect, it } from 'vitest'
import { data, definition } from '../examples/comparison/definition'
import { resolveWorkbook, renderWorkbookReport, renderWorkbookForm } from '../src/index'
import { importAuthoredWorkbook, openWorkbook } from './xlsx'

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
    const book = await openWorkbook(await renderWorkbookReport(saved, data))
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
    const template = await importAuthoredWorkbook(book => {
      const sheet = book.addWorksheet('Comparison')
      sheet.addRows([
        ['Comparison'], [], [null, null, '{#offers | axis=columns}'], [null, null, '{.name}'],
        ['{#items}'], [null, null, '{#lines}'],
        ['{.name}', '{.quantity}', '{?.price}', { formula: 'IF(C7="","",C7*$B7)' }],
        [null, null, '{#notes}'], [null, null, '{.text}'], [null, null, null, '{/notes}'],
        [null, '{/items}'], [null, null, null, '{/lines}'],
        [null, null, 'Total', { formula: 'IFERROR(SUM(D7:D12),0)' }], [null, null, null, '{/offers}'],
      ])
      sheet.mergeCells('C9:D9')
    })
    const counts = [[3, 0, 1], [0, 3, 0], [1, 1, 1]]
    const input = { ...data, offers: data.offers.map((offer, index) => ({ ...offer, lines: offer.lines.map((line, row) => ({ ...line, notes: Array.from({ length: counts[index][row] }, (_, n) => ({ text: `Note ${n}` })) })) })) }
    const cells = resolveWorkbook(template, input).sheets[0].cells
    expect(cells.filter(cell => 'literal' in cell.value && ['Item A', 'Item B', 'Item C'].includes(String(cell.value.literal))).map(cell => cell.at.row)).toEqual([4, 8, 12])
    expect(cells.filter(cell => 'formula' in cell.value && cell.value.formula.startsWith('IFERROR')).map(cell => cell.at.row)).toEqual([14, 14, 14])
    expect(cells.find(cell => cell.at.row === 12 && cell.at.column === 8)?.value).toEqual({ formula: 'IF(G12="","",G12*$B12)' })
    const sheet = (await openWorkbook(await renderWorkbookReport(template, input))).worksheets[0]
    expect([4, 8, 12].map(row => sheet.getCell(row, 1).value)).toEqual(['Item A', 'Item B', 'Item C'])
    expect([4, 6, 8].map(column => sheet.getCell(14, column).formula)).toEqual(['IFERROR(SUM(D4:D13),0)', 'IFERROR(SUM(F4:F13),0)', 'IFERROR(SUM(H4:H13),0)'])
    expect(sheet.getCell('H12').formula).toBe('IF(G12="","",G12*$B12)')
  })

  it('rejects overlapping growth bands and preserves blank cells beyond column 256', async () => {
    await expect(importAuthoredWorkbook(book => book.addWorksheet('Overlap').addRows([
      ['{#left}'], [null, null, '{#right}'], ['{.name}', null, '{.name}'],
      [null, '{/left}'], [], [null, null, null, '{/right}'],
    ]))).rejects.toThrow('reserve the same band')
    const template = await importAuthoredWorkbook(book => {
      const sheet = book.addWorksheet('Wide')
      sheet.getCell('A1').value = '{#items | axis=columns}'
      sheet.getCell('A2').value = '{.name}'
      sheet.getCell(2, 129).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFCC' } }
      sheet.getCell(3, 129).value = '{/items}'
      sheet.getCell(2, 130).value = 'Fixed'
    })
    const sheet = (await openWorkbook(await renderWorkbookReport(template, { items: [{ name: 'First' }, { name: 'Second' }] }))).worksheets[0]
    expect([1, 130, 259].map(column => sheet.getCell(1, column).value)).toEqual(['First', 'Second', 'Fixed'])
    for (const column of [129, 258]) {
      expect(sheet.getCell(1, column).value).toBeNull()
      expect(sheet.getCell(1, column).fill).toMatchObject({ fgColor: { argb: 'FFFFFFCC' } })
    }
  })
})

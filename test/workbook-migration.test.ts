import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { readWorkbookForm, renderWorkbookForm, renderWorkbookReport } from '../src/index'
import { loadTemplate, invoiceData, serviceData, comparisonData } from '../examples/migration/templates'
import { openWorkbook as open } from './xlsx'

it('renders migrated invoice rows while preserving fixed neighbours, merges and metadata', async () => {
  const config = await loadTemplate('invoice')
  for (const count of [0, 3]) {
    const data = { ...invoiceData, items: invoiceData.items.slice(0, count) }
    const height = Math.max(1, count)
    const book = await open(await renderWorkbookReport(config, data))
    const report = book.worksheets[0]
    expect(report.getCell('G4').value).toBe('Jordan Lee')
    expect(report.getCell('G4').note).toBe('One approval for the entire invoice')
    expect(report.getCell('G5').note).toBeUndefined()
    expect(report.model.merges).toContain(`F3:F${height + 4}`)
    expect(report.getCell(height + 4, 5).formula).toBe(`IFERROR(SUM(E4:E${height + 3}),0)`)
    expect(report.getCell(height + 6, 1).hyperlink).toBe('https://example.com/instructions')
    expect(report.getImages()[0].range.tl.nativeRow).toBe(height + 6)
    // The empty repeat contributes no addresses, even when its neighbour retains the row.
    expect(book.definedNames.getRanges('AmountColumn').ranges).toEqual([count ? `Invoice!$E$3:$E$${height + 3}` : 'Invoice!$E$3'])
    for (let index = 0; index < count; index++) {
      expect(report.getCell(4 + index, 1).note).toBe('Keep leading zeros')
      expect(report.getCell(4 + index, 3).dataValidation.prompt).toBe('Enter a nonnegative value')
    }
    if (!count) {
      expect(report.getCell('A4').note).toBeUndefined()
      expect(report.getCell('C4').dataValidation).toBeUndefined()
    }
    const form = await renderWorkbookForm(config, data)
    expect(await readWorkbookForm(config, form)).toEqual({ success: true, data })
  }
})

it('reads nested editable records from tagged XLSX without inventing empty records', async () => {
  const config = await loadTemplate('service')
  const bytes = await renderWorkbookReport(config, serviceData)
  const report = (await open(bytes)).worksheets[0]
  expect(report.getCell('G5').value).toBe('Planner')
  expect(report.getCell('A5').note).toBeUndefined()
  expect(report.getCell('E6').formula).toBe('IFERROR(SUM(E5:E5),0)')
  expect(report.getCell('A8').value).toBe('S-02')
  expect(report.getCell('E13').formula).toBe('IFERROR(SUM(E10:E12),0)')
  expect(report.getCell('E16').formula).toBe('IFERROR(SUMIF(D3:D14,"Subtotal",E3:E14),0)')
  expect(report.getCell('A17').hyperlink).toBe('https://example.com/instructions')
  const xml = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml')!.async('string')
  expect(xml).toContain('sqref="C10 C11 C12"')
  expect(xml).toContain('<formula>C10&gt;0</formula>')
  for (const data of [{ sites: [] }, serviceData]) {
    expect(await readWorkbookForm(config, await renderWorkbookForm(config, data))).toEqual({ success: true, data })
  }
})

it('stretches fixed headings across growing columns without duplicating their content', async () => {
  const config = await loadTemplate('comparison')
  for (const count of [0, 1, 3]) {
    const data = { offers: comparisonData.offers.slice(0, count) }
    const end = count === 3 ? 'H' : 'D'
    const report = (await open(await renderWorkbookReport(config, data))).worksheets[0]
    expect(report.model.merges).toContain(`A1:${end}1`)
    expect(report.model.merges).toContain(`A6:${end}6`)
    expect(report.getCell(`${end}4`).value).toBe('Reference')
    expect(report.getCell(`${end}7`).note).toBe('Keep this note with the footer')
    expect(report.getImages()[0].range.tl.nativeCol).toBe(count === 3 ? 7 : 3)
    for (let index = 0; index < count; index++) {
      expect(report.getCell(4, 2 + index * 2).value).toBe(data.offers[index].price)
      expect(report.getCell(4, 2 + index * 2).note).toBe('Quoted price')
    }
    if (!count) {
      expect(report.getCell('B4').note).toBeUndefined()
    }
  }
})

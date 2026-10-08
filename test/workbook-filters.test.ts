import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { importWorkbookXlsx, renderWorkbookReport, renderWorkbookForm, readWorkbookForm } from '../src/index'
import { xmlAttributes, xmlElements } from '../src/xlsx/xml'
import { saveWorkbook } from './xlsx'

function author(axis: 'rows' | 'columns'): ExcelJS.Workbook {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Items')
  sheet.addRows([
    [], [`{#items | axis=${axis}}`], ['Name', 'Amount'], ['{.name}', '{.amount}'], [], [null, '{/items}'],
  ])
  sheet.getCell('D1').value = 'Heading'
  sheet.getCell('D7').value = 'Footer'
  return book
}

async function withMetadata(book: ExcelJS.Workbook, metadata: string) {
  const zip = await JSZip.loadAsync(await saveWorkbook(book))
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file('xl/worksheets/sheet1.xml', xml.replace('</sheetData>', () => '</sheetData>' + metadata))
  return importWorkbookXlsx(await zip.generateAsync({ type: 'nodebuffer' }))
}

const data = (count: number) => ({ items: Array.from({ length: count }, (_, index) => ({ name: `Item ${index + 1}`, amount: index + 1 })) })
const attributes = (xml: string, name: string) => xmlElements(xml, name).map(node => xmlAttributes(node.split('>')[0]))

it.each(['rows', 'columns'] as const)('removes a collapsed autoFilter along %s and preserves surviving filter settings', async axis => {
  const filter = '<filterColumn colId="1" hiddenButton="0"><customFilters><customFilter operator="greaterThan" val="0"/></customFilters></filterColumn>'
  const template = await withMetadata(author(axis), `<autoFilter ref="A3:B4">${filter}</autoFilter>`)
  for (const count of [0, 1, 3]) {
    const zip = await JSZip.loadAsync(await renderWorkbookReport(template, data(count)))
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    if (!count) {
      expect(xmlElements(xml, 'autoFilter')).toEqual([])
      expect(xmlElements(xml, 'filterColumn')).toEqual([])
    }
    else {
      expect(attributes(xml, 'autoFilter')).toEqual([{ ref: axis === 'rows' ? `A2:B${count * 3}` : count === 1 ? 'A2:B3' : 'A2:F3' }])
      expect(xmlElements(xml, 'filterColumn')).toEqual([filter])
    }
  }
})

it.each(['rows', 'columns'] as const)('removes a collapsed sortState along %s and retains surviving sort options', async axis => {
  const options = axis === 'columns' ? ' columnSort="1"' : ''
  const condition = axis === 'columns' ? 'A4:B4' : 'B4:B4'
  const sort = `<sortState ref="A3:B4" caseSensitive="1"${options}><sortCondition ref="${condition}" descending="1"/></sortState>`
  const template = await withMetadata(author(axis), axis === 'rows' ? `<autoFilter ref="A3:B4">${sort}</autoFilter>` : sort)
  for (const count of [0, 1, 3]) {
    const zip = await JSZip.loadAsync(await renderWorkbookReport(template, data(count)))
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    if (!count) {
      expect(xmlElements(xml, 'sortState')).toEqual([])
      expect(xmlElements(xml, 'sortCondition')).toEqual([])
    }
    else {
      const ref = axis === 'rows' ? `A2:B${count * 3}` : count === 1 ? 'A2:B3' : 'A2:F3'
      const key = axis === 'rows' ? count === 1 ? 'B3' : `B3:B${count * 3}` : count === 1 ? 'A3:B3' : 'A3:F3'
      expect(attributes(xml, 'sortState')).toEqual([{ ref, caseSensitive: '1', ...(axis === 'columns' ? { columnSort: '1' } : {}) }])
      expect(attributes(xml, 'sortCondition')).toEqual([{ ref: key, descending: '1' }])
    }
  }
})

it('maps filters around form control rows without attaching them to empty records', async () => {
  const template = await withMetadata(author('rows'), '<autoFilter ref="A3:B4"><sortState ref="A4:B4"><sortCondition ref="B4:B4" descending="1"/></sortState></autoFilter>')
  for (const count of [0, 1, 3]) {
    const bytes = await renderWorkbookForm(template, data(count))
    const zip = await JSZip.loadAsync(bytes)
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(attributes(xml, 'autoFilter')).toEqual(count ? [{ ref: `A4:B${count * 5}` }] : [])
    expect(attributes(xml, 'sortState')).toEqual(count ? [{ ref: `A5:B${count * 5}` }] : [])
    expect(attributes(xml, 'sortCondition')).toEqual(count ? [{ ref: count === 1 ? 'B5' : 'B5:B15', descending: '1' }] : [])
    expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data: data(count) })
  }
})

it.each([false, true])('removes vanished sort keys inside a surviving range (another key: %s)', async another => {
  const book = author('columns')
  book.worksheets[0].getCell('F4').value = 'Reference'
  const sort = `<sortState ref="A4:F4" caseSensitive="1"><sortCondition ref="B4:B4" descending="1"/>${another ? '<sortCondition ref="F4:F4" customList="Reference,Other"/>' : ''}</sortState>`
  const template = await withMetadata(book, `<autoFilter ref="A3:F4">${sort}</autoFilter>`)
  const zip = await JSZip.loadAsync(await renderWorkbookReport(template, data(0)))
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  expect(attributes(xml, 'autoFilter')).toEqual([{ ref: 'A2:D3' }])
  expect(attributes(xml, 'sortState')).toEqual(another ? [{ ref: 'A3:D3', caseSensitive: '1' }] : [])
  expect(attributes(xml, 'sortCondition')).toEqual(another ? [{ ref: 'D3', customList: 'Reference,Other' }] : [])
})

it('relocates native table filters and sorts once through the same metadata path', async () => {
  const book = author('rows')
  book.worksheets[0].addTable({
    name: 'ReferenceValues', ref: 'D9', headerRow: true,
    columns: [{ name: 'Name', filterButton: true }, { name: 'Amount', filterButton: true }], rows: [['First', 1], ['Second', 2]],
  })
  const source = await JSZip.loadAsync(await saveWorkbook(book))
  const table = await source.file('xl/tables/table1.xml')!.async('string')
  source.file('xl/tables/table1.xml', table.replace('</autoFilter>', '</autoFilter><sortState ref="D10:E11"><sortCondition descending="1" ref="E10:E11"/></sortState>'))
  const template = await importWorkbookXlsx(await source.generateAsync({ type: 'nodebuffer' }))
  for (const count of [0, 1, 3]) {
    const zip = await JSZip.loadAsync(await renderWorkbookReport(template, data(count)))
    const xml = await zip.file('xl/tables/table1.xml')!.async('string')
    const row = 4 + count * 3
    expect(attributes(xml, 'table')[0].ref).toBe(`D${row}:E${row + 2}`)
    expect(attributes(xml, 'autoFilter')).toEqual([{ ref: `D${row}:E${row + 2}` }])
    expect(attributes(xml, 'sortState')).toEqual([{ ref: `D${row + 1}:E${row + 2}` }])
    expect(attributes(xml, 'sortCondition')).toEqual([{ ref: `E${row + 1}:E${row + 2}`, descending: '1' }])
  }
})

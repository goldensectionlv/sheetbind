import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { importWorkbookXlsx, renderWorkbookReport, renderWorkbookForm, readWorkbookForm } from '../src/index'
import { decodeXml, encodeXml, xmlAttributes, xmlBody, xmlElements } from '../src/xlsx/xml'
import { saveWorkbook } from './xlsx'

const sheetName = "Item's data"
const prefix = "'Item''s data'!"
const data = (count: number) => ({ items: Array.from({ length: count }, (_, index) => ({ name: `Item ${index + 1}`, amount: (index + 1) * 10 })) })

async function author(axis: 'rows' | 'columns') {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet(sheetName)
  sheet.getCell('B2').value = `{#items | axis=${axis}}`
  sheet.getCell('B3').value = '{.name}'
  sheet.getCell('C3').value = '{.amount}'
  sheet.getCell('B4').value = 'Item note'
  sheet.getCell('C4').value = 0
  sheet.getCell('C5').value = '{/items}'
  sheet.getCell(axis === 'rows' ? 'C1' : 'A3').value = 100
  sheet.getCell(axis === 'rows' ? 'C6' : 'D3').value = 500
  const wholeAxis = axis === 'rows' ? '$3:$4' : '$B:$C'
  const summary = book.addWorksheet('Summary')
  summary.getCell('A1').value = { formula: 'IFERROR(SUM(Amounts),0)' }
  summary.getCell('A2').value = { formula: `IFERROR(SUM(${prefix}${wholeAxis}),0)` }
  sheet.addConditionalFormatting({ ref: 'F1', rules: [{ type: 'expression', priority: 1, formulae: ['SUM($B$3:$C$4)>0'], style: { font: { bold: true } } }] })
  const source = await JSZip.loadAsync(await saveWorkbook(book))
  const xml = await source.file('xl/workbook.xml')!.async('string')
  const names = [
    `<definedName name="Amounts">${prefix}$B$3:$C$4</definedName>`,
    '<definedName name="LocalAmounts" localSheetId="0" hidden="1">B$3:$C4</definedName>',
    `<definedName name="FirstAmount">${prefix}$C$3</definedName>`,
    `<definedName name="ItemLines">${prefix}${wholeAxis}</definedName>`,
    `<definedName name="WithNeighbours">${prefix}${axis === 'rows' ? '$C$1:$C$6' : '$A$3:$D$3'}</definedName>`,
    `<definedName name="LiteralText">${encodeXml('"B3:C4"')}</definedName>`,
    '<definedName name="ExternalRange">[1]External!$B$3:$C$4</definedName>',
  ].join('')
  source.file('xl/workbook.xml', xml.replace('<calcPr', () => `<definedNames>${names}</definedNames><calcPr`))
  return importWorkbookXlsx(await source.generateAsync({ type: 'nodebuffer' }))
}

async function names(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes)
  const workbook = await zip.file('xl/workbook.xml')!.async('string')
  return new Map(xmlElements(workbook, 'definedName').map(node => {
    const attributes = xmlAttributes(node.split('>')[0])
    return [attributes.name, { attributes, value: decodeXml(xmlBody(node)) }]
  }))
}

it.each(['rows', 'columns'] as const)('invalidates vanished named ranges along %s and preserves surviving references', async axis => {
  const template = await author(axis)
  for (const count of [0, 1, 3]) {
    const bytes = await renderWorkbookReport(template, data(count))
    const result = await names(bytes)
    const end = axis === 'rows' ? `$C$${count * 2 + 1}` : count === 1 ? '$C$3' : '$G$3'
    const range = count ? `$B$2:${end}` : '#REF!'
    const wholeAxis = count ? axis === 'rows' ? `$2:$${count * 2 + 1}` : count === 1 ? '$B:$C' : '$B:$G' : '#REF!'
    expect(result.get('Amounts')?.value).toBe(prefix + range)
    expect(result.get('FirstAmount')?.value).toBe(prefix + (count ? '$C$2' : '#REF!'))
    expect(result.get('LocalAmounts')).toEqual({
      attributes: { name: 'LocalAmounts', localSheetId: '0', hidden: '1' },
      value: count ? `B$2:${end.replace(/\$(\d+)$/, '$1')}` : '#REF!',
    })
    expect(result.get('ItemLines')?.value).toBe(prefix + wholeAxis)
    const remaining = axis === 'rows' ? `$C$1:$C$${count * 2 + 2}` : `$A$2:$${count === 0 ? 'B' : count === 1 ? 'D' : 'H'}$2`
    expect(result.get('WithNeighbours')?.value).toBe(prefix + remaining)
    expect(result.get('LiteralText')?.value).toBe('"B3:C4"')
    expect(result.get('ExternalRange')?.value).toBe('[1]External!$B$3:$C$4')
    const zip = await JSZip.loadAsync(bytes)
    const summary = await zip.file('xl/worksheets/sheet2.xml')!.async('string')
    expect(xmlElements(summary, 'f').map(node => decodeXml(xmlBody(node)))).toEqual(['IFERROR(SUM(Amounts),0)', `IFERROR(SUM(${prefix}${wholeAxis}),0)`])
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(xmlElements(sheet, 'formula').map(node => decodeXml(xmlBody(node)))).toEqual([`SUM(${range})>0`])
  }
})

it('maps named ranges around form control rows and preserves form reading', async () => {
  const template = await author('rows')
  for (const count of [0, 1, 3]) {
    const bytes = await renderWorkbookForm(template, data(count))
    const result = await names(bytes)
    // The last bound includes hidden block and region closing rows.
    expect(result.get('Amounts')?.value).toBe(prefix + (count ? `$B$5:$C$${count * 4 + 4}` : '#REF!'))
    expect(result.get('ItemLines')?.value).toBe(prefix + (count ? `$5:$${count * 4 + 4}` : '#REF!'))
    expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data: data(count) })
  }
})

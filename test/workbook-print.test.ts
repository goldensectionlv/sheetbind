import { expect, it } from 'vitest'
import JSZip from 'jszip'
import ExcelJS from 'exceljs'
import { decodeXml, encodeXml, xmlAttributes, xmlBody, xmlElements } from '../src/xlsx/xml'
import { data, dictionaries, declaredData } from '../examples/choices/definition'
import { resolveWorkbook, importWorkbookXlsx, renderWorkbookReport } from '../src/xlsx/workbook-template'
import { renderWorkbookForm, readWorkbookForm } from '../src/xlsx/workbook-form'
import { exampleFile, openWorkbook as load, saveWorkbook, importAuthoredWorkbook } from './xlsx'

const source = await load(await exampleFile('choices/template.xlsx'))
const sheet = source.worksheets[0]
sheet.getCell('A6').value = 'Prepared by'
sheet.pageSetup = { ...sheet.pageSetup, paperSize: 8 as ExcelJS.PaperSize, orientation: 'landscape', printArea: 'A1:D6', printTitlesRow: '1:2', fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true,
  margins: { left: 0.25, right: 0.25, top: 0.7, bottom: 0.7, header: 0.3, footer: 0.3 } }
const bytes = await saveWorkbook(source)
const template = await importWorkbookXlsx(bytes)

it('moves print bounds with the report and form while preserving native page settings', async () => {
  expect(resolveWorkbook(template, data, { dictionaries }).sheets[0].print!.area!.end.row).toBe(5)
  const report = await load(await renderWorkbookReport(template, data, { dictionaries }))
  expect(report.worksheets[0].pageSetup).toMatchObject({ paperSize: 8, orientation: 'landscape', printArea: 'A1:D5', printTitlesRow: '1:2', fitToPage: true, fitToHeight: 0, horizontalCentered: true, margins: { left: 0.25, right: 0.25 } })
  const form = await renderWorkbookForm(template, data, { dictionaries })
  const book = await load(form)
  const sheet = book.worksheets[0]
  let footerRow = 0
  sheet.eachRow(row => {
    if (row.getCell(1).value === 'Prepared by') {
      footerRow = row.number
    }
  })
  expect(sheet.pageSetup.printArea).toBe('A1:D' + footerRow)
  expect(sheet.pageSetup.printTitlesRow).toBe('1:2')
  expect(sheet.pageSetup.paperSize).toBe(8)
  expect(await readWorkbookForm(template, form)).toEqual({ success: true, data: declaredData })
})

it('preserves a print area beyond column 256', async () => {
  const wide = await importAuthoredWorkbook(book => {
    const sheet = book.addWorksheet('Wide')
    sheet.getCell(1, 300).value = 'Edge'
    sheet.pageSetup.printArea = 'A1:KN1'
  })
  expect((await load(await renderWorkbookReport(wide, {}))).worksheets[0].pageSetup.printArea).toBe('A1:KN1')
})

it('retains Normal font and theme references without a style projection', async () => {
  const zip = await JSZip.loadAsync(bytes)
  const styles = await zip.file('xl/styles.xml')!.async('string')
  zip.file('xl/styles.xml', styles.replace('name val="Calibri"', 'name val="Arial"').replace('sz val="11"', 'sz val="12"').replace('<scheme val="minor"/>', '').replace('color theme="1"', 'color theme="4"'))
  const imported = await importWorkbookXlsx(await zip.generateAsync({ type: 'nodebuffer' }))
  const report = await load(await renderWorkbookReport(imported, data, { dictionaries }))
  expect(report.worksheets[0].getCell('A2').font).toMatchObject({ name: 'Arial', size: 12, color: { theme: 4 } })
})

it('moves disjoint native print areas without narrowing their geometry', async () => {
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('xl/workbook.xml')!.async('string')
  zip.file('xl/workbook.xml', xml.replace(/(<definedName name="_xlnm.Print_Area"[^>]*>)[^<]+/, (_, head: string) => head + "'Order'!$A$1:$B$2,'Order'!$C$1:$D$6"))
  const result = await JSZip.loadAsync(await renderWorkbookReport(await importWorkbookXlsx(await zip.generateAsync({ type: 'nodebuffer' })), data, { dictionaries }))
  const workbook = await result.file('xl/workbook.xml')!.async('string')
  expect(workbook).toContain('$A$1:$B$2')
  expect(workbook).toContain('$C$1:$D$5')
})

it.each(["Client's projects", "O'Neil's, plans!", 'Team plans'])('quotes print references to %s in reports and forms', async (name) => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet(name)
  sheet.getCell('A1').value = '{name}'
  sheet.getCell('B2').value = 'Footer'
  sheet.pageSetup.printArea = 'A1:B2'
  sheet.pageSetup.printTitlesRow = '1:1'
  // Supply valid source OOXML: ExcelJS itself omits apostrophe escaping here.
  const zip = await JSZip.loadAsync(await saveWorkbook(book))
  const sourceXml = await zip.file('xl/workbook.xml')!.async('string')
  zip.file('xl/workbook.xml', sourceXml.replace(/(<definedName\b[^>]*>)([^<]*)(<\/definedName>)/g, (_, head: string, value: string, tail: string) =>
    head + encodeXml(decodeXml(value).replace(`'${name}'!`, `'${name.replace(/'/g, "''")}'!`)) + tail))
  const template = await importWorkbookXlsx(await zip.generateAsync({ type: 'nodebuffer' }))
  for (const render of [renderWorkbookReport, renderWorkbookForm]) {
    const output = await render(template, { name: 'Alex' })
    const result = await JSZip.loadAsync(output)
    const xml = await result.file('xl/workbook.xml')!.async('string')
    const names = xmlElements(xml, 'definedName').filter(node => xmlAttributes(node.split('>')[0]).name.startsWith('_xlnm.Print_'))
    expect(names).toHaveLength(2)
    for (const node of names) {
      expect(decodeXml(xmlBody(node))).toMatch(new RegExp('^' + `'${name.replace(/'/g, "''")}'!`.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')))
    }
    const reopened = await load(output)
    expect(reopened.worksheets[0].pageSetup.printArea).toBe('A1:B2')
  }
})

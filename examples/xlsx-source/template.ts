import ExcelJS from 'exceljs'

export const data = { title: 'Quarterly review', items: [{ code: '0007', quantity: 2 }, { code: '0008', quantity: 0 }] }

/** An ordinary Excel-authored template with features independent of its bindings. */
export async function createTemplate(): Promise<Buffer> {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Template')
  sheet.views = [{ state: 'frozen', ySplit: 1, showGridLines: false }]
  sheet.properties.tabColor = { argb: 'FF336699' }
  sheet.headerFooter.oddHeader = 'Internal review'
  sheet.getColumn(1).width = 25
  sheet.getColumn(2).width = 15
  sheet.getCell('A1').value = { richText: [{ text: 'Review ', font: { bold: true } }, { text: 'form {instructions}', font: { italic: true } }] }
  sheet.getCell('A2').value = { richText: [{ text: '{title}' }, { text: '{@validate:required|string}' }] }
  sheet.getCell('A2').font = { name: 'Arial', underline: true, color: { argb: 'FF336699' } }
  sheet.getCell('A2').alignment = { textRotation: 15, wrapText: true }
  sheet.getCell('A2').border = { diagonal: { up: true, style: 'thin', color: { argb: 'FF336699' } } }
  sheet.getCell('A3').value = '{#items}'
  sheet.getCell('A3').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDDDDD' } }
  sheet.getCell('A4').value = '{.code}{@validate:required|string}'
  sheet.getCell('A4').note = 'Keep the complete identifier'
  sheet.getCell('B4').value = '{.quantity}{@validate:required|number|min:0}'
  sheet.getCell('B4').numFmt = '0.00'
  sheet.getCell('C4').value = { text: 'Reference', hyperlink: 'https://example.com/reference' }
  sheet.getCell('D4').dataValidation = { type: 'whole', operator: 'greaterThanOrEqual', formulae: [0], showErrorMessage: true }
  sheet.addConditionalFormatting({ ref: 'B4', rules: [{ type: 'expression', formulae: ['B4>0'], priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } } } }] })
  sheet.getCell('D5').value = '{/items}'
  sheet.getCell('A6').value = 'Footer'
  sheet.getRow(6).outlineLevel = 1
  sheet.getCell('B6').value = new Date('2026-01-02T00:00:00Z')
  sheet.getCell('B6').numFmt = 'yyyy-mm-dd'
  sheet.getCell('E7').value = 'Label'
  sheet.getCell('F7').value = 'Value'
  sheet.getCell('E8').value = 'Reference'
  sheet.getCell('F8').value = 12
  book.definedNames.add("'Template'!$B$4", 'Quantity')
  const image = book.addImage({ base64: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=', extension: 'png' })
  sheet.addImage(image, { tl: { col: 0, row: 5 }, ext: { width: 12, height: 12 } })
  return Buffer.from(await book.xlsx.writeBuffer())
}

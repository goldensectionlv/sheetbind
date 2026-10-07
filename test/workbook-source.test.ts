import { expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { createTemplate, data } from '../examples/xlsx-source/template'
import { importWorkbookXlsx, renderWorkbookReport, renderWorkbookForm, readWorkbookForm } from '../src/index'
import { xmlAttributes } from '../src/xlsx/xml'

async function open(bytes: Uint8Array) {
  const book = new ExcelJS.Workbook()
  await book.xlsx.load(Uint8Array.from(bytes).buffer)
  return book
}

it.each(['rows', 'columns'] as const)('keeps blank styled cells inside their %s repeat or fixed outside it', async axis => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Styled blanks')
  sheet.getCell('A1').value = `{#items | axis=${axis}}`
  sheet.getCell('A2').value = '{.name}'
  sheet.getCell('B3').value = '{/items}'
  const color = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFCC' } } as const
  sheet.getCell('B2').fill = color
  sheet.getCell('D2').fill = color
  sheet.getCell('B4').fill = color
  const source = Buffer.from(await book.xlsx.writeBuffer())
  const template = await importWorkbookXlsx(source)
  for (const count of [0, 1, 3]) {
    const input = { items: Array.from({ length: count }, (_, index) => ({ name: `Item ${index}` })) }
    for (const render of axis === 'rows' ? [renderWorkbookReport, renderWorkbookForm] : [renderWorkbookReport]) {
      const result = (await open(await render(template, input))).worksheets[0]
      const colors: string[] = []
      result.eachRow({ includeEmpty: true }, row => row.eachCell({ includeEmpty: true }, cell => {
        if (cell.fill?.type === 'pattern' && cell.fill.fgColor?.argb === color.fgColor.argb) {
          colors.push(cell.address)
        }
      }))
      const repeated = render === renderWorkbookForm ? Math.max(1, count) : count
      expect(colors).toHaveLength(repeated + 2)
      const offset = render === renderWorkbookForm ? 1 : 0
      const expected = axis === 'rows'
        ? [...Array.from({ length: repeated }, (_, index) => `B${index + 1 + offset}`), `D${1 + offset}`, `B${Math.max(1, repeated) + 1 + offset + (render === renderWorkbookForm ? 1 : 0)}`]
        : [...Array.from({ length: repeated }, (_, index) => `${String.fromCharCode(66 + index * 2)}1`), `${String.fromCharCode(68 + Math.max(0, repeated - 1) * 2)}1`, 'B2']
      expect(colors.sort()).toEqual(expected.sort())
      if (render === renderWorkbookForm) {
        expect(await readWorkbookForm(template, Buffer.from(await result.workbook.xlsx.writeBuffer()))).toEqual({ success: true, data: input })
      }
    }
  }
})

it('retains Excel content through tagged templates, both outputs and repeated rows', async () => {
  const bytes = await createTemplate()
  const source = await JSZip.loadAsync(bytes)
  const template = await importWorkbookXlsx(bytes)
  const report = await renderWorkbookReport(template, data)
  const zip = await JSZip.loadAsync(report)
  const book = await open(report)
  const sheet = book.worksheets[0]
  expect(sheet.getCell('A2').value).toBe(data.title)
  expect(sheet.getCell('A1').text).toBe('Review form {instructions}')
  expect(sheet.getCell('A3').value).toBe('0007')
  expect(sheet.getCell('A4').value).toBe('0008')
  expect(sheet.getCell('A3').note).toEqual(sheet.getCell('A4').note)
  expect(sheet.getCell('A3').note).toBeTruthy()
  expect(sheet.getCell('C3').hyperlink).toBe('https://example.com/reference')
  expect(sheet.getCell('C4').hyperlink).toBe(sheet.getCell('C3').hyperlink)
  expect(sheet.getCell('D3').dataValidation).toMatchObject({ type: 'whole', formulae: [0] })
  expect(sheet.getCell('D4').dataValidation).toMatchObject({ type: 'whole', formulae: [0] })
  expect(sheet.getCell('A2').font.underline).toBe(true)
  expect(sheet.getCell('A2').alignment.textRotation).toBe(15)
  expect(sheet.headerFooter.oddHeader).toBe('Internal review')
  expect(sheet.properties.tabColor).toEqual({ argb: 'FF336699' })
  expect(sheet.getCell('B5').value).toEqual(new Date('2026-01-02T00:00:00Z'))
  expect(await zip.file('xl/media/image1.png')!.async('uint8array')).toEqual(await source.file('xl/media/image1.png')!.async('uint8array'))
  expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).toContain('sqref="B3 B4"')
  expect(await zip.file('xl/drawings/drawing1.xml')!.async('string')).toContain('<xdr:row>4</xdr:row>')
  expect(book.definedNames.getRanges('Quantity').ranges).toEqual(['Template!$B$3'])
  const form = await renderWorkbookForm(template, data)
  expect(await readWorkbookForm(template, form)).toEqual({ success: true, data })
  const edited = (await open(await renderWorkbookReport(template, { ...data, title: "A $& B $' C $1" }))).worksheets[0].getCell('A2')
  expect(edited.value).toBe("A $& B $' C $1")
  expect(edited.font.underline).toBe(true)

})

it('duplicates native metadata along columns and removes it with an empty body', async () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Columns')
  sheet.getCell('A1').value = '{#items | axis=columns}'
  sheet.getCell('A2').value = '{.code}{@validate:string}'
  sheet.getCell('A2').note = 'Column note'
  sheet.getCell('B2').value = { text: 'Reference', hyperlink: 'https://example.com/reference' }
  sheet.getCell('B3').value = '{/items}'
  sheet.getCell('E2').value = { formula: 'LOG10(100)+SUM($B:$B)' }
  book.addWorksheet('Summary').getCell('A1').value = { formula: "SUM('Columns'!$B:$B)" }
  const image = book.addImage({ base64: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=', extension: 'png' })
  sheet.addImage(image, { tl: { col: 0, row: 1 }, ext: { width: 8, height: 8 } })
  const template = await importWorkbookXlsx(Buffer.from(await book.xlsx.writeBuffer()))
  const rendered = await open(await renderWorkbookReport(template, data))
  const result = rendered.worksheets[0]
  expect(result.getCell('A1').value).toBe('0007')
  expect(result.getCell('C1').value).toBe('0008')
  expect(result.getCell('A1').note).toBeTruthy()
  expect(result.getCell('C1').note).toEqual(result.getCell('A1').note)
  expect(result.getCell('D1').hyperlink).toBe(result.getCell('B1').hyperlink)
  expect(result.getImages().map(image => image.range.tl.nativeCol)).toEqual([0, 2])
  expect(result.getCell('G1').formula).toBe('LOG10(100)+SUM($B:$D)')
  expect(rendered.getWorksheet('Summary')!.getCell('A1').formula).toBe("SUM('Columns'!$B:$D)")
  const empty = (await open(await renderWorkbookReport(template, { items: [] }))).worksheets[0]
  expect(empty.getImages()).toHaveLength(0)
  expect(empty.getCell('A1').note).toBeUndefined()
})

it.each(['rows', 'columns'])('preserves backslashes in native formula names during %s growth', async axis => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Names')
  sheet.addRows([
    [`{#items | axis=${axis}}`], ['{.n}', null, null, { formula: 'IFERROR(SUM($A:$A),0)+SUM(\\A1,A1\\Rate)' }], ['{/items}'],
  ])
  sheet.getCell('Z2').value = 10
  book.definedNames.add('Names!$Z$2', '\\A1')
  book.definedNames.add('Names!$Z$2', 'A1\\Rate')
  const source = Buffer.from(await book.xlsx.writeBuffer())
  const template = await importWorkbookXlsx(source)
  for (const count of [0, 1, 3]) {
    const data = { items: Array.from({ length: count }, (_, index) => ({ n: index + 1 })) }
    for (const render of axis === 'rows' ? [renderWorkbookReport, renderWorkbookForm] : [renderWorkbookReport]) {
      const bytes = await render(template, data)
      const result = await open(bytes)
      const formulas: string[] = []
      result.worksheets[0].eachRow(row => row.eachCell(cell => {
        if (cell.formula) {
          formulas.push(cell.formula)
        }
      }))
      expect(formulas).toHaveLength(1)
      expect(formulas[0]).toContain('SUM(\\A1,A1\\Rate)')
      if (render === renderWorkbookForm) {
        expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data })
      }
    }
  }
})

it('keeps a single active cell and viewport anchor as selected rows repeat or disappear', async () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Rows', { views: [{ state: 'frozen', ySplit: 2, topLeftCell: 'A3', activeCell: 'B3', zoomScale: 85 }] })
  sheet.getCell('A1').value = 'Heading'
  sheet.getCell('A2').value = '{#items}'
  sheet.getCell('A3').value = '{.name}'
  sheet.getCell('B3').value = '{.value}'
  sheet.getCell('B4').value = '{/items}'
  sheet.getCell('A5').value = 'Footer'
  const template = await importWorkbookXlsx(Buffer.from(await book.xlsx.writeBuffer()))
  for (const count of [0, 1, 3]) {
    const input = { items: Array.from({ length: count }, (_, index) => ({ name: `Item ${index + 1}`, value: index + 1 })) }
    for (const render of [renderWorkbookReport, renderWorkbookForm]) {
      const bytes = await render(template, input)
      const saved = (await open(bytes)).worksheets[0]
      const xml = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml')!.async('string')
      const selection = xmlAttributes(xml.match(/<selection\b[^>]*>/)![0])
      const pane = xmlAttributes(xml.match(/<pane\b[^>]*>/)![0])
      const rows: number[] = []
      saved.eachRow(row => {
        if (typeof row.getCell(1).value === 'string' && row.getCell(1).text.startsWith('Item ')) {
          rows.push(row.number)
        }
      })
      expect(selection.activeCell).toMatch(/^B[1-9]\d*$/)
      expect(selection.sqref).toBe(rows.length ? rows.map(row => `B${row}`).join(' ') : selection.activeCell)
      expect(selection.activeCell).toBe(selection.sqref.split(' ')[0])
      expect(pane.topLeftCell).toBe(selection.activeCell.replace('B', 'A'))
      expect(selection.pane).toBe('bottomLeft')
      expect(xmlAttributes(xml.match(/<sheetView\b[^>]*>/)![0]).zoomScale).toBe('85')
      if (render === renderWorkbookReport) {
        expect(selection.activeCell).toBe('B2')
        expect(pane.ySplit).toBe('1')
      }
      else {
        expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data: input })
      }
    }
  }
})

it('remaps the active range index and falls back to a surviving selection after column repeats', async () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Columns', { views: [{ state: 'normal', activeCell: 'D3' }] })
  sheet.getCell('A1').value = 'Heading'
  sheet.getCell('B2').value = '{#items | axis=columns}'
  sheet.getCell('B3').value = '{.name}'
  sheet.getCell('B4').value = '{/items}'
  sheet.getCell('D3').value = 'Reference'
  sheet.getCell('D5').value = 'Footer'
  const original = Buffer.from(await book.xlsx.writeBuffer())
  for (const scenario of [
    { activeCell: 'D3', count: 3, expected: { activeCell: 'F2', sqref: 'B2 C2 D2 F2:F3', activeCellId: '3' } },
    { activeCell: 'D3', count: 1, expected: { activeCell: 'D2', sqref: 'B2 D2:D3', activeCellId: '1' } },
    { activeCell: 'D3', count: 0, expected: { activeCell: 'C2', sqref: 'C2:C3' } },
    { activeCell: 'B3', count: 0, expected: { activeCell: 'C2', sqref: 'C2:C3' } },
    { activeCell: 'B2', count: 3, expected: { activeCell: 'F2', sqref: 'F2:F3' } },
  ]) {
    const source = await JSZip.loadAsync(original)
    const xml = await source.file('xl/worksheets/sheet1.xml')!.async('string')
    const first = scenario.activeCell === 'B2' ? 'B2' : 'B3'
    source.file('xl/worksheets/sheet1.xml', xml.replace('<sheetView ', '<sheetView topLeftCell="B3" ')
      .replace(/<selection\b[^>]*\/>/, `<selection activeCell="${scenario.activeCell}" activeCellId="${scenario.activeCell === 'D3' ? 1 : 0}" sqref="${first} D3:D5"/>`))
    const template = await importWorkbookXlsx(await source.generateAsync({ type: 'nodebuffer' }))
    const bytes = await renderWorkbookReport(template, { items: Array.from({ length: scenario.count }, (_, index) => ({ name: `Item ${index + 1}` })) })
    const result = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml')!.async('string')
    expect(xmlAttributes(result.match(/<selection\b[^>]*>/)![0])).toEqual(scenario.expected)
    expect(xmlAttributes(result.match(/<sheetView\b[^>]*>/)![0]).topLeftCell).toBe('B2')
  }
})

it('reads only submitted fields while allowing Excel annotations and unrelated worksheet content', async () => {
  const source = await open(await createTemplate())
  source.addWorksheet('Reference').getCell('A1').value = 'Instructions'
  const template = await importWorkbookXlsx(Buffer.from(await source.xlsx.writeBuffer()))
  const book = await open(await renderWorkbookForm(template, data))
  const sheet = book.worksheets[0]
  book.removeWorksheet('Reference')
  let title: ExcelJS.Cell | undefined
  sheet.eachRow(row => row.eachCell(cell => {
    if (cell.value === data.title) {
      title = cell
    }
  }))
  title!.value = { text: 'Updated title', hyperlink: 'https://example.com/updated' }
  title!.note = 'Checked by the reviewer'
  sheet.getCell('A1').value = 'Edited heading'
  sheet.getCell('Z100').value = 'Personal calculations'
  sheet.mergeCells('Z101:AA102')
  book.addWorksheet('Notes').getCell('A1').value = 'Working notes'
  expect(await readWorkbookForm(template, Buffer.from(await book.xlsx.writeBuffer()))).toEqual({ success: true, data: { ...data, title: 'Updated title' } })
})

it('keeps source sheet identities, scoped names and helper strings', async () => {
  const original = new ExcelJS.Workbook()
  original.addWorksheet('Reference').getCell('A1').value = 'Reference title'
  const status = original.addWorksheet('Status')
  status.getCell('A1').value = '{status}{@list:statuses}'
  status.pageSetup.printArea = 'A1:A1'
  const lines = original.addWorksheet('Lines')
  lines.getCell('A1').value = '{#items}'
  lines.getCell('A2').value = '{.name}'
  lines.getCell('A3').value = '{/items}'
  lines.getCell('A4').value = 'Footer'
  const source = await JSZip.loadAsync(await original.xlsx.writeBuffer())
  const xml = await source.file('xl/workbook.xml')!.async('string')
  source.file('xl/workbook.xml', xml.replace('<calcPr', '<definedNames><definedName name="Footer" localSheetId="2">Lines!$A$4</definedName><definedName name="Title" localSheetId="0">Reference!$A$1</definedName></definedNames><calcPr'))
  const template = await importWorkbookXlsx(await source.generateAsync({ type: 'nodebuffer' }))
  const data = { status: 'Accepted', items: [{ name: 'First' }, { name: 'Second' }] }
  const options = { dictionaries: { statuses: ['Accepted', 'Pending'] } }
  for (const render of [renderWorkbookReport, renderWorkbookForm]) {
    const bytes = await render(template, data, options)
    const book = await open(bytes)
    const zip = await JSZip.loadAsync(bytes)
    const workbook = await zip.file('xl/workbook.xml')!.async('string')
    expect(book.worksheets.map(sheet => sheet.name)).toEqual(['Reference', 'Status', 'Lines', '_sheetbind_lists'])
    expect(book.getWorksheet('Reference')!.getCell('A1').text).toBe('Reference title')
    const cells: Record<string, string> = {}
    book.getWorksheet('Lines')!.eachRow(row => row.eachCell(cell => {
      cells[cell.text] = cell.address
    }))
    expect(cells).toHaveProperty('First')
    expect(cells).toHaveProperty('Second')
    expect(workbook).toContain(`<definedName name="Footer" localSheetId="2">'Lines'!$A$${cells.Footer.slice(1)}</definedName>`)
    expect(workbook).toContain('<definedName name="Title" localSheetId="0">\'Reference\'!$A$1</definedName>')
    expect(workbook).toMatch(/<definedName name="_xlnm.Print_Area" localSheetId="1">/)
    expect(book.getWorksheet('Status')!.pageSetup.printArea).toBe('A1:A1')
    if (render === renderWorkbookForm) {
      expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data })
    }
  }
})

it('installs a shared string table when rendering inline-string bindings', async () => {
  const book = new ExcelJS.Workbook()
  book.addWorksheet('Values').getCell('A1').value = 7
  const source = await JSZip.loadAsync(await book.xlsx.writeBuffer())
  const xml = await source.file('xl/worksheets/sheet1.xml')!.async('string')
  source.file('xl/worksheets/sheet1.xml', xml.replace(/<c\b[^>]*>[\s\S]*?<\/c>/, '<c r="A1" t="inlineStr"><is><t>{text}</t></is></c>'))
  const bytes = await source.generateAsync({ type: 'nodebuffer' })
  expect((await JSZip.loadAsync(bytes)).file('xl/sharedStrings.xml')).toBeNull()
  const template = await importWorkbookXlsx(Buffer.from(bytes))
  const result = await renderWorkbookReport(template, { text: 'Added text _x0041_' })
  expect((await open(result)).worksheets[0].getCell('A1').text).toBe('Added text _x0041_')
  const zip = await JSZip.loadAsync(result)
  expect(await zip.file('xl/_rels/workbook.xml.rels')!.async('string')).toContain('relationships/sharedStrings')
  expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('PartName="/xl/sharedStrings.xml"')
})

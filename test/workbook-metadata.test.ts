import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { importWorkbookXlsx, renderWorkbookForm, renderWorkbookReport, readWorkbookForm } from '../src/index'
import { formatAddress, parseRange } from '../src/grid/geometry'
import { xmlAttributes, xmlElements } from '../src/xlsx/xml'
import { saveWorkbook } from './xlsx'

const pixel = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII='

it('keeps each sheet with its own growth and native metadata regardless of sheet order', async () => {
  for (const names of [['Rows', 'Columns'], ['Columns', 'Rows']]) {
    const book = new ExcelJS.Workbook()
    for (const name of names) {
      const sheet = book.addWorksheet(name)
      sheet.getCell('A1').value = `{#${name.toLowerCase()} | axis=${name.toLowerCase()}}`
      sheet.getCell('A2').value = '{.value}'
      sheet.getCell('B3').value = `{/${name.toLowerCase()}}`
      sheet.getCell('D4').value = name + ' footer'
      decorate(book, sheet, 'B2', 'comment')
      decorate(book, sheet, 'B2', 'validation')
    }
    const template = await importWorkbookXlsx(await saveWorkbook(book))
    for (const [rows, columns] of [[0, 3], [1, 1], [3, 2]]) {
      const data = { rows: Array.from({ length: rows }, (_, index) => ({ value: index + 1 })), columns: Array.from({ length: columns }, (_, index) => ({ value: (index + 1) * 10 })) }
      const bytes = await renderWorkbookReport(template, data)
      const saved = new ExcelJS.Workbook()
      await saved.xlsx.load(Uint8Array.from(bytes).buffer)
      const zip = await JSZip.loadAsync(bytes)
      expect(saved.worksheets.map(sheet => sheet.name)).toEqual(names)
      for (const name of names) {
        const sheet = saved.getWorksheet(name)!
        const vertical = name === 'Rows'
        const values = data[vertical ? 'rows' : 'columns']
        const comments = await zip.file(`xl/comments${names.indexOf(name) + 1}.xml`)!.async('string')
        const notes = xmlElements(comments, 'comment').map(node => xmlAttributes(node.split('>')[0]).ref)
        expect(notes).toEqual(values.map((_item, index) => formatAddress({ row: vertical ? index + 1 : 1, column: vertical ? 2 : index * 2 + 2 })))
        for (const [index, item] of values.entries()) {
          const row = vertical ? index + 1 : 1
          const column = vertical ? 1 : index * 2 + 1
          expect(sheet.getCell(row, column).value).toBe(item.value)
          expect(sheet.getCell(row, column + 1).dataValidation).toMatchObject({ type: 'list', formulae: ['"One,Two"'] })
        }
        expect(sheet.getCell(vertical ? rows + 1 : 2, vertical ? 4 : columns * 2 + 2).value).toBe(name + ' footer')
        if (!values.length) {
          expect(sheet.getCell('B1').dataValidation).toBeUndefined()
        }
      }
    }
  }
})

type Feature = 'validation' | 'comment' | 'conditional' | 'image'
function decorate(book: ExcelJS.Workbook, sheet: ExcelJS.Worksheet, address: string, feature: Feature) {
  const cell = sheet.getCell(address)
  if (feature === 'validation') {
    cell.dataValidation = { type: 'list', formulae: ['"One,Two"'], allowBlank: true }
  }
  if (feature === 'comment') {
    cell.note = 'Authored note'
  }
  if (feature === 'conditional') {
    sheet.addConditionalFormatting({ ref: `${address}:${address}`, rules: [{
      type: 'expression', formulae: ['TRUE'], priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } } },
    }] })
  }
  if (feature === 'image') {
    sheet.addImage(book.addImage({ base64: pixel, extension: 'png' }), {
      tl: { row: Number(cell.row) - 1, col: Number(cell.col) - 1 }, ext: { width: 8, height: 8 },
    })
  }
}

async function locations(bytes: Uint8Array, feature: Feature): Promise<string[]> {
  const zip = await JSZip.loadAsync(bytes)
  if (feature === 'image') {
    const xml = await zip.file('xl/drawings/drawing1.xml')!.async('string')
    return xmlElements(xml, 'xdr:from').map(node => formatAddress({
      row: Number(/<xdr:row>(\d+)</.exec(node)![1]) + 1, column: Number(/<xdr:col>(\d+)</.exec(node)![1]) + 1,
    })).sort()
  }
  if (feature === 'comment') {
    const xml = await zip.file('xl/comments1.xml')!.async('string')
    const refs = xmlElements(xml, 'comment').map(node => xmlAttributes(node.split('>')[0]).ref).sort()
    const vml = await zip.file('xl/drawings/vmlDrawing1.vml')!.async('string')
    const shapes = xmlElements(vml, 'v:shape').map(node => formatAddress({
      row: Number(/<x:Row>(\d+)</.exec(node)![1]) + 1, column: Number(/<x:Column>(\d+)</.exec(node)![1]) + 1,
    })).sort()
    expect(shapes).toEqual(refs)
    return refs
  }
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  return xmlElements(xml, feature === 'validation' ? 'dataValidation' : 'conditionalFormatting').flatMap(node => {
    return xmlAttributes(node.split('>')[0]).sqref.split(/\s+/).flatMap(ref => {
      const range = parseRange(ref)
      const cells = []
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let column = range.start.column; column <= range.end.column; column++) {
          cells.push(formatAddress({ row, column }))
        }
      }
      return cells
    })
  }).sort()
}

it.each(['validation', 'comment', 'conditional', 'image'] as const)('places empty-cell %s by region ownership on both axes', async feature => {
  for (const axis of ['rows', 'columns'] as const) {
    const book = new ExcelJS.Workbook()
    const sheet = book.addWorksheet('Metadata')
    sheet.getCell('A1').value = `{#items | axis=${axis}}`
    sheet.getCell('A2').value = '{.name}'
    sheet.getCell('B3').value = '{/items}'
    for (const address of ['B2', 'D2', 'B4']) {
      decorate(book, sheet, address, feature)
    }
    const template = await importWorkbookXlsx(await saveWorkbook(book))
    for (const count of [0, 1, 3]) {
      const data = { items: Array.from({ length: count }, (_, index) => ({ name: `Item ${index + 1}` })) }
      for (const render of axis === 'rows' ? [renderWorkbookReport, renderWorkbookForm] : [renderWorkbookReport]) {
        const form = render === renderWorkbookForm
        const copies = form ? Math.max(1, count) : count
        const offset = form ? 1 : 0
        const expected = axis === 'rows'
          ? [...Array.from({ length: copies }, (_, index) => `B${index + 1 + offset}`), `D${1 + offset}`, `B${Math.max(1, copies) + 1 + offset + Number(form)}`]
          : [...Array.from({ length: copies }, (_, index) => `${String.fromCharCode(66 + index * 2)}1`), `${String.fromCharCode(68 + Math.max(0, copies - 1) * 2)}1`, 'B2']
        const bytes = await render(template, data)
        expect(await locations(bytes, feature)).toEqual(expected.sort())
        if (form) {
          expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data })
        }
      }
    }
  }
})

it.each(['validation', 'comment', 'conditional', 'image'] as const)('removes %s owned only by an empty repeat', async feature => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Empty')
  sheet.getCell('A1').value = '{#items}'
  sheet.getCell('A2').value = '{.name}'
  sheet.getCell('B3').value = '{/items}'
  decorate(book, sheet, 'B2', feature)
  const template = await importWorkbookXlsx(await saveWorkbook(book))
  expect(await locations(await renderWorkbookReport(template, { items: [] }), feature)).toEqual([])
})

it('keeps native ranges outside a multi-row block out of the inserted gaps', async () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Ranges')
  sheet.getCell('B2').value = '{#items}'
  sheet.getCell('B3').value = '{.name}'
  sheet.getCell('C6').value = '{/items}'
  sheet.getCell('B7').value = 'Footer'
  for (const ref of ['C3:C5', 'E1:E7']) {
    sheet.addConditionalFormatting({ ref, rules: [{ type: 'expression', formulae: ['TRUE'], priority: 1 }] })
  }
  const template = await importWorkbookXlsx(await saveWorkbook(book))
  for (const count of [0, 3]) {
    const data = { items: Array.from({ length: count }, () => ({ name: 'Item' })) }
    const expected = count
      ? [...Array.from({ length: 9 }, (_, index) => `C${index + 2}`), 'E1', 'E2', 'E3', 'E4', 'E11']
      : ['E1', 'E2', 'E3', 'E4', 'E5']
    expect(await locations(await renderWorkbookReport(template, data), 'conditional')).toEqual(expected.sort())
  }
})

it('uses each sibling collection count when its neighbour shares the same row band', async () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Siblings')
  sheet.getCell('A1').value = '{#left}'
  sheet.getCell('D2').value = '{#right}'
  sheet.getCell('A3').value = '{.name}'
  sheet.getCell('D3').value = '{.name}'
  sheet.getCell('B4').value = '{/left}'
  sheet.getCell('E5').value = '{/right}'
  for (const address of ['B3', 'C3', 'E3']) {
    decorate(book, sheet, address, 'validation')
  }
  const template = await importWorkbookXlsx(await saveWorkbook(book))
  for (const count of [0, 1]) {
    const data = { left: [{ name: 'A' }, { name: 'B' }, { name: 'C' }], right: Array.from({ length: count }, () => ({ name: 'D' })) }
    expect(await locations(await renderWorkbookReport(template, data), 'validation')).toEqual(['B1', 'B2', 'B3', 'C1', ...(count ? ['E1'] : [])])
  }
})

it.each(['validation', 'conditional'] as const)('keeps %s ranges off control rows in multirow and nested forms', async feature => {
  for (const nested of [false, true]) {
    const book = new ExcelJS.Workbook()
    const sheet = book.addWorksheet('Blocks')
    sheet.addRows(nested
      ? [['Heading'], ['{#groups}'], ['{.name}'], ['{#.items}'], ['{.label}'], [null, '{/.items}'], ['Footer'], [null, '{/groups}'], ['End']]
      : [['Heading'], ['{#groups}'], ['{.name}'], ['Middle'], ['Footer'], [null, '{/groups}'], ['End']])
    // One range owned by the block and one passing through its boundaries.
    for (const ref of ['B3:B5', `D1:D${nested ? 9 : 7}`]) {
      if (feature === 'validation') {
        Reflect.get(sheet, 'dataValidations').add(ref, { type: 'whole', operator: 'greaterThan', formulae: [0] })
      }
      else {
        sheet.addConditionalFormatting({ ref, rules: [{ type: 'expression', formulae: ['TRUE'], priority: 1 }] })
      }
    }
    const template = await importWorkbookXlsx(await saveWorkbook(book))
    for (const count of [0, 1, 3]) {
      const input = { groups: Array.from({ length: count }, (_, index) => ({ name: `Group ${index}`, ...(nested ? { items: [{ label: 'One' }, { label: 'Two' }] } : {}) })) }
      const bytes = await renderWorkbookForm(template, input)
      const result = new ExcelJS.Workbook()
      await result.xlsx.load(Uint8Array.from(bytes).buffer)
      const cells = await locations(bytes, feature)
      expect(cells.filter(address => result.worksheets[0].getRow(parseRange(address).start.row).hidden)).toEqual([])
      expect(cells.filter(address => address.startsWith('B'))).toHaveLength(count * 3)
      expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data: input })
    }
  }
})

it('keeps nested native content in its own outer and inner iterations', async () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Nested')
  sheet.getCell('A1').value = '{#groups}'
  sheet.getCell('A2').value = '{.name}'
  sheet.getCell('A3').value = '{#.items}'
  sheet.getCell('A4').value = '{.label}'
  sheet.getCell('B5').value = '{/.items}'
  sheet.getCell('E7').value = '{/groups}'
  for (const address of ['B4', 'E4', 'E6', 'G4']) {
    decorate(book, sheet, address, 'comment')
  }
  const template = await importWorkbookXlsx(await saveWorkbook(book))
  const data = { groups: [{ name: 'Empty', items: [] }, { name: 'Filled', items: [{ label: 'One' }, { label: 'Two' }] }] }
  expect(await locations(await renderWorkbookReport(template, data), 'comment')).toEqual(['B5', 'B6', 'E2', 'E3', 'E5', 'E7', 'G2'])
})

it('retains native-only geometry when scopes are flattened for a form', async () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Scoped')
  sheet.getCell('A1').value = '{#with info}'
  sheet.getCell('A2').value = '{#items}'
  sheet.getCell('A3').value = '{.name}'
  sheet.getCell('B4').value = '{/items}'
  sheet.getCell('D5').value = '{/with}'
  decorate(book, sheet, 'D3', 'validation')
  const template = await importWorkbookXlsx(await saveWorkbook(book))
  const data = { info: { items: [{ name: 'One' }, { name: 'Two' }] } }
  const form = await renderWorkbookForm(template, data)
  expect(await locations(form, 'validation')).toEqual(['D2'])
  expect(await readWorkbookForm(template, form)).toEqual({ success: true, data })
})

it('moves both drawing anchors with their owner without stretching into neighbouring copies', async () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Drawing')
  sheet.getCell('A1').value = '{#items | axis=columns}'
  sheet.getCell('A2').value = '{.name}'
  sheet.getCell('B3').value = '{/items}'
  sheet.addImage(book.addImage({ base64: pixel, extension: 'png' }), 'A4:C4')
  const template = await importWorkbookXlsx(await saveWorkbook(book))
  const zip = await JSZip.loadAsync(await renderWorkbookReport(template, { items: [{ name: 'A' }, { name: 'B' }, { name: 'C' }] }))
  const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string')
  expect(xmlElements(drawing, 'xdr:from')).toHaveLength(1)
  expect(xmlElements(drawing, 'xdr:from')[0]).toContain('<xdr:col>0</xdr:col>')
  expect(xmlElements(drawing, 'xdr:to')[0]).toContain('<xdr:col>3</xdr:col>')
})

it('keeps named single-cell ranges fixed outside a repeat even when written with two endpoints', async () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Names')
  sheet.getCell('A1').value = '{#items}'
  sheet.getCell('A2').value = '{.name}'
  sheet.getCell('B3').value = '{/items}'
  decorate(book, sheet, 'D2', 'validation')
  const source = await JSZip.loadAsync(await saveWorkbook(book))
  const xml = await source.file('xl/workbook.xml')!.async('string')
  source.file('xl/workbook.xml', xml.replace('<calcPr', '<definedNames><definedName name="Fixed">Names!$D$2:$D$2</definedName><definedName name="Items">Names!$A$2:$B$2</definedName></definedNames><calcPr'))
  const template = await importWorkbookXlsx(await source.generateAsync({ type: 'nodebuffer' }))
  for (const count of [0, 3]) {
    const bytes = await renderWorkbookReport(template, { items: Array.from({ length: count }, () => ({ name: 'Item' })) })
    const zip = await JSZip.loadAsync(bytes)
    const rendered = await zip.file('xl/workbook.xml')!.async('string')
    expect(rendered).toContain('<definedName name="Fixed">\'Names\'!$D$1:$D$1</definedName>')
    expect(rendered).toContain(`<definedName name="Items">'Names'!${count ? '$A$1:$B$3' : '#REF!'}</definedName>`)
  }
})

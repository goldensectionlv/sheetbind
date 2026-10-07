import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'
import * as project from '../examples/regions/template'
import { TemplateError } from '../src/core/template'
import { expandWorkbookScopes } from '../src/form/workbook'
import type { WorkbookCell, WorkbookDefinition } from '../src/grid/workbook'
import { resolveWorkbook as resolveDefinition } from '../src/grid/workbook-layout'
import { WorkbookTemplate } from '../src/xlsx/template'
import { resolveWorkbook, renderWorkbookReport, importWorkbookXlsx } from '../src/xlsx/workbook-template'
import { importAuthoredWorkbook, openWorkbook } from './xlsx'

const cell = (id: string, row: number, column = 1): WorkbookCell => ({ id, at: { row, column }, size: { rows: 1, columns: 1 }, value: { literal: id } })
async function load(bytes: Uint8Array) {
  return (await openWorkbook(bytes)).worksheets[0]
}

describe('scope and repeat placement', () => {
  it.each([0, 1, 3])('owns the rectangle between two markers for %i items, including empty columns', async count => {
    const template = await importAuthoredWorkbook(book => {
      const sheet = book.addWorksheet('Corners')
      sheet.getCell('B2').value = '{#items}'
      sheet.getCell('B3').value = '{.name}'
      sheet.getCell('D5').value = '{/items}'
      sheet.getCell('F3').value = 'Fixed neighbour'
      sheet.getCell('B6').value = 'Footer'
      sheet.getRow(3).height = 27
    })
    const region = WorkbookTemplate.content(template).definition.sheets[0].regions![0]
    expect(region).toMatchObject({ column: 2, width: 3, height: 2 })
    const report = await load(await renderWorkbookReport(template, { items: Array.from({ length: count }, (_, index) => ({ name: `Item ${index}` })) }))
    expect(report.getCell('F2').value).toBe('Fixed neighbour')
    expect(report.getCell(Math.max(1, count) * 2 + 2, 2).value).toBe('Footer')
    for (let index = 0; index < count; index++) {
      expect(report.getCell(index * 2 + 2, 2).value).toBe(`Item ${index}`)
      expect(report.getRow(index * 2 + 2).height).toBe(27)
    }
  })

  it('pairs adjacent repeats of the same source by their corners', async () => {
    const template = await importAuthoredWorkbook(book => {
      const sheet = book.addWorksheet('Adjacent')
      sheet.getCell('A1').value = '{#items}'
      sheet.getCell('D2').value = '{#items}'
      sheet.getCell('A3').value = '{.name}'
      sheet.getCell('D3').value = '{.name}'
      sheet.getCell('B4').value = '{/items}'
      sheet.getCell('E5').value = '{/items}'
    })
    const report = await load(await renderWorkbookReport(template, { items: [{ name: 'First' }, { name: 'Second' }] }))
    expect([report.getCell('A1').value, report.getCell('D1').value, report.getCell('A2').value, report.getCell('D2').value]).toEqual(['First', 'First', 'Second', 'Second'])
  })

  it('treats markers in one column as a one-column region', async () => {
    const template = await importAuthoredWorkbook(book => {
      book.addWorksheet('One').addRows([['{#items}'], ['{.name}', 'Fixed'], ['{/items}']])
    })
    const report = await load(await renderWorkbookReport(template, { items: [{ name: 'First' }, { name: 'Second' }] }))
    expect(report.getCell('A2').value).toBe('Second')
    expect(report.getCell('B1').value).toBe('Fixed')
    expect(report.getCell('B2').value).toBeNull()
  })

  it.each(['manual width', 'close on the left', 'merge beyond right edge'])('rejects %s instead of inferring a different rectangle', async kind => {
    await expect(importAuthoredWorkbook(book => {
      const sheet = book.addWorksheet('Invalid')
      sheet.getCell('B1').value = kind === 'manual width' ? '{#items | width=2}' : '{#items}'
      sheet.getCell('B2').value = '{.name}'
      sheet.getCell(kind === 'close on the left' ? 'A3' : 'C3').value = '{/items}'
      if (kind === 'merge beyond right edge') {
        sheet.mergeCells('B2:D2')
      }
    })).rejects.toThrow()
  })

  it('flattens scopes into paths while preserving XLSX geometry and source formatting', async () => {
    const template = await importAuthoredWorkbook(book => {
      const sheet = book.addWorksheet('Scoped')
      sheet.addRows([
        ['Title'], [], ['{#with company.provider}'], ['{.name}'], ['{$root.name}', '{?.address.city}'], [], [null, null, '{/with}'], [],
        ['{#items}'], ['{.name}'], ['{/items}'],
      ])
      sheet.mergeCells('A4:C4')
      sheet.getCell('A4').font = { bold: true }
      sheet.getCell('A4').numFmt = '@'
      sheet.getRow(4).height = 26
      sheet.getRow(6).height = 9
      sheet.getRow(6).hidden = true
    })
    const definition = WorkbookTemplate.content(template).definition
    const expanded = expandWorkbookScopes(definition)
    expect(expanded.sheets[0].regions!.map(region => region.type)).toEqual(['repeat'])
    expect(expandWorkbookScopes(expanded)).toEqual(expanded)
    for (const count of [0, 1, 3]) {
      const data = { name: 'Root name', company: { provider: { name: '0007' } }, items: Array.from({ length: count }, (_, index) => ({ name: index })) }
      const values = (cells: readonly WorkbookCell[]) => cells.map(cell => ({ at: cell.at, size: cell.size, value: cell.value }))
      expect(values(resolveDefinition(expanded, data).sheets[0].cells)).toEqual(values(resolveWorkbook(template, data).sheets[0].cells))
      const output = await load(await renderWorkbookReport(template, data))
      expect(output.getCell('A3').value).toBe('0007')
      expect(output.getCell('A4').value).toBe('Root name')
      expect(output.getCell('A3').font.bold).toBe(true)
      expect(output.model.merges).toContain('A3:C3')
      expect(output.getRow(5).hidden).toBe(true)
    }
  })

  it.each([0, 1, 3])('places %i repeated bodies with provenance, metadata and a shifted footer', async count => {
    const data = { ...project.data, work: project.data.work.slice(0, count) }
    const result = resolveWorkbook(project.definition, data)
    const sheet = result.sheets[0]
    expect(sheet.cells.find(cell => 'literal' in cell.value && cell.value.literal === 'Prepared by')!.at.row).toBe(9 + 3 * count)
    expect(sheet.rows!.find(row => row.index === 11 + 3 * count)).toMatchObject({ hidden: true, height: 9 })
    const work = sheet.cells.filter(cell => cell.origin.iterations.length)
    expect(work).toHaveLength(6 * count)
    expect(new Set(sheet.cells.map(cell => cell.id)).size).toBe(sheet.cells.length)
    if (count) {
      expect(work[0].origin).toMatchObject({ nodeId: work[0].definitionId, dataPath: '$data.work[0].code', iterations: [{ index: 0 }] })
    }
    const report = await load(await renderWorkbookReport(project.definition, data))
    expect(report.getCell('A' + (9 + 3 * count)).value).toBe('Prepared by')
    expect(report.getRow(11 + 3 * count).hidden).toBe(true)
    expect(report.getCell('B3').value).toBe('Northwind Service')
    expect(report.getCell('B4').value).toBe('SR-0042')
    expect(report.model.merges).toContain('A' + (9 + 3 * count) + ':E' + (9 + 3 * count))
    if (count) {
      expect(report.getCell('A8').value).toBe('0001')
      expect(report.getCell('E8').value).toBe(0)
      expect(report.getCell('E9').value).toBe(false)
      expect(report.getCell('A10').fill).toMatchObject({ type: 'pattern', fgColor: { argb: 'FFF1F6F4' } })
      expect(report.getRow(8).height).toBe(26)
      expect(report.getRow(10).height).toBe(8)
    }
  })

  it('keeps leading and trailing blank rows and adjacent scope bodies', async () => {
    const template = await importAuthoredWorkbook(book => {
      const sheet = book.addWorksheet('Manual')
      sheet.getCell('A2').value = '{#items}'
      sheet.getCell('B4').value = '{code}'
      sheet.getRow(3).height = 7
      sheet.getRow(5).height = 9
      sheet.getCell('B6').value = '{/items}'
      sheet.getCell('A7').value = '{#with provider}'
      sheet.getCell('A8').value = '{name}'
      sheet.getCell('A9').value = '{/with}'
      sheet.getCell('A11').value = 'Footer'
      sheet.getRow(12).hidden = true
      sheet.getRow(12).getCell(1)
    })
    const report = await load(await renderWorkbookReport(template, { items: [{ code: '005' }, { code: '006' }], provider: { name: 'Vendor' } }))
    expect(report.getCell('B3').value).toBe('005')
    expect(report.getCell('B6').value).toBe('006')
    expect(report.getCell('A8').value).toBe('Vendor')
    expect(report.getCell('A10').value).toBe('Footer')
    expect(report.getRow(11).hidden).toBe(true)
  })
  it('accumulates growth and shrinkage of independent repeats once', () => {
    const config: WorkbookDefinition = { sheets: [{ id: 's', name: 'Two', cells: [cell('footer', 12)], regions: [
      { id: 'a', type: 'repeat', row: 2, height: 2, source: { path: 'a' }, cells: [cell('a-cell', 1)] },
      { id: 'b', type: 'repeat', row: 6, height: 3, source: { path: 'b' }, cells: [cell('b-cell', 2)], rows: [{ index: 3, hidden: true }] },
    ] }] }
    for (const [a, b, footer] of [[0, 0, 7], [2, 0, 11], [0, 2, 13], [2, 2, 17]]) {
      const result = resolveDefinition(config, { a: Array.from({ length: a }, () => ({})), b: Array.from({ length: b }, () => ({})) }).sheets[0]
      expect(result.cells.find(cell => cell.definitionId === 'footer')!.at.row).toBe(footer)
    }
  })

  it.each([
    [{ ...project.data, provider: [] }, 'missing-source', '$data.provider.name'],
    [{ ...project.data, work: {} }, 'invalid-collection', '$data.work'],
    [{ ...project.data, work: [0] }, 'invalid-item', '$data.work[0]'],
    [{ ...project.data, work: [{}] }, 'missing-source', '$data.work[0].code'],
  ])('reports concrete data paths for invalid sources', (data, code, path) => {
    try {
      resolveWorkbook(project.definition, data)
      expect.fail('Expected rejection')
    }
    catch (error) {
      expect(error).toBeInstanceOf(TemplateError)
      expect((error as TemplateError).issues[0]).toMatchObject({ code, path })
    }
  })
  it.each(['crossed merge', 'empty body'])('rejects ambiguous %s at the carrier boundary', async kind => {
    const book = new ExcelJS.Workbook()
    const sheet = book.addWorksheet('Invalid')
    sheet.getCell('A1').value = '{#items}'
    sheet.getCell('A2').value = '{code}'
    sheet.getCell('A3').value = '{/items}'
    if (kind === 'crossed merge') {
      sheet.mergeCells('B1:B2')
    }
    if (kind === 'empty body') {
      sheet.getCell('A2').value = '{/items}'
      sheet.getCell('A3').value = null
    }
    await expect(importWorkbookXlsx(Buffer.from(await book.xlsx.writeBuffer()))).rejects.toThrow()
  })
})

import { describe, expect, it } from 'vitest'
import * as project from '../examples/regions/nested'
import { workbookCells, workbookRegions, workbookRows } from '../src/grid/workbook'
import { expandWorkbookScopes } from '../src/form/workbook'
import type { WorkbookDefinition } from '../src/grid/workbook'
import { resolveWorkbook as resolveDefinition } from '../src/grid/workbook-layout'
import { resolveWorkbook, renderWorkbookReport } from '../src/xlsx/workbook-template'
import { WorkbookTemplate } from '../src/xlsx/template'
import { importAuthoredWorkbook, openWorkbook } from './xlsx'

const definition = WorkbookTemplate.content(project.definition).definition
const plain = (config: WorkbookDefinition, data: unknown) => resolveDefinition(config, data).sheets[0].cells.map(cell => ({ at: cell.at, size: cell.size, value: cell.value }))

describe('nested row placement', () => {
  it('keeps local ownership, absolute projections and every iteration of the data origin', () => {
    const regions = workbookRegions(definition.sheets[0])
    expect(regions.map(region => [region.source.path, region.row, region.depth])).toEqual([['sites', 4, 0], ['work', 6, 1]])
    const code = workbookCells(definition.sheets[0]).find(cell => 'path' in cell.value && cell.value.path === 'code')!
    expect(code.at).toEqual({ row: 6, column: 1 })
    expect(workbookRows(definition.sheets[0])).toContainEqual({ index: 6, height: 26 })
    const result = resolveWorkbook(project.definition, project.data).sheets[0]
    expect(result.cells.filter(cell => /^\$data.sites\[\d+\].name$/.test(cell.origin.dataPath)).map(cell => cell.at.row)).toEqual([4, 9, 16])
    expect(result.cells.filter(cell => /^\$data.sites\[\d+\].note$/.test(cell.origin.dataPath)).map(cell => cell.at.row)).toEqual([6, 13, 24])
    expect(result.cells.find(cell => 'literal' in cell.value && cell.value.literal === 'Prepared by')!.at.row).toBe(28)
    expect(result.cells.find(cell => cell.definitionId === code.id)!.origin).toEqual({ nodeId: code.id, dataPath: '$data.sites[1].work[0].code', iterations: [{ nodeId: regions[0].id, index: 1 }, { nodeId: regions[1].id, index: 0 }] })
    expect(new Set(result.cells.map(cell => cell.id)).size).toBe(result.cells.length)
  })

  it.each([0, 1, 3])('renders %i sites with unequal nested repeats', async count => {
    const data = { ...project.data, sites: project.data.sites.slice(0, count) }
    const output = (await openWorkbook(await renderWorkbookReport(project.definition, data))).worksheets[0]
    const end = ({ 0: 5, 1: 10, 3: 28 })[count as 0 | 1 | 3]
    expect(output.getCell('A' + end).value).toBe('Prepared by')
    if (count > 1) {
      expect(output.getCell('A11').value).toBe('0007')
      expect(output.getCell('D11').value).toBe(0)
      expect(output.getCell('D12').value).toBe(false)
    }
  })

  it('retains nested definitions when flattening object scopes', () => {
    expect(expandWorkbookScopes(definition)).toEqual(definition)
    const config: WorkbookDefinition = { sheets: [{ id: 's', name: 'Scopes', cells: [], regions: [{
      id: 'outer', type: 'repeat', row: 1, height: 2, source: { path: 'items' }, cells: [], regions: [{
        id: 'object', type: 'scope', row: 1, height: 2, source: { path: 'detail' }, cells: [{ id: 'name', at: { row: 1, column: 1 }, size: { rows: 1, columns: 1 }, value: { path: 'name', optional: true } }], regions: [{
          id: 'inner', type: 'repeat', row: 2, height: 1, source: { path: 'work' }, cells: [{ id: 'code', at: { row: 1, column: 1 }, size: { rows: 1, columns: 1 }, value: { path: 'code' } }],
        }],
      }],
    }] }] }
    const expanded = expandWorkbookScopes(config)
    const data = { items: [{ detail: { work: [{ code: '0007' }] } }, { detail: { name: 'Second', work: [] } }] }
    expect(plain(expanded, data)).toEqual(plain(config, data))
    expect(workbookRegions(expanded.sheets[0]).map(region => region.source)).toEqual([{ path: 'items' }, { path: 'detail.work', from: 'current' }])
  })

  it('collapses coincident empty boundaries at three nested levels', async () => {
    const template = await importAuthoredWorkbook(book => {
      book.addWorksheet('Deep').addRows([
        ['{#a}'], ['{#.b}'], ['{#.c}'], ['{.value}'], ['{/.c}'], ['{/.b}'], ['{/a}'], ['End'],
      ])
    })
    const data = { a: [{ b: [] }, { b: [{ c: [] }, { c: [{ value: '0007' }, { value: false }] }] }] }
    const cells = resolveWorkbook(template, data).sheets[0].cells
    expect(cells.map(cell => [cell.at.row, cell.value])).toEqual([[1, { literal: '0007' }], [2, { literal: false }], [3, { literal: 'End' }]])
    expect(resolveWorkbook(template, { a: [] }).sheets[0].cells).toMatchObject([{ at: { row: 1 }, value: { literal: 'End' } }])
  })
})

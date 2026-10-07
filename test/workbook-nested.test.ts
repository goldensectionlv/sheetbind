import { describe, expect, it } from 'vitest'
import * as project from '../examples/regions/nested'
import { resolveWorkbook, renderWorkbookReport, renderWorkbookForm, readWorkbookForm } from '../src/index'
import { importAuthoredWorkbook, openWorkbook } from './xlsx'

describe('nested row placement', () => {
  it('keeps nested field positions, row settings and every iteration of the data origin', () => {
    const result = resolveWorkbook(project.definition, project.data).sheets[0]
    const site = result.cells.find(cell => cell.origin.dataPath === '$data.sites[1].name')!
    const code = result.cells.find(cell => cell.origin.dataPath === '$data.sites[1].work[0].code')!
    expect(result.cells.filter(cell => /^\$data.sites\[\d+\].name$/.test(cell.origin.dataPath)).map(cell => cell.at.row)).toEqual([4, 9, 16])
    expect(result.cells.filter(cell => /^\$data.sites\[\d+\].note$/.test(cell.origin.dataPath)).map(cell => cell.at.row)).toEqual([6, 13, 24])
    expect(result.cells.find(cell => 'literal' in cell.value && cell.value.literal === 'Prepared by')!.at.row).toBe(28)
    expect(code).toMatchObject({ at: { row: 11, column: 1 },
      origin: { nodeId: code.definitionId, iterations: [site.origin.iterations[0], { nodeId: expect.any(String), index: 0 }] },
    })
    expect(site.origin.iterations).toEqual([{ nodeId: expect.any(String), index: 1 }])
    expect(new Set(code.origin.iterations.map(item => item.nodeId)).size).toBe(2)
    expect(result.rows).toContainEqual({ index: 11, height: 26 })
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

  it('reads nested form records through object scopes and explicit root bindings', async () => {
    const template = await importAuthoredWorkbook(book => book.addWorksheet('Scopes').addRows([
      ['{#items}'], ['{#with .detail}'], ['{?.name}', '{$root.title}'],
      ['{#work}'], ['{.code}'], ['{/work}'], [null, '{/with}'], [null, '{/items}'],
    ]))
    const title = '_x000a_ = SUM(A1)\nText'
    const data = { title, items: [{ detail: { work: [{ code: '0007' }] } }, { detail: { name: 'Second', work: [] } }] }
    const issued = await renderWorkbookForm(template, data)
    expect(await readWorkbookForm(template, issued)).toEqual({ success: true, data: {
      title, items: [{ detail: { name: null, work: [{ code: '0007' }] } }, { detail: { name: 'Second', work: [] } }],
    } })
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

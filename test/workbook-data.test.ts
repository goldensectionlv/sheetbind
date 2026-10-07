import { expect, it, vi } from 'vitest'
import { resolveWorkbook, renderWorkbookReport } from '../src/index'
import type { WorkbookTemplate } from '../src/index'
import { importAuthoredWorkbook, openWorkbook } from './xlsx'

const scalarTemplate = await importAuthoredWorkbook(book => book.addWorksheet('Values').getCell('A1').value = '{input}')
const repeatTemplate = await importAuthoredWorkbook(book => book.addWorksheet('Values').addRows([
  ['Header'], ['{#items}'], ['{.name}'], ['{/items}'], ['End'],
]))
function values(template: WorkbookTemplate, data: unknown) {
  return resolveWorkbook(template, data).sheets[0].cells.map(cell => 'literal' in cell.value ? cell.value.literal : cell.value)
}

it.each(['00123', '=SUM(A1:A2) {value}', '_x0041_', '', 0, false, null])('preserves scalar %j through XLSX', async input => {
  expect(values(scalarTemplate, { input })).toEqual([input])
  const book = await openWorkbook(await renderWorkbookReport(scalarTemplate, { input }))
  expect(book.worksheets[0].getCell('A1').value).toBe(input)
})

it.each([{}, [], NaN, Infinity, new Date()])('rejects undeclared structured or nonfinite values %#', input => {
  expect(() => values(scalarTemplate, { input })).toThrow()
})

it('does not fall back to a parent when the current field is absent', async () => {
  const warning = vi.spyOn(console, 'warn').mockImplementation(() => {})
  try {
    expect(values(repeatTemplate, { name: 'Wrong', items: [{}] })).toEqual(['Header', null, 'End'])
    expect(warning).toHaveBeenCalledExactlyOnceWith(expect.stringContaining('$data.items[0].name'))
  }
  finally {
    warning.mockRestore()
  }
  const optional = await importAuthoredWorkbook(book => book.addWorksheet('Values').addRows([
    ['{#items}'], ['{?.name}'], ['{/items}'],
  ]))
  expect(values(optional, { name: 'Wrong', items: [{}] })).toEqual([null])
})

it.each([{}, { details: null }])('renders blank fields for absent object scopes %#', async data => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Scope').addRows([
    ['{#with details}'], ['{.name}'], ['{/with}'],
  ]))
  expect(values(template, data)).toEqual([null])
})

it('composes scopes, nested repeats and explicit root access with concrete origins', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Nested').addRows([
    ['{#items}'], ['{$root.title}'], ['{#with .details}'], ['{.code}'], ['{/with}'],
    ['{#notes}'], ['{.text}'], ['{/notes}'], ['{/items}'],
  ]))
  const data = { title: 'Overview', items: [{ details: { code: '001' }, notes: [{ text: 'Ready' }, { text: 'Next' }] }] }
  expect(values(template, data)).toEqual(['Overview', '001', 'Ready', 'Next'])
  const cell = resolveWorkbook(template, data).sheets[0].cells.at(-1)!
  expect(cell.origin).toMatchObject({ nodeId: cell.definitionId, dataPath: '$data.items[0].notes[1].text', iterations: [{ index: 0 }, { index: 1 }] })
  expect(cell.contextPath).toBe('$data.items[0].notes[1]')
})

it('validates bindings inside empty repeats when importing the template', async () => {
  await expect(importAuthoredWorkbook(book => book.addWorksheet('Unsafe').addRows([
    ['{#items}'], ['{.constructor.name}'], ['{/items}'],
  ]))).rejects.toThrow()
})

it('supports null-prototype objects without accepting inherited values', () => {
  expect(values(scalarTemplate, Object.assign(Object.create(null), { input: 'Own' }))).toEqual(['Own'])
  expect(() => values(scalarTemplate, Object.create({ input: 'Inherited' }))).toThrow()
})

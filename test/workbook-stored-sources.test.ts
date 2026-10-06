import { expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import { readFile } from 'node:fs/promises'
import { importAuthoredWorkbook, saveWorkbook } from './xlsx'
import { importWorkbookXlsx, readWorkbookForm, renderWorkbookForm, renderWorkbookReport } from '../src/index'

const sourceTag = '{.answer}{@choice:.answers; key=id; label=label; return=key; emptySource=input}'

it.each(['x'.repeat(32768), 'A\u0001B', '\ud800'])('rejects invalid text in an unselected local choice %#', async label => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = '{answer}{@choice:.options; key=id; label=label; return=key}')
  for (const render of [renderWorkbookForm, renderWorkbookReport]) {
    await expect(render(template, { answer: null, options: [{ id: 'a', label }] })).rejects.toThrow(/32767|XML/)
  }
})
async function load(bytes: Uint8Array) {
  const book = new ExcelJS.Workbook()
  await book.xlsx.load(Uint8Array.from(bytes).buffer)
  return book
}

it('reads named key, object, root and string-list sources without a second input or dictionary', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Form').addRows([
    ['{key}{@choice:Options; key=id; label=label; return=key}'],
    ['{object}{@choice:Options; key=id; label=label}'],
    ['{rootKey}{@choice:$root.catalog; key=id; label=label; return=key}'],
    ['{status}{@list:Statuses}'],
  ]))
  const object = { id: '006', label: 'Original', nested: { rate: 0, enabled: false } }
  const dictionaries = { Options: [object], Statuses: ['Open', 'Closed'] }
  const issued = await renderWorkbookForm(template, { key: '006', object, rootKey: 0, catalog: [{ id: 0, label: 'Zero' }], status: 'Open' }, { dictionaries })
  object.label = 'Changed later'
  dictionaries.Statuses.splice(0)
  expect(await readWorkbookForm(template, issued)).toEqual({ success: true, data: {
    key: '006', object: { ...object, label: 'Original' }, rootKey: 0, status: 'Open',
  } })
  const book = await load(issued)
  book.worksheets[0].getCell('A4').value = 'Unknown'
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [{ code: 'list' }] })
  const range = book.definedNames.getRanges('_sb_object_sources').ranges[0]
  book.getWorksheet('_sheetbind_lists')!.getCell(range.split('!')[1].split(':')[0].replaceAll('$', '')).value = 'broken'
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [{ code: 'choice-source' }] })
})

it('keeps distinct local sources after row sorting, copying, deletion and a fresh template import', async () => {
  const authored = new ExcelJS.Workbook()
  authored.addWorksheet('Survey').addRows([
    ['{#questions}'], ['{.question}', sourceTag], [null, '{/questions}'],
  ])
  const bytes = await saveWorkbook(authored)
  const issued = await renderWorkbookForm(await importWorkbookXlsx(bytes), { questions: [
    { question: 'A', answer: 'same', answers: [{ id: 'same', label: 'First' }] },
    { question: 'B', answer: 'same', answers: [{ id: 'same', label: 'Second' }] },
    { question: 'C', answer: 0, answers: [] },
  ] })
  const book = await load(issued)
  const sheet = book.worksheets[0]
  const first = sheet.getRow(2).values
  sheet.getRow(2).values = sheet.getRow(3).values
  sheet.getRow(3).values = first
  sheet.spliceRows(4, 0, sheet.getRow(2).values as ExcelJS.CellValue[])
  sheet.spliceRows(3, 1)
  const template = await importWorkbookXlsx(bytes)
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toEqual({ success: true, data: { questions: [
    { question: 'B', answer: 'same' }, { question: 'B', answer: 'same' }, { question: 'C', answer: 0 },
  ] } })
  sheet.getCell('B2').value = 'First'
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [{ code: 'choice', address: 'B2' }] })
  sheet.getCell('B2').value = 'Second'
  sheet.getCell(2, sheet.columnCount).value = null
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [{ code: 'choice-source', address: 'B2' }] })
})

it('keeps sources for two independent local fields in nested records, including object payloads', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Survey').addRows([
    ['{#groups}'], ['{.name}'], ['{#.questions}'],
    [sourceTag, '{.detail}{@choice:.details; key=id; label=label}'],
    [null, '{/.questions}'], [null, '{/groups}'],
  ]))
  const detail = { id: 'x', label: 'Detail', rate: 0, code: '006' }
  const data = { groups: [{ name: 'Group', questions: [{ answer: 'a', answers: [{ id: 'a', label: 'Answer' }], detail, details: [detail] }] }] }
  expect(await readWorkbookForm(template, await renderWorkbookForm(template, data))).toEqual({ success: true, data: {
    groups: [{ name: 'Group', questions: [{ answer: 'a', detail }] }],
  } })
})

it('issues twenty blank inputs with row-local sources and keeps only completed rows', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Survey').addRows([
    ['{#questions}'], [sourceTag], ['{/questions}'],
  ]))
  const book = await load(await renderWorkbookForm(template, { questions: Array.from({ length: 20 }, () => ({ answer: null, answers: [{ id: 'a', label: 'Answer' }] })) }))
  book.worksheets[0].getCell('A2').value = 'Answer'
  book.worksheets[0].getCell('A21').value = 'Answer'
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toEqual({ success: true, data: { questions: [{ answer: 'a' }, { answer: 'a' }] } })
})

it.each([{}, { options: undefined }, { options: null }, { options: [] }])('rejects input for an absent or empty local source unless free input was explicitly enabled %#', async empty => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Survey').addRows([
    ['{#items}'], ['{.answer}{@choice:.options; key=id; label=label; return=key}'], ['{/items}'],
  ]))
  const book = await load(await renderWorkbookForm(template, { items: [empty, { options: [{ id: 'a', label: 'A' }], answer: 'a' }] }))
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toEqual({ success: true, data: { items: [{ answer: 'a' }] } })
  book.worksheets[0].getCell('A2').value = 'A'
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [{ code: 'choice', address: 'A2' }] })
})

it('accepts omitted, undefined and null local sources as empty without weakening populated choices or required fields', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Survey').addRows([
    ['{#questions}'], ['{.question}', sourceTag + '{@validate:required}'], [null, '{/questions}'],
  ]))
  const data = { questions: [
    { question: 'Omitted' }, { question: 'Undefined', answers: undefined },
    { question: 'Null', answers: null }, { question: 'Empty', answers: [] },
    { question: 'Choice', answers: [{ id: 'a', label: 'A' }] },
  ] }
  const before = structuredClone(data)
  const book = await load(await renderWorkbookForm(template, data))
  const required = await readWorkbookForm(template, await saveWorkbook(book))
  expect(required).toMatchObject({ success: false })
  if (!required.success) {
    expect(required.issues.map(issue => issue.rule)).toEqual(Array(5).fill('required'))
  }
  const values = ['Free text', '006', 0, 'Empty list', 'A']
  values.forEach((value, index) => {
    book.worksheets[0].getCell(index + 2, 2).value = value
  })
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toEqual({ success: true, data: { questions:
    data.questions.map((question, index) => ({ question: question.question, answer: index === 4 ? 'a' : values[index] })),
  } })
  expect(data).toStrictEqual(before)
  book.worksheets[0].getCell('B6').value = 'Free text'
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [{ code: 'choice', address: 'B6' }] })
})

it.each([{}, { catalog: undefined }, { catalog: null }, { catalog: {} }, { catalog: { options: null } }])('treats absent nested root sources as empty in reports and self-contained forms %#', async source => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = '{answer}{@choice:$root.catalog.options; key=id; label=label; return=key; emptySource=input}')
  const data = { ...source, answer: 'Free text' }
  const before = structuredClone(data)
  expect((await load(await renderWorkbookReport(template, data))).worksheets[0].getCell('A1').value).toBe('Free text')
  expect(await readWorkbookForm(template, await renderWorkbookForm(template, data))).toEqual({ success: true, data: { answer: 'Free text' } })
  expect(data).toStrictEqual(before)
})

it('still rejects invalid source types, missing named dictionaries and lost local metadata', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = sourceTag)
  for (const answers of ['', 0, false, {}]) {
    await expect(renderWorkbookForm(template, { answers })).rejects.toMatchObject({ issues: [{ code: 'choice-source' }] })
  }
  const named = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = '{answer}{@choice:Options; key=id; label=label; return=key; emptySource=input}')
  await expect(renderWorkbookForm(named, {})).rejects.toMatchObject({ issues: [{ code: 'missing-dictionary' }] })
  const book = await load(await renderWorkbookForm(template, { answer: 'Free text' }))
  const range = book.definedNames.getRanges('_sb_object_sources').ranges[0]
  const address = range.split('!')[1].split(':')[1].replaceAll('$', '')
  const cell = book.getWorksheet('_sheetbind_lists')!.getCell(address)
  const payload = JSON.parse(cell.value as string)
  payload.local = {}
  cell.value = JSON.stringify(payload)
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [{ code: 'choice-source' }] })
})

it('reads an Excel-saved form after native sorting and copying, with independent local dropdowns', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Survey').addRows([
    ['{status}{@list:Statuses}'], ['{#questions}'],
    ['{.question}', sourceTag, '{.detail}{@choice:.details; key=id; label=label}'],
    [null, null, '{/questions}'],
  ]))
  const bytes = await readFile(new URL('./fixtures/local-choices-excel.xlsx', import.meta.url))
  expect(await readWorkbookForm(template, bytes)).toEqual({ success: true, data: { status: 'Open', questions: [
    { question: 'Text', answer: 'A free answer', detail: { id: 'same', label: 'Text detail' } },
    { question: 'Beta', answer: 'same', detail: { id: 'same', label: 'Beta detail', code: '007', rate: 20 } },
    { question: 'Beta', answer: 'other', detail: { id: 'same', label: 'Beta detail', code: '007', rate: 20 } },
    { question: 'Alpha', answer: 'same', detail: { id: 'same', label: 'Alpha detail', code: '006', rate: 0 } },
  ] } })
})

import { expect, it, vi } from 'vitest'
import ExcelJS from 'exceljs'
import { importWorkbookXlsx, readWorkbookForm, renderWorkbookForm, renderWorkbookReport, resolveWorkbook } from '../src/index'
import { openWorkbook as load, importAuthoredWorkbook, saveWorkbook } from './xlsx'

const sourceTag = '{.answer}{@choice:.answers; key=id; label=label; return=key; emptySource=input}'

it.each(['Options', '$root.options'])('validates shared choice source %s before issuing an empty multirow repeat', async source => {
  const book = new ExcelJS.Workbook()
  book.addWorksheet('Form').addRows([
    ['{#items}'], ['Name', '{.name}'], ['Selected', `{.selected}{@choice:${source}; key=id; label=name; return=key}`], [null, '{/items}'],
  ])
  const authored = await saveWorkbook(book)
  const template = await importWorkbookXlsx(authored)
  for (const values of [[{ id: 'a' }], ['Allowed'], [{ id: 'a', name: 'A' }, { id: 'a', name: 'B' }]]) {
    for (const render of [resolveWorkbook, renderWorkbookReport, renderWorkbookForm]) {
      await expect(Promise.resolve().then<unknown>(() => render(template, { items: [], options: values }, { dictionaries: { Options: values } }))).rejects.toMatchObject({
        issues: [{ phase: 'data', code: 'choice-source', path: source === 'Options' ? '$dictionaries.Options' : '$data.options', sheetName: 'Form', address: 'B3' }],
      })
    }
  }
  for (const values of [[], [{ id: 'a', name: 'Allowed' }]]) {
    const bytes = await renderWorkbookForm(template, { items: [], options: values }, { dictionaries: { Options: values } })
    expect(await readWorkbookForm(await importWorkbookXlsx(authored), bytes)).toEqual({ success: true, data: { items: [] } })
  }
})

it.each(['x'.repeat(32768), 'A\u0001B', '\ud800'])('rejects invalid text in an unselected local choice %#', async label => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = '{answer}{@choice:.options; key=id; label=label; return=key}')
  for (const render of [renderWorkbookForm, renderWorkbookReport]) {
    await expect(render(template, { answer: null, options: [{ id: 'a', label }] })).rejects.toThrow(/32767|XML/)
  }
})

it('captures named, root, local and string-list sources before asynchronous form writing', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Form').addRows([
    ['{key}{@choice:Options; key=id; label=label; return=key}'],
    ['{object}{@choice:Options; key=id; label=label}'],
    ['{rootKey}{@choice:$root.catalog; key=id; label=label; return=key}'],
    ['{status}{@list:Statuses}'],
    ['{local}{@choice:.catalog; key=id; label=label}'],
  ]))
  const object = { id: '006', label: 'Original', nested: { rate: 0, enabled: false } }
  const root = { id: 0, label: 'Zero', nested: { retained: true } }
  const dictionaries = { Options: [object], Statuses: ['Open', 'Closed'] }
  const data = { key: '006', object, rootKey: 0, status: 'Open', local: root }
  const expected = structuredClone(data)
  const pending = renderWorkbookForm(template, { ...data, catalog: [root] }, { dictionaries })
  object.label = 'Changed later'
  object.nested.rate = 99
  root.label = 'Changed while writing'
  root.nested.retained = false
  dictionaries.Statuses.splice(0)
  const issued = await pending
  expect(await readWorkbookForm(template, issued)).toEqual({ success: true, data: expected })
  const book = await load(issued)
  book.worksheets[0].getCell('A4').value = 'Unknown'
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [{ code: 'list' }] })
  const range = book.definedNames.getRanges('_sb_object_sources').ranges[0]
  book.getWorksheet('_sheetbind_lists')!.getCell(range.split('!')[1].split(':')[0].replaceAll('$', '')).value = 'broken'
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [{ code: 'choice-source' }] })
})

it('keeps distinct local sources for issued fields without movable worksheet references', async () => {
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
  expect(sheet.getCell('B2').dataValidation.formulae[0]).toMatch(/^_sb_list_/)
  expect(sheet.getCell('B3').dataValidation.formulae[0]).not.toEqual(sheet.getCell('B2').dataValidation.formulae[0])
  expect(sheet.getCell('B4').dataValidation).toBeUndefined()
  expect(sheet.columnCount).toBe(257)
  const template = await importWorkbookXlsx(bytes)
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toEqual({ success: true, data: { questions: [
    { question: 'A', answer: 'same' }, { question: 'B', answer: 'same' }, { question: 'C', answer: 0 },
  ] } })
  sheet.getCell('B2').value = 'Second'
  expect(await readWorkbookForm(template, await saveWorkbook(book))).toMatchObject({ success: false, issues: [{ code: 'choice', address: 'B2' }] })
  sheet.getCell('B2').value = 'First'
  expect((await readWorkbookForm(template, await saveWorkbook(book))).success).toBe(true)
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

it('still rejects invalid source types', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = sourceTag)
  for (const answers of ['', 0, false, {}]) {
    await expect(renderWorkbookForm(template, { answers })).rejects.toMatchObject({ issues: [{ code: 'choice-source' }] })
  }
})

it.each(['named', 'root', 'local', 'list'])('distinguishes absent dictionaries from damaged contextual sources (%s)', async kind => {
  const rule = kind === 'list' ? '{@list:Options}' : `{@choice:${kind === 'named' ? 'Options' : kind === 'root' ? '$root.options' : '.options'}; key=id; label=name; return=key}`
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = '{selected}' + rule)
  const options = [{ id: 'a', name: 'Allowed' }]
  for (const input of [null, 'Allowed', { error: '#VALUE!' }] as const) {
    const book = await load(await renderWorkbookForm(template, { selected: null, options }, { dictionaries: { Options: kind === 'list' ? ['Allowed'] : options } }))
    book.worksheets[0].getCell('A1').value = input
    const range = book.definedNames.getRanges('_sb_object_sources').ranges[0]
    const cell = book.getWorksheet('_sheetbind_lists')!.getCell(range.split('!')[1].split(':')[1].replaceAll('$', ''))
    const payload = JSON.parse(String(cell.value))
    if (kind === 'named' || kind === 'list') {
      delete payload.dictionaries.Options
    }
    else if (kind === 'root') {
      delete payload.context.options
    }
    else {
      payload.local = {}
    }
    cell.value = JSON.stringify(payload)
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => {})
    try {
      const result = await readWorkbookForm(template, await saveWorkbook(book))
      if (kind === 'named' || kind === 'list') {
        expect(warning).toHaveBeenCalledExactlyOnceWith(expect.stringContaining('dictionary Options was not supplied'))
        expect(result).toEqual(typeof input === 'object' && input !== null
          ? { success: false, data: { selected: null }, issues: [expect.objectContaining({ code: 'non-scalar' })] }
          : { success: true, data: { selected: input } })
      }
      else {
        expect(result).toMatchObject({ success: false, issues: [{ phase: 'xlsx', code: 'choice-source', sheetName: 'Input', address: 'A1' }] })
        expect(result).not.toHaveProperty('data')
      }
    }
    finally {
      warning.mockRestore()
    }
  }
})

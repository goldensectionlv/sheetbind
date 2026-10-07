import { describe, expect, it, vi } from 'vitest'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import * as project from '../examples/fields/template'
import { parseDictionaries } from '../src/core/dictionaries'
import { validateList } from '../src/core/field-rules'
import { resolveWorkbook, workbookDictionarySources, renderWorkbookReport, importWorkbookXlsx, renderWorkbookForm, readWorkbookForm, workbookChoiceRange } from '../src/index'
import { openWorkbook as load, exampleFile, importAuthoredWorkbook, saveWorkbook } from './xlsx'

const options = { dictionaries: project.dictionaries }

describe('named string dictionaries', () => {
  it('keeps dependencies on fields and dictionary values outside the template', async () => {
    expect(workbookDictionarySources(project.definition)).toEqual(['workCodes'])

    const bytes = await exampleFile('fields/template.xlsx')
    const book = await load(bytes)
    expect(book.worksheets).toHaveLength(1)
    expect(book.definedNames.model).toEqual([])
    expect(book.worksheets[0].getCell('A8').value).toBe('{.code}\n{@validate:required|string}\n{@list:workCodes}')
    const zip = await JSZip.loadAsync(bytes)
    expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).not.toContain('<dataValidations')
    expect(await zip.file('xl/sharedStrings.xml')!.async('string')).not.toContain('TASK-300')
    const imported = await importWorkbookXlsx(bytes)
    expect(workbookDictionarySources(imported)).toEqual(['workCodes'])
  })

  it('warns about unavailable lists without validating report values or empty repeats', () => {
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => {})
    try {
      expect(() => resolveWorkbook(project.definition, { ...project.data, sites: [] })).not.toThrow()
      expect(warning).toHaveBeenCalledExactlyOnceWith(expect.stringContaining('dictionary workCodes is missing or empty'))
      expect(() => resolveWorkbook(project.definition, { ...project.data, sites: [] }, { dictionaries: { workCodes: [] } })).not.toThrow()
    }
    finally {
      warning.mockRestore()
    }
    expect(() => resolveWorkbook(project.definition, { ...project.data, sites: [] }, options)).not.toThrow()
    const data = structuredClone(project.data) as { sites: { work: { code: string }[] }[] }
    data.sites[1].work[0].code = '7'
    expect(() => resolveWorkbook(project.definition, data, options)).not.toThrow()
    expect(validateList('00042', { validation: [{ rule: 'string' }], list: 'codes' }, ['00042'])).toBeUndefined()
    expect(validateList(42, { validation: [{ rule: 'string' }], list: 'codes' }, ['00042'])?.code).toBe('list')
  })

  it('warns once per missing dictionary and preserves ordinary input through saved reports and forms', async () => {
    const book = new ExcelJS.Workbook()
    const sheet = book.addWorksheet('Input')
    sheet.addRows([
      ['{code}{@list:Codes}{@validate:required}', '{product}{@choice:Products; key=id; label=name; return=key}', '{details}{@choice:Products; key=id; label=name}', '{state}{@list:States}'],
      ['{second}{@list:Codes}'],
    ])
    const range = workbookChoiceRange({ source: { dictionary: 'Products' }, key: 'id', label: 'name' })
    sheet.getCell('E1').value = { formula: `COUNTA(${range})` }
    const authored = await saveWorkbook(book)
    const template = await importWorkbookXlsx(authored)
    const data = { code: '006', second: 'Free text', product: '007', details: { id: '008', name: 'Label from input' }, state: 'Open' }
    const dictionaries = { States: ['Open'] }
    const before = structuredClone({ data, dictionaries })
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => {})
    try {
      for (const render of [renderWorkbookReport, renderWorkbookForm]) {
        warning.mockClear()
        const rendered = await load(await render(template, data, { dictionaries }))
        expect(warning.mock.calls.map(([message]) => message)).toEqual([
          expect.stringContaining('dictionary Codes is missing or empty'), expect.stringContaining('dictionary Products is missing or empty'),
        ])
        const input = rendered.getWorksheet('Input')!
        expect(input.getRow(1).values).toEqual([undefined, '006', '007', 'Label from input', 'Open', { formula: `COUNTA(${range})` }])
        for (const address of ['A1', 'B1', 'C1', 'A2']) {
          expect(input.getCell(address).dataValidation).toBeUndefined()
        }
        expect(input.getCell('D1').dataValidation.type).toBe('list')
        expect(rendered.definedNames.getRanges(range).ranges).toHaveLength(1)
        if (render === renderWorkbookForm) {
          warning.mockClear()
          const fresh = await importWorkbookXlsx(authored)
          expect(await readWorkbookForm(fresh, await saveWorkbook(rendered))).toEqual({ success: true, data: { ...data, details: 'Label from input' } })
          expect(warning).toHaveBeenCalledTimes(2)
          input.getCell('A1').value = null
          input.getCell('D1').value = 'Unknown'
          expect(await readWorkbookForm(fresh, await saveWorkbook(rendered))).toMatchObject({ success: false, issues: [
            { rule: 'required', address: 'A1' }, { code: 'list', address: 'D1' },
          ] })
        }
      }
      expect({ data, dictionaries }).toStrictEqual(before)
    }
    finally {
      warning.mockRestore()
    }
  })

  it('writes one range per source after nested placement from tagged workbooks', async () => {
    const bytes = await renderWorkbookReport(project.definition, project.data, options)

    const book = await load(bytes)
    const sheet = book.worksheets[0]
    const lists = book.worksheets[1]
    expect(book.worksheets).toHaveLength(2)
    expect(lists.state).toBe('veryHidden')
    expect(book.definedNames.model).toEqual([{ name: '_sb_list_1', ranges: ["'_sheetbind_lists'!$A$1:$A$307"] }])
    for (const address of ['A11', 'A18', 'A20', 'A22']) {
      expect(sheet.getCell(address).dataValidation).toMatchObject({ type: 'list', formulae: ['_sb_list_1'], showErrorMessage: true, errorStyle: 'stop' })
    }
    expect(sheet.getCell('A11').value).toBe('0007')
    expect(sheet.getCell('B11').dataValidation).toBeUndefined()
    expect(lists.getCell('A5').value).toBe('00042')
    expect(lists.getCell('A6').value).toBe('Review, follow-up')
    expect(lists.getCell('A7').value).toBe('Special "quoted" work')
    expect(lists.getCell('A5').numFmt).toBe('@')
    expect(lists.getCell('A307').value).toBe('TASK-300')
    expect((await load(await renderWorkbookReport(project.definition, { ...project.data, sites: [] }, options))).worksheets).toHaveLength(1)
  })

  it('merges native prompts by their exact ranges when rendering reports and forms', async () => {
    const template = await importAuthoredWorkbook(book => {
      const sheet = book.addWorksheet('Prompts')
      sheet.addRows([
        ['{#items}'],
        ['{.code}{@choice:Codes; key=id; label=name; return=key}{@validate:required}', '{.other}{@list:Other}'],
        [null, '{/items}'], ['{footer}{@list:Other}'],
      ])
      const validation = Reflect.get(sheet, 'dataValidations')
      validation.add('A2:C2', { type: 'any', formulae: [], showInputMessage: true, promptTitle: 'Choose code', prompt: 'Agreed workflow', errorTitle: 'Authored error' })
      validation.add('A4', { type: 'any', formulae: [], showInputMessage: true, prompt: 'Footer instruction' })
    })
    const dictionaries = { Codes: [{ id: '001', name: 'Paper' }], Other: ['Open'] }
    for (const render of [renderWorkbookReport, renderWorkbookForm]) {
      for (const count of [0, 1, 3]) {
        const data = { items: Array.from({ length: count }, () => ({ code: '001', other: 'Open' })), footer: 'Open' }
        const saved = await load(await render(template, data, { dictionaries }))
        const sheet = saved.worksheets[0]
        const selected: ExcelJS.Cell[] = []
        sheet.eachRow(row => row.eachCell(cell => {
          if (cell.value === 'Paper') {
            selected.push(cell)
          }
        }))
        expect(selected).toHaveLength(count)
        for (const cell of selected) {
          expect(cell.dataValidation).toMatchObject({ type: 'list', promptTitle: 'Choose code', prompt: 'Agreed workflow', showInputMessage: true, errorTitle: 'Authored error' })
          expect(cell.dataValidation.allowBlank).toBeFalsy()
          expect(sheet.getCell(Number(cell.row), 2).dataValidation).toMatchObject({ type: 'list', prompt: 'Agreed workflow' })
        }
        const prompts: string[] = []
        for (const rule of Object.values(Reflect.get(sheet, 'dataValidations').model) as ExcelJS.DataValidation[]) {
          if (rule.prompt) {
            prompts.push(rule.prompt)
          }
        }
        expect(prompts).toContain('Footer instruction')
        // The prompt outside the repeated rectangle remains on its first row only.
        const outside = Object.entries(Reflect.get(sheet, 'dataValidations').model).filter(([address]) => address.startsWith('C'))
        expect(outside).toHaveLength(1)
        expect(outside[0][1]).toMatchObject({ prompt: 'Agreed workflow' })
      }
    }
  })

  it.each(['list', 'custom', 'whole'] as const)('preserves native %s validation and skips only the overlapping generated dropdown', async type => {
    const template = await importAuthoredWorkbook(book => {
      const sheet = book.addWorksheet('Conflict')
      sheet.addRows([['{name}{@list:Names}', '{second}{@list:Names}'], ['Authored only']])
      Reflect.get(sheet, 'dataValidations').add('A1:A2', { type, formulae: [type === 'list' ? '"Other"' : '1'] })
    })
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => {})
    try {
      for (const render of [renderWorkbookReport, renderWorkbookForm]) {
        warning.mockClear()
        const bytes = await render(template, { name: 'Alex', second: 'Alex' }, { dictionaries: { Names: ['Alex'] } })
        const output = (await load(bytes)).worksheets[0]
        for (const address of ['A1', 'A2']) {
          expect(output.getCell(address).dataValidation).toMatchObject({ type, formulae: [type === 'list' ? '"Other"' : type === 'custom' ? '1' : 1] })
        }
        expect(output.getCell('B1').dataValidation).toMatchObject({ type: 'list', formulae: [expect.stringMatching(/^_sb_list_/)] })
        expect(warning).toHaveBeenCalledExactlyOnceWith(expect.stringContaining('existing validation at Conflict!A1'))
      }
    }
    finally {
      warning.mockRestore()
    }
  })

  it('owns declared sources, skips unavailable sources and ignores unused input', () => {
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => {})
    try {
      for (const value of [null, [], { codes: [42] }, { codes: [''] }, { codes: new Array(1) }]) {
        expect(parseDictionaries(value, ['codes'])).toEqual({})
      }
      expect(parseDictionaries({ codes: ['a', 'a'] }, ['codes'])).toEqual({ codes: ['a'] })
      expect(warning).toHaveBeenCalledWith(expect.stringContaining('duplicate strings'))
      warning.mockClear()
      const unused = { get unused() {
        throw new Error('must not be inspected')
      }, bad: new Date(), constructor: ['a'] }
      expect(parseDictionaries(unused, [])).toEqual({})
      expect(warning).not.toHaveBeenCalled()
    }
    finally {
      warning.mockRestore()
    }
    expect(parseDictionaries({ codes: Array.from({ length: 10001 }, (_, i) => String(i)) }, ['codes']).codes).toHaveLength(10001)
    const original = { codes: ['00042', 'a,b', '"x"'] }
    const parsed = parseDictionaries(original, ['codes'])
    original.codes.push('after')
    expect(parsed.codes).toHaveLength(3)

  })
})

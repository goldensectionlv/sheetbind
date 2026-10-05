import { describe, expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import * as project from '../examples/fields/template'
import { parseDictionaries } from '../src/core/dictionaries'
import { validateList } from '../src/core/field-rules'
import { TemplateError } from '../src/core/template'
import { resolveWorkbook, workbookDictionarySources, renderWorkbookReport, importWorkbookXlsx } from '../src/xlsx/workbook-template'
import { exampleFile } from './xlsx'
import { writeWorkbookDropdowns } from '../src/xlsx/workbook-dropdowns'
import { renderWorkbookForm } from '../src/xlsx/workbook-form'
import { importAuthoredWorkbook } from './xlsx'

const options = { dictionaries: project.dictionaries }
async function load(bytes: Buffer) {
  const book = new ExcelJS.Workbook()
  await book.xlsx.load(Uint8Array.from(bytes).buffer)
  return book
}

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

  it('validates dependencies before execution even when a repeat is empty', () => {
    for (const [dictionaries, code] of [[{}, 'missing-dictionary'], [{ workCodes: [] }, 'empty-dictionary']] as const) {
      try {
        resolveWorkbook(project.definition, { ...project.data, sites: [] }, { dictionaries })
        expect.fail('Expected a dependency error')
      }
      catch (error) {
        expect(error).toBeInstanceOf(TemplateError)
        expect((error as TemplateError).issues).toMatchObject([{ phase: 'data', code, path: '$dictionaries.workCodes' }])
      }
    }
    expect(() => resolveWorkbook(project.definition, { ...project.data, sites: [] }, options)).not.toThrow()
    const data = structuredClone(project.data) as { sites: { work: { code: string }[] }[] }
    data.sites[1].work[0].code = '7'
    expect(() => resolveWorkbook(project.definition, data, options)).toThrow('$data.sites[1].work[0].code: must be a value from workCodes')
    expect(validateList('00042', { validation: [{ rule: 'string' }], list: 'codes' }, ['00042'])).toBeUndefined()
    expect(validateList(42, { validation: [{ rule: 'string' }], list: 'codes' }, ['00042'])?.code).toBe('list')
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

  it('uses fresh names, reuses lists across sheets and preserves authored prompts', async () => {
    const book = new ExcelJS.Workbook()
    const sheet = book.addWorksheet('_SHEETBIND_LISTS')
    const second = book.addWorksheet('Second')
    sheet.getCell('A1').value = 'Foreign content'
    book.definedNames.add("'_SHEETBIND_LISTS'!$A$1", '_SB_LIST_1')
    const existing = { type: 'any', formulae: [], showInputMessage: true, promptTitle: 'Choose code', prompt: 'Provided by the application', errorTitle: 'Custom error' }
    Reflect.set(sheet.getCell('B2'), 'dataValidation', existing)
    const rules = { validation: [{ rule: 'string' }], list: 'codes' }
    writeWorkbookDropdowns(book, [{ sheet, address: 'B2', rules }, { sheet: second, address: 'C4', rules }, { sheet, address: 'D2', rules: { ...rules, list: 'other' } }], { codes: ['00042', 'a,b', 'say "yes"'], other: ['Other'] })
    const saved = await load(Buffer.from(await book.xlsx.writeBuffer()))
    const first = saved.worksheets[0]
    expect(saved.worksheets.map(sheet => sheet.name)).toEqual(['_SHEETBIND_LISTS', 'Second', '_sheetbind_lists_2'])
    expect(saved.definedNames.model.map(entry => entry.name)).toEqual(['_SB_LIST_1', '_sb_list_2', '_sb_list_3'])
    expect(first.getCell('A1').value).toBe('Foreign content')
    expect(first.getCell('B2').dataValidation).toMatchObject({ type: 'list', formulae: ['_sb_list_2'], allowBlank: true, showInputMessage: true, promptTitle: existing.promptTitle, prompt: existing.prompt, errorTitle: existing.errorTitle })
    expect(saved.worksheets[1].getCell('C4').dataValidation.formulae).toEqual(['_sb_list_2'])
    expect(existing.type).toBe('any')
    expect(existing.formulae).toEqual([])
    expect(saved.worksheets[2].getCell('B1').value).toBe('Other')
  })

  it('rejects a conflicting validation or unrepresentable text before changing the workbook', () => {
    const book = new ExcelJS.Workbook()
    const sheet = book.addWorksheet('Report')
    const rules = { validation: [{ rule: 'string' }], list: 'codes' }
    sheet.getCell('A2').dataValidation = { type: 'whole', operator: 'greaterThan', formulae: [0] }
    expect(() => writeWorkbookDropdowns(book, [{ sheet, address: 'A1', rules }, { sheet, address: 'A2', rules }], { codes: ['ok'] })).toThrow('Existing validation')
    expect(book.worksheets).toHaveLength(1)
    expect(sheet.getCell('A1').dataValidation).toBeUndefined()
    expect(() => writeWorkbookDropdowns(book, [{ sheet, address: 'A1', rules }], { codes: ['\u0001'] })).toThrow('XML cannot represent')
    expect(book.worksheets).toHaveLength(1)
    expect(book.definedNames.model).toEqual([])
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

  it.each(['list', 'custom', 'whole'] as const)('rejects native %s conflicts at the saved source boundary', async type => {
    const template = await importAuthoredWorkbook(book => {
      const sheet = book.addWorksheet('Conflict')
      sheet.getCell('A1').value = '{name}{@list:Names}'
      sheet.getCell('A1').dataValidation = { type, formulae: [type === 'list' ? '"Other"' : '1'] }
    })
    for (const render of [renderWorkbookReport, renderWorkbookForm]) {
      await expect(render(template, { name: 'Alex' }, { dictionaries: { Names: ['Alex'] } })).rejects.toThrow('Existing validation at Conflict!A')
    }
  })

  it('rejects invalid dictionary shapes without coercion or aliases', () => {
    for (const value of [null, [], { codes: [42] }, { codes: [''] }, { codes: new Array(1) }, { codes: ['a', 'a'] }, { 'bad|name': ['a'] }, JSON.parse('{"__proto__":["a"]}'), { constructor: ['a'] }]) {
      expect(() => parseDictionaries(value)).toThrow()
    }
    expect(parseDictionaries({ codes: Array.from({ length: 10001 }, (_, i) => String(i)) }).codes).toHaveLength(10001)
    const original = { codes: ['00042', 'a,b', '"x"'] }
    const parsed = parseDictionaries(original)
    original.codes.push('after')
    expect(parsed.codes).toHaveLength(3)

  })
})

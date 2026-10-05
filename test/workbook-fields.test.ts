import { describe, expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import { parseFieldRules } from '../src/core/field-rules'
import { createValidation } from '../src/core/validation'
import { TemplateError } from '../src/core/template'
import { resolveWorkbook, renderWorkbookReport, importWorkbookXlsx } from '../src/xlsx/workbook-template'
import { parseFieldTag } from '../src/xlsx/field-tag'

const book = new ExcelJS.Workbook()
const sheet = book.addWorksheet('Report')
sheet.addRows([
  [], ['{#groups}'], ['{.name}{@validate:required|string|maxLength:12}'], ['{#.items}'],
  ['{?.hours}{@validate:required|number|min:0|max:24}', '{.approved}{@validate:required|boolean}', '{.code}{@validate:string|maxLength:5}'],
  [null, null, '{/.items}'], [], [null, null, '{/groups}'],
])
sheet.mergeCells('A3:B3')
const project = {
  data: { groups: [{ name: 'Empty', items: [] }, { name: 'Main', items: [{ hours: 0, approved: false, code: '00042' }] }] },
  definition: await importWorkbookXlsx(Buffer.from(await book.xlsx.writeBuffer())),
}

describe('shared field rules', () => {
  it('keeps scalar types, zero and false through tagged XLSX inputs', async () => {
    const result = resolveWorkbook(project.definition, project.data)
    expect(result.sheets[0].cells.map(cell => 'literal' in cell.value ? cell.value.literal : undefined)).toEqual(['Empty', 'Main', 0, false, '00042'])
    const output = new ExcelJS.Workbook()
    await output.xlsx.load(Uint8Array.from(await renderWorkbookReport(project.definition, project.data)).buffer)
    expect(output.worksheets[0].getCell('A5').value).toBe(0)
    expect(output.worksheets[0].getCell('B5').value).toBe(false)
    expect(output.worksheets[0].getCell('C5').value).toBe('00042')
  })

  it('reports each invalid field with its definition identity and concrete nested data path', async () => {
    const data = { groups: [{ name: 'Main', items: [{ hours: -1, approved: 'false', code: 'too-long' }, { approved: false, code: '' }] }] }
    try {
      resolveWorkbook(project.definition, data)
      expect.fail('Expected field errors')
    }
    catch (error) {
      expect(error).toBeInstanceOf(TemplateError)
      expect((error as TemplateError).issues.map(({ code, path }) => ({ code, path }))).toEqual([
        { code: 'min', path: '$data.groups[0].items[0].hours' },
        { code: 'boolean', path: '$data.groups[0].items[0].approved' },
        { code: 'maxLength', path: '$data.groups[0].items[0].code' },
        { code: 'required', path: '$data.groups[0].items[1].hours' },
      ])
    }
    await expect(renderWorkbookReport(project.definition, data)).rejects.toThrow('$data.groups[0].items[0].hours')
  })

  it('uses binding and directive tags and rejects invalid choice options', async () => {
    expect(parseFieldTag('{?.hours}\n{@validate:required|number|min:0|max:24}')).toEqual({ value: { path: 'hours', from: 'current', optional: true }, rules: { validation: [{ rule: 'required' }, { rule: 'number' }, { rule: 'min', args: [0] }, { rule: 'max', args: [24] }] } })
    for (const options of ['key=id', 'label=name', 'key=__proto__.id; label=name', 'key=id; label=name; return=boolean']) {
      expect(() => parseFieldTag(`{x}{@choice:Products; ${options}}`)).toThrow()
    }
    const workbook = new ExcelJS.Workbook()
    workbook.addWorksheet('Report').getCell('A1').value = '{x | required}'
    await expect(importWorkbookXlsx(Buffer.from(await workbook.xlsx.writeBuffer()))).rejects.toThrow('Use a dotted path')

  })

  it('checks predicates without converting present zero, false or numeric text', () => {
    const prepare = createValidation()
    const context = { root: {}, current: {}, path: '$data.value' }
    expect(prepare('required|number|min:0')(0, context)).toBeUndefined()
    expect(prepare('required|boolean')(false, context)).toBeUndefined()
    expect(prepare('number')('', context)).toBeUndefined()
    expect(prepare('required')('   ', context)?.code).toBe('required')
    expect(prepare('number')('0', context)?.code).toBe('number')
    expect(prepare('number|max:24')(25, context)?.code).toBe('max')
    expect(prepare('string|maxLength:0')('a', context)?.code).toBe('maxLength')
    expect(() => parseFieldRules({ type: 'date' })).toThrow()
    expect(() => prepare('max:bad')).toThrow('Invalid arguments')
  })
})

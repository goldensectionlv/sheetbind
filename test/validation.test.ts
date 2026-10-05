import { expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import { definition, data, options } from '../examples/validation/definition'
import { createValidation, parseValidation, ValidationExecutionError } from '../src/core/validation'
import type { ValidationRule } from '../src/core/validation'
import { parseFieldTag } from '../src/xlsx/field-tag'
import { importAuthoredWorkbook, saveWorkbook } from './xlsx'
import { WorkbookTemplate } from '../src/xlsx/template'
import { resolveWorkbook, importWorkbookXlsx, renderWorkbookReport } from '../src/xlsx/workbook-template'
import { readWorkbookForm, renderWorkbookForm } from '../src/xlsx/workbook-form'

async function load(bytes: Buffer) {
  const book = new ExcelJS.Workbook()
  await book.xlsx.load(Uint8Array.from(bytes).buffer)
  return book
}
const context = { root: {}, current: {}, path: '$data.amount' }

it('carries ordered repeated rules and escaped JSON arguments/messages through field tags and XLSX', async () => {
  const pipeline = 'required|probe:"001",[1,true],{"text":"a|b:c,=;{d}"}|probe:false'
  const validation = parseValidation(pipeline)
  expect(validation[1].args).toEqual(['001', [1, true], { text: 'a|b:c,=;{d}' }])
  const rules = { validation: validation.map((use, index) => ({ ...use, ...(index === 1 ? { message: 'Ошибка "x" | : = ; {a}\nnext' } : {}) })), validationMessages: { probe: 'Fallback: "x"' } }
  const tag = '{amount}{@validate:required}{@validate:probe:"001",[1,true],{"text":"a|b:c,=;{d}"}; message=' + JSON.stringify(rules.validation[1].message) + '}{@validate:probe:false}{@validationMessage:probe:' + JSON.stringify(rules.validationMessages.probe) + '}'
  expect(parseFieldTag(tag).rules).toEqual(rules)
  const book = new ExcelJS.Workbook()
  book.addWorksheet('Data').getCell('A1').value = tag
  const restored = await importWorkbookXlsx(await saveWorkbook(book))
  expect(WorkbookTemplate.content(restored).definition.sheets[0].cells[0].rules).toEqual(rules)
  for (const invalid of ['{x}{y}', '{x}{@unknown:a}', '{x}{@validate:a|b; message="ambiguous"}', '{x}{@validationMessage:a:"unclosed}', '{x}{@choice:Items; key=id; key=other; label=name}']) {
    expect(() => parseFieldTag(invalid)).toThrow()
  }
})

it('uses occurrence, field, runtime and handler messages without aliases for repeated rules', () => {
  const calls: unknown[] = []
  const limit: ValidationRule = { validateArgs: args => args.length === 1, validate(value, args) {
    calls.push(args[0])
    return (value as number) <= (args[0] as number)
  }, message: ({ args }) => `default ${args[0]}` }
  const uses = [{ rule: 'limit', args: [20] }, { rule: 'limit', args: [5] }, { rule: 'limit', args: [1] }]
  const prepare = createValidation({ validationRules: { limit }, validationMessages: { limit: ({ args }) => `runtime ${args[0]}` } })
  expect(prepare(uses)(10, context)).toMatchObject({ rule: 'limit', args: [5], index: 1, message: 'runtime 5' })
  expect(calls).toEqual([20, 5])
  expect(prepare(uses, { limit: 'field' })(10, context)?.message).toBe('field')
  expect(prepare(uses.map(use => ({ ...use, message: '' })), { limit: 'field' })(10, context)?.message).toBe('')
  expect(createValidation({ validationRules: { limit } })(uses)(10, context)?.message).toBe('default 5')
})

it('prepares rules in empty repeats and blank forms, and keeps handler failures separate from value issues', async () => {
  expect(() => resolveWorkbook(definition, { ceiling: 10, items: [] })).toThrow('Unknown validation rule: decimalPlaces')
  await expect(renderWorkbookForm(definition, { ceiling: null, items: [] })).rejects.toThrow('Unknown validation rule')
  await expect(renderWorkbookForm(definition, { ceiling: null, items: [] }, options)).resolves.toBeInstanceOf(Buffer)
  expect(() => createValidation({ validationRules: { required: { validate: () => true } } })).toThrow('reserved')
  expect(() => createValidation(options)('decimalPlaces:bad')).toThrow('Invalid arguments')
  const config = await importAuthoredWorkbook(book => {
    book.addWorksheet('Data').getCell('A1').value = '{v}{@validate:broken}'
  })
  for (const validate of [() => {
    throw new RangeError('handler bug')
  }, () => Promise.resolve(true)]) {
    const runtime = { validationRules: { broken: { validate: validate as unknown as ValidationRule['validate'] } } }
    const form = await renderWorkbookForm(config, { v: 1 }, runtime)
    await expect(readWorkbookForm(config, form, runtime)).rejects.toBeInstanceOf(ValidationExecutionError)
    await expect(renderWorkbookReport(config, { v: 1 }, runtime)).rejects.toThrow('$data.v: validation rule broken')
  }
})

it('validates complete submitted records after sorting and blank-row removal, keeping actual cell addresses', async () => {
  const book = await load(await renderWorkbookForm(definition, data, options))
  const sheet = book.worksheets[0]
  let row = 0
  sheet.eachRow(current => {
    if (current.getCell(1).value === 'line-a') {
      row = current.number
    }
  })
  const first = sheet.getRow(row).values
  sheet.getRow(row).values = sheet.getRow(row + 1).values
  sheet.getRow(row + 1).values = first
  sheet.spliceRows(row, 0, [null, null, null])
  sheet.getCell(row + 1, 2).value = 12
  const seen: number[] = []
  const original = options.validationRules!.maxFromField
  const runtime = { ...options, validationRules: { ...options.validationRules, maxFromField: { ...original, validate(value: unknown, args: Parameters<ValidationRule['validate']>[1], context: Parameters<ValidationRule['validate']>[2]) {
    seen.push((context.root.items as unknown[]).length)
    return original.validate(value, args, context)
  } } } }
  const result = await readWorkbookForm(definition, Buffer.from(await book.xlsx.writeBuffer()), runtime)
  expect(result).toMatchObject({ success: false, issues: [{ rule: 'maxFromField', args: ['$root.ceiling'], index: 5, path: '$data.items[0].quantity', address: `B${row + 1}`, message: 'Quantity exceeds the order limit' }] })
  expect(new Set(seen)).toEqual(new Set([2]))
  sheet.getCell(row + 1, 2).value = 2
  expect(await readWorkbookForm(definition, Buffer.from(await book.xlsx.writeBuffer()), options)).toMatchObject({ success: true, data: { ceiling: 10, items: [data.items[1], data.items[0]] } })
})

it('signs validation semantics independently of directive layout and messages', async () => {
  const book = new ExcelJS.Workbook()
  const field = book.addWorksheet('Data').getCell('A1')
  field.value = '{v}{@validate:required|number|min:1}'
  const config = await importWorkbookXlsx(await saveWorkbook(book))
  const bytes = await renderWorkbookForm(config, { v: 2 })
  field.value = '{v}{@validate:required}{@validate:number}{@validate:min:1; message="New message"}{@validationMessage:min:"Fallback message"}'
  const changed = await importWorkbookXlsx(await saveWorkbook(book))
  expect(await readWorkbookForm(changed, bytes)).toEqual({ success: true, data: { v: 2 } })
  field.value = '{v}{@validate:required|number|min:3}'
  expect(await readWorkbookForm(await importWorkbookXlsx(await saveWorkbook(book)), bytes)).toMatchObject({ success: false, issues: [{ phase: 'structure' }] })
})

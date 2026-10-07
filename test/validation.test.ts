import { expect, it, vi } from 'vitest'
import ExcelJS from 'exceljs'
import { definition, data, options } from '../examples/validation/definition'
import { createValidation } from '../src/core/validation'
import type { ValidationRule } from '../src/index'
import { parseValidation, registerValidationRule, ValidationExecutionError, resolveWorkbook, importWorkbookXlsx, renderWorkbookReport, readWorkbookForm, renderWorkbookForm } from '../src/index'
import { openWorkbook as load, importAuthoredWorkbook, saveWorkbook } from './xlsx'

const context = { root: {}, current: {}, path: '$data.amount' }

it('carries ordered repeated rules and escaped JSON arguments/messages through field tags and XLSX', async () => {
  const pipeline = 'required|probe:"001",[1,true],{"text":"a|b:c,=;{d}"}|probe:false'
  const validation = parseValidation(pipeline)
  expect(validation[1].args).toEqual(['001', [1, true], { text: 'a|b:c,=;{d}' }])
  const rules = { validation: validation.map((use, index) => ({ ...use, ...(index === 1 ? { message: 'Ошибка "x" | : = ; {a}\nnext' } : {}) })), validationMessages: { probe: 'Fallback: "x"' } }
  const tag = '{amount}{@validate:required}{@validate:probe:"001",[1,true],{"text":"a|b:c,=;{d}"}; message=' + JSON.stringify(rules.validation[1].message) + '}{@validate:probe:false}{@validationMessage:probe:' + JSON.stringify(rules.validationMessages.probe) + '}'
  const book = new ExcelJS.Workbook()
  book.addWorksheet('Data').getCell('A1').value = tag
  const restored = await importWorkbookXlsx(await saveWorkbook(book))
  expect(resolveWorkbook(restored, { amount: 10 }).sheets[0].cells[0].rules).toEqual(rules)
  for (const invalid of ['{x}{y}', '{x}{@unknown:a}', '{x}{@validate:a|b; message="ambiguous"}', '{x}{@validationMessage:a:"unclosed}', '{x}{@choice:Items; key=id; key=other; label=name}']) {
    await expect(importAuthoredWorkbook(book => book.addWorksheet('Invalid').getCell('A1').value = invalid)).rejects.toThrow()
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

it('renders values and empty repeats without looking up or executing validation handlers', async () => {
  const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
  const validate = vi.fn(() => false)
  const validateArgs = vi.fn(() => false)
  registerValidationRule('testRenderOnly', { validate, validateArgs })
  try {
    const config = await importAuthoredWorkbook(book => {
      book.addWorksheet('Data').addRows([
        ['{v}{@validate:required|number|min:1|testRenderOnly:2}'],
        ['{blank}{@validate:required}'],
      ])
    })
    const input = { v: 'not a number', blank: null }
    expect(resolveWorkbook(config, input).sheets[0].cells[0].value).toEqual({ literal: input.v })
    for (const render of [renderWorkbookReport, renderWorkbookForm]) {
      const bytes = await render(config, input)
      expect((await load(bytes)).worksheets[0].getColumn('A').values).toContain(input.v)
      await expect(render(definition, { ceiling: null, items: [] })).resolves.toBeInstanceOf(Buffer)
    }
    expect(validateArgs).not.toHaveBeenCalled()
    expect(validate).not.toHaveBeenCalled()
    expect(warn).not.toHaveBeenCalled()
  }
  finally {
    warn.mockRestore()
  }
})

it('rejects unknown rules before reading, including empty repeats, and preserves occurrence indexes', async () => {
  const config = await importAuthoredWorkbook(book => book.addWorksheet('Data').addRows([
    ['Heading', { formula: '1+1' }], ['{#items}'], ['{.label}'],
    ['{.value}{@validate:external:{"key":1}|number|external|required|min:1}'], ['{/items}'],
  ]))
  for (const items of [[], [{ value: 0 }, { value: 2 }]]) {
    const bytes = await renderWorkbookForm(config, { items })
    await expect(readWorkbookForm(config, bytes)).rejects.toMatchObject({ issues: [{
      phase: 'template', code: 'invalid-rules', sheetName: 'Data', address: 'A4', message: 'Unknown validation rule: external',
    }] })
  }
  const bytes = await renderWorkbookForm(config, { items: [{ value: 0 }, { value: 2 }] })
  const handler = { validateArgs: () => true, validate: () => true }
  expect(await readWorkbookForm(config, bytes, { validationRules: { external: handler } })).toMatchObject({ success: false,
    issues: [{ rule: 'min', index: 4, path: '$data.items[0].value' }],
  })
  const runtime = { validationRules: { external: { ...handler, validate: () => false, message: 'Domain check failed' } } }
  expect(await readWorkbookForm(config, bytes, runtime)).toMatchObject({ success: false, issues: [
    { rule: 'external', index: 0, message: 'Domain check failed' },
    { rule: 'external', index: 0, message: 'Domain check failed' },
  ] })
})

it('rejects a misspelled required rule instead of accepting a blank field', async () => {
  const config = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = '{amount}{@validate:requred|number}')
  const bytes = await renderWorkbookForm(config, { amount: null })
  await expect(readWorkbookForm(config, bytes)).rejects.toThrow('Unknown validation rule: requred')
})

it('keeps malformed handlers, known arguments and execution failures separate from value issues', async () => {
  expect(() => createValidation({ validationRules: { required: { validate: () => true } } })('required')).toThrow('reserved')
  expect(() => createValidation(options)('decimalPlaces:bad')).toThrow('Invalid arguments')
  const config = await importAuthoredWorkbook(book => {
    book.addWorksheet('Data').getCell('A1').value = '{v}{@validate:broken}'
  })
  for (const validate of [() => {
    throw new RangeError('handler bug')
  }, () => Promise.resolve(true)]) {
    const runtime = { validationRules: { broken: { validate: validate as unknown as ValidationRule['validate'] } } }
    const form = await renderWorkbookForm(config, { v: 1 })
    await expect(readWorkbookForm(config, form, runtime)).rejects.toBeInstanceOf(ValidationExecutionError)
  }
})

it('ignores unused runtime options and keeps validation failures when message formatting fails', async () => {
  const config = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = '{amount}{@validate:required}')
  const custom = await importAuthoredWorkbook(book => book.addWorksheet('Input').getCell('A1').value = '{amount}{@validate:custom}')
  const bytes = await renderWorkbookForm(config, { amount: null })
  const warning = vi.spyOn(console, 'warn').mockImplementation(() => {})
  try {
    const unused = { validationRules: { unused: {} as ValidationRule }, validationMessages: { unused: 42 as unknown as string } }
    expect(await readWorkbookForm(config, bytes, unused)).toMatchObject({ success: false, data: { amount: null }, issues: [{ rule: 'required', message: 'a value is required' }] })
    expect(warning).not.toHaveBeenCalled()
    for (const message of [42 as unknown as string, () => {
      throw new Error('message failed')
    }]) {
      expect(await readWorkbookForm(config, bytes, { validationMessages: { required: message } })).toMatchObject({ success: false,
        data: { amount: null }, issues: [{ rule: 'required', message: 'failed validation: required' }],
      })
      expect(await readWorkbookForm(custom, bytes, { validationRules: { custom: { validate: () => false, skipEmpty: false, message } } })).toMatchObject({ success: false,
        data: { amount: null }, issues: [{ rule: 'custom', message: 'failed validation: custom' }],
      })
    }
    expect(warning).toHaveBeenCalledTimes(4)
  }
  finally {
    warning.mockRestore()
  }
})

it('validates complete submitted records after sorting and blank-row removal, keeping actual cell addresses', async () => {
  const book = await load(await renderWorkbookForm(definition, data))
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

it('uses current field paths, validation and messages when reading an already issued form', async () => {
  const book = new ExcelJS.Workbook()
  const field = book.addWorksheet('Data').getCell('A1')
  field.value = '{v}{@validate:required|number|min:1}'
  const config = await importWorkbookXlsx(await saveWorkbook(book))
  const bytes = await renderWorkbookForm(config, { v: 2 })
  field.value = '{current}{@validate:required}{@validate:number}{@validate:min:1; message="New message"}{@validationMessage:min:"Fallback message"}'
  const changed = await importWorkbookXlsx(await saveWorkbook(book))
  expect(await readWorkbookForm(changed, bytes)).toEqual({ success: true, data: { current: 2 } })
  field.value = '{current}{@validate:required|number}{@validate:min:3; message="New message"}{@validationMessage:min:"Fallback message"}'
  expect(await readWorkbookForm(await importWorkbookXlsx(await saveWorkbook(book)), bytes)).toMatchObject({ success: false, issues: [{ phase: 'value', path: '$data.current', code: 'min', args: [3], message: 'New message' }] })
})

it('checks builtin predicates without converting present zero, false or numeric text', () => {
  const prepare = createValidation()
  expect(prepare('required|number|min:0')(0, context)).toBeUndefined()
  expect(prepare('required|boolean')(false, context)).toBeUndefined()
  expect(prepare('number')('', context)).toBeUndefined()
  expect(prepare('required')('   ', context)?.code).toBe('required')
  expect(prepare('number')('0', context)?.code).toBe('number')
  expect(prepare('number|max:24')(25, context)?.code).toBe('max')
  expect(prepare('string|maxLength:0')('a', context)?.code).toBe('maxLength')
  expect(() => prepare('max:bad')).toThrow('Invalid arguments')
})

it('renders invalid input and reports nested field errors only when reading the completed form', async () => {
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Input').addRows([
    ['{#groups}'], ['{.name}{@validate:required|string|maxLength:12}'], ['{#.items}'],
    ['{?.hours}{@validate:required|number|min:0|max:24}', '{.approved}{@validate:required|boolean}', '{.code}{@validate:string|maxLength:5}'],
    [null, null, '{/.items}'], [null, null, '{/groups}'],
  ]))
  const data = { groups: [{ name: 'Main', items: [{ hours: -1, approved: 'false', code: 'too-long' }, { approved: false, code: '' }] }] }
  expect(() => resolveWorkbook(template, data)).not.toThrow()
  await expect(renderWorkbookReport(template, data)).resolves.toBeInstanceOf(Buffer)
  const result = await readWorkbookForm(template, await renderWorkbookForm(template, data))
  expect(result).toMatchObject({ success: false, issues: [
    { code: 'min', path: '$data.groups[0].items[0].hours' },
    { code: 'boolean', path: '$data.groups[0].items[0].approved' },
    { code: 'maxLength', path: '$data.groups[0].items[0].code' },
    { code: 'required', path: '$data.groups[0].items[1].hours' },
  ] })
})

it('registers validation globally without replacement, runs it only on read and allows per-call overrides', async () => {
  const validate = vi.fn(value => value === 'accepted')
  registerValidationRule('testRegistered', { validate, message: 'Expected accepted' })
  expect(() => registerValidationRule('testRegistered', { validate: () => true })).toThrow('already registered')
  const template = await importAuthoredWorkbook(book => book.addWorksheet('Data').getCell('A1').value = '{v}{@validate:testRegistered}')
  const bytes = await renderWorkbookForm(template, { v: 'rejected' })
  await renderWorkbookReport(template, { v: 'rejected' })
  expect(validate).not.toHaveBeenCalled()
  expect(await readWorkbookForm(template, bytes)).toMatchObject({ success: false, issues: [{ phase: 'value', rule: 'testRegistered', message: 'Expected accepted' }] })
  expect(await readWorkbookForm(template, bytes, { validationRules: { testRegistered: { validate: () => true } } })).toEqual({ success: true, data: { v: 'rejected' } })
  expect(() => registerValidationRule('required', { validate: () => true })).toThrow('reserved')
})

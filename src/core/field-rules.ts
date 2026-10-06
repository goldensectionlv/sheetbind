import { isDictionaryName } from './dictionaries'
import { parseChoiceRule } from './choices'
import type { ChoiceRule } from './choices'
import { assertJson, isDataObject } from './json'
import { isBlank, parseValidation } from './validation'
import type { Validation } from './validation'
import { parseFormatting } from './formatters'
import type { Formatting } from './formatters'

/** Serializable field behavior shared by template I/O, execution and form reading. */
export interface FieldRules {
  readonly validation?: Validation
  readonly validationMessages?: Readonly<Record<string, string>>
  readonly list?: string
  readonly choice?: ChoiceRule
  readonly format?: Formatting
}

export function parseFieldRules(value: unknown): FieldRules {
  assertJson(value)
  if (!isDataObject(value) || Object.keys(value).some(key => !['validation', 'validationMessages', 'list', 'choice', 'format'].includes(key))) {
    throw new SyntaxError('Unknown or invalid field rules')
  }
  const { validation, validationMessages, list, choice, format } = value
  if (format !== undefined && (list !== undefined || choice !== undefined)) {
    throw new SyntaxError('Formatting cannot change list or choice labels')
  }
  if (list !== undefined && !isDictionaryName(list)) {
    throw new SyntaxError('A list requires a dictionary name')
  }
  if (list !== undefined && choice !== undefined) {
    throw new SyntaxError('A field cannot declare both list and choice')
  }
  if (validationMessages !== undefined && (!isDataObject(validationMessages) || Object.values(validationMessages).some(text => typeof text !== 'string'))) {
    throw new SyntaxError('Field validation messages must be strings keyed by rule name')
  }
  if (validationMessages) {
    parseValidation(Object.keys(validationMessages).map(rule => ({ rule })))
  }
  const uses = validation === undefined ? [] : parseValidation(validation)
  return {
    ...(uses.length ? { validation: uses } : {}),
    ...(validationMessages && Object.keys(validationMessages).length ? { validationMessages: structuredClone(validationMessages) as Record<string, string> } : {}),
    ...(list !== undefined ? { list: list as string } : {}),
    ...(choice !== undefined ? { choice: parseChoiceRule(choice) } : {}),
    ...(format !== undefined ? { format: parseFormatting(format) } : {}),
  }
}

export function hasValidation(rules: FieldRules | undefined, name: string): boolean {
  const validation = rules?.validation
  return !!validation && (typeof validation === 'string' ? parseValidation(validation) : validation).some(use => use.rule === name)
}

export function validateList(value: unknown, rules: FieldRules, choices?: readonly string[]): { code: string, message: string } | undefined {
  if (!rules.list || isBlank(value)) {
    return undefined
  }
  if (!choices) {
    return { code: 'missing-dictionary', message: `dictionary ${rules.list} was not supplied` }
  }
  if (!choices.includes(value as string)) {
    return { code: 'list', message: `must be a value from ${rules.list}` }
  }
  return undefined
}

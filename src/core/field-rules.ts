import { isDictionaryName } from './dictionaries'
import type { ChoiceRule } from './choices'
import { isBlank, parseValidation } from './validation'
import type { Validation, ValidationRuleUse } from './validation'
import { parseFormatting } from './formatters'
import type { Formatting } from './formatters'
import { isDataPath, parseDataReference } from './template'
import type { ValueExpression } from './template'
import { parseRuleArgument, splitRuleText } from './rule-syntax'

/** Serializable field behavior shared by template I/O, execution and form reading. */
export interface FieldRules {
  readonly validation?: Validation
  readonly validationMessages?: Readonly<Record<string, string>>
  readonly list?: string
  readonly choice?: ChoiceRule
  readonly format?: Formatting
}

function tags(text: string): string[] {
  const result: string[] = []
  let depth = 0
  let start = 0
  let quoted = false
  let escaped = false
  for (let index = 0; index < text.length; index++) {
    const char = text[index]
    if (!depth && char !== '{' && char.trim()) {
      throw new SyntaxError('Only a binding and its directives may occupy a field cell')
    }
    if (quoted) {
      if (escaped) {
        escaped = false
      }
      else if (char === '\\') {
        escaped = true
      }
      else if (char === '"') {
        quoted = false
      }
    }
    else if (char === '"' && depth) {
      quoted = true
    }
    else if (char === '{') {
      if (!depth++) {
        start = index + 1
      }
    }
    else if (char === '}') {
      if (!depth) {
        throw new SyntaxError('Unexpected closing brace')
      }
      if (!--depth) {
        result.push(text.slice(start, index).trim())
      }
    }
  }
  if (depth || quoted || !result.length) {
    throw new SyntaxError('Unclosed or empty field tag')
  }
  return result
}

function assignments(parts: string[], allowed: readonly string[]): Record<string, unknown> {
  const result: Record<string, unknown> = {}
  for (const part of parts) {
    const [name, value, extra] = splitRuleText(part, '=')
    if (!allowed.includes(name) || value === undefined || extra !== undefined || Object.hasOwn(result, name)) {
      throw new SyntaxError(`Invalid or duplicate directive option: ${name}`)
    }
    result[name] = name === 'message' || value.startsWith('"') ? parseRuleArgument(value) : value
  }
  return result
}

/** One binding and its ordered field directives in the same cell. */
export function parseFieldTag(text: string): { value: Exclude<ValueExpression, { literal: unknown }>, rules?: FieldRules } {
  const parts = tags(text)
  const bindings = parts.filter(part => !part.startsWith('@'))
  if (bindings.length !== 1) {
    throw new SyntaxError('A field cell requires exactly one binding')
  }
  const [binding, ...pipeline] = splitRuleText(bindings[0], '|')
  const optional = binding.startsWith('?')
  const value = { ...parseDataReference(optional ? binding.slice(1) : binding), optional }
  const validation: ValidationRuleUse[] = []
  const messages: Record<string, string> = {}
  const rules: { -readonly [Key in keyof FieldRules]: FieldRules[Key] } = {}
  if (pipeline.length) {
    format(pipeline.join('|'))
  }
  const directives: Record<string, (text: string) => void> = { validate, validationMessage, choice, list, format }
  function format(text: string): void {
    if (rules.format !== undefined) {
      throw new SyntaxError('A field has one formatting pipeline')
    }
    const chain = parseFormatting(text)
    if (!chain.length) {
      throw new SyntaxError('A formatting directive needs a formatter')
    }
    rules.format = chain
  }
  function validate(text: string): void {
    const [pipeline, ...options] = splitRuleText(text, ';')
    const chain = parseValidation(pipeline)
    if (!chain.length) {
      throw new SyntaxError('A validation directive needs a rule')
    }
    const inline = assignments(options, ['message'])
    if (inline.message !== undefined && (chain.length !== 1 || typeof inline.message !== 'string')) {
      throw new SyntaxError('An inline message belongs to one rule occurrence and must be text')
    }
    validation.push(...chain.map(use => ({ ...use, ...(inline.message !== undefined ? { message: inline.message as string } : {}) })))
  }
  function validationMessage(text: string): void {
    const [name, value, extra] = splitRuleText(text, ':')
    parseValidation([{ rule: name }])
    if (value === undefined || extra !== undefined || Object.hasOwn(messages, name)) {
      throw new SyntaxError('A validation message directive requires one rule name and one message')
    }
    const parsed = parseRuleArgument(value)
    if (typeof parsed !== 'string') {
      throw new SyntaxError('Validation messages must be text')
    }
    messages[name] = parsed
  }
  function choice(text: string): void {
    if (rules.choice !== undefined) {
      throw new SyntaxError('A field has one choice source')
    }
    const [name, ...options] = splitRuleText(text, ';')
    const { key, label, return: mode, emptySource } = assignments(options, ['key', 'label', 'return', 'emptySource'])
    if (!isDataPath(key) || !isDataPath(label)) {
      throw new SyntaxError('Choice requires a source and safe key/label paths')
    }
    if (mode !== undefined && mode !== 'object' && mode !== 'key') {
      throw new SyntaxError('Choice return must be object or key')
    }
    if (emptySource !== undefined && (emptySource !== 'input' || mode !== 'key')) {
      throw new SyntaxError('Choice emptySource=input requires return=key')
    }
    const reference = name.startsWith('.') || name.startsWith('$root.') ? parseDataReference(name) : undefined
    if (!reference && !isDictionaryName(name)) {
      throw new SyntaxError('Choice dictionary must be a named source')
    }
    const source = reference ? reference.from === 'root' ? reference : { path: reference.path } : { dictionary: name }
    rules.choice = { source, key, label, ...(mode !== undefined ? { return: mode } : {}), ...(emptySource !== undefined ? { emptySource } : {}) }
  }
  function list(text: string): void {
    if (rules.list !== undefined) {
      throw new SyntaxError('A field has one string list')
    }
    const name = text.trim()
    if (!isDictionaryName(name)) {
      throw new SyntaxError('A list requires a dictionary name')
    }
    rules.list = name
  }
  for (const part of parts.filter(part => part.startsWith('@'))) {
    const colon = part.indexOf(':')
    const name = part.slice(1, colon)
    if (colon < 0 || !Object.hasOwn(directives, name)) {
      throw new SyntaxError(`Unknown field directive: ${part}`)
    }
    directives[name](part.slice(colon + 1).trim())
  }
  if (validation.length) {
    rules.validation = validation
  }
  if (Object.keys(messages).length) {
    rules.validationMessages = messages
  }
  if (rules.format !== undefined && (rules.list !== undefined || rules.choice !== undefined)) {
    throw new SyntaxError('Formatting cannot change list or choice labels')
  }
  if (rules.list !== undefined && rules.choice !== undefined) {
    throw new SyntaxError('A field cannot declare both list and choice')
  }
  return { value, ...(Object.keys(rules).length ? { rules } : {}) }
}

export function hasValidation(rules: FieldRules | undefined, name: string): boolean {
  const validation = rules?.validation
  return !!validation && (typeof validation === 'string' ? parseValidation(validation) : validation).some(use => use.rule === name)
}

export function validateList(value: unknown, rules: FieldRules, choices?: readonly unknown[]): { code: string, message: string } | undefined {
  if (!rules.list || !choices?.length || !choices.every(item => typeof item === 'string') || isBlank(value)) {
    return undefined
  }
  if (!choices.includes(value as string)) {
    return { code: 'list', message: `must be a value from ${rules.list}` }
  }
  return undefined
}

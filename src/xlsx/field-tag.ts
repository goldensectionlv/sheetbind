import { parseDataReference } from '../core/reference'
import type { ValueExpression } from '../core/template'
import { parseFieldRules } from '../core/field-rules'
import type { FieldRules } from '../core/field-rules'
import { parseValidation } from '../core/validation'
import type { ValidationRuleUse } from '../core/validation'
import { parseRuleArgument, splitRuleText } from '../core/rule-syntax'
import { parseFormatting } from '../core/formatters'

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
  const rules: Record<string, unknown> = {}
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
    const settings = assignments(options, ['key', 'label', 'return', 'emptySource'])
    const reference = name.startsWith('.') || name.startsWith('$root.') ? parseDataReference(name) : undefined
    const source = reference ? reference.from === 'root' ? reference : { path: reference.path } : { dictionary: name }
    rules.choice = { source, ...settings }
  }
  function list(text: string): void {
    if (rules.list !== undefined) {
      throw new SyntaxError('A field has one string list')
    }
    rules.list = text.trim()
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
  return { value, ...(Object.keys(rules).length ? { rules: parseFieldRules(rules) } : {}) }
}

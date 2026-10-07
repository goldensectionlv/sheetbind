import { assertJson, isDataObject } from './json'
import type { JsonValue } from './json'
import { parseRuleArgument, splitRuleText } from './rule-syntax'

export interface ValidationRuleUse {
  readonly rule: string
  readonly args?: readonly JsonValue[]
  readonly message?: string
}
export type Validation = string | readonly ValidationRuleUse[]
export interface ValidationContext {
  readonly root: Readonly<Record<string, unknown>>
  readonly current: Readonly<Record<string, unknown>>
  readonly path: string
}
export interface ValidationMessageContext extends ValidationContext {
  readonly value: unknown
  readonly rule: string
  readonly args: readonly JsonValue[]
  readonly index: number
}
export type ValidationMessage = string | ((context: ValidationMessageContext) => string)
export interface ValidationRule {
  readonly validate: (value: unknown, args: readonly JsonValue[], context: ValidationContext) => boolean
  /** Without this guard a rule accepts no arguments. Runs before any field values. */
  readonly validateArgs?: (args: readonly JsonValue[]) => boolean
  readonly skipEmpty?: boolean
  readonly message?: ValidationMessage
}
export interface ValidationOptions {
  readonly validationRules?: Readonly<Record<string, ValidationRule>>
  readonly validationMessages?: Readonly<Record<string, ValidationMessage>>
}
export interface ValidationIssue {
  readonly code: string
  readonly rule: string
  readonly args: readonly JsonValue[]
  readonly index: number
  readonly message: string
}
export type ValidateValue = (value: unknown, context: ValidationContext) => ValidationIssue | undefined

export class ValidationExecutionError extends Error {
  constructor(readonly rule: string, readonly path: string, cause: unknown) {
    super(`${path}: validation rule ${rule} failed to execute: ${cause instanceof Error ? cause.message : String(cause)}`, { cause })
    this.name = 'ValidationExecutionError'
  }
}

const ruleName = /^[A-Za-z_][A-Za-z0-9_.-]*$/

export function parseValidation(value: unknown): readonly ValidationRuleUse[] {
  if (typeof value === 'string') {
    if (!value.trim()) {
      return []
    }
    return parseValidation(splitRuleText(value, '|').map(text => {
      const [rule, ...parts] = splitRuleText(text, ':')
      if (parts.length > 1) {
        throw new SyntaxError('A rule has one colon followed by comma-separated arguments')
      }
      return { rule, ...(parts.length ? { args: splitRuleText(parts[0], ',').map(parseRuleArgument) } : {}) }
    }))
  }
  if (!Array.isArray(value)) {
    throw new SyntaxError('Validation must be a pipeline or an ordered rule array')
  }
  assertJson(value)
  return value.map(use => {
    if (!isDataObject(use) || Object.keys(use).some(key => !['rule', 'args', 'message'].includes(key))
      || typeof use.rule !== 'string' || !ruleName.test(use.rule) || ['__proto__', 'constructor', 'prototype'].includes(use.rule)
      || use.args !== undefined && !Array.isArray(use.args)
      || use.message !== undefined && typeof use.message !== 'string') {
      throw new SyntaxError('Invalid validation rule occurrence')
    }
    return { rule: use.rule, ...(use.args?.length ? { args: structuredClone(use.args) as JsonValue[] } : {}), ...(use.message !== undefined ? { message: use.message } : {}) }
  })
}

/** Individual messages belong to structured occurrences or separate XLSX directives. */
export function formatValidation(value: Validation): string {
  return parseValidation(value).map(use => use.rule + (use.args?.length ? ':' + use.args.map(arg => JSON.stringify(arg)).join(',') : '')).join('|')
}

export function isBlank(value: unknown): boolean {
  return value === undefined || value === null || typeof value === 'string' && value.trim() === ''
}

function numericArgument(args: readonly JsonValue[]): boolean {
  return args.length === 1 && typeof args[0] === 'number' && Number.isFinite(args[0])
}
function lengthArgument(args: readonly JsonValue[]): boolean {
  return numericArgument(args) && Number.isSafeInteger(args[0]) && (args[0] as number) >= 0
}
function required(value: unknown): boolean {
  return !isBlank(value)
}
function string(value: unknown): boolean {
  return typeof value === 'string'
}
function number(value: unknown): boolean {
  return typeof value === 'number' && Number.isFinite(value)
}
function boolean(value: unknown): boolean {
  return typeof value === 'boolean'
}
function object(value: unknown): boolean {
  return isDataObject(value)
}
function min(value: unknown, args: readonly JsonValue[]): boolean {
  return number(value) && (value as number) >= (args[0] as number)
}
function max(value: unknown, args: readonly JsonValue[]): boolean {
  return number(value) && (value as number) <= (args[0] as number)
}
function maxLength(value: unknown, args: readonly JsonValue[]): boolean {
  return string(value) && (value as string).length <= (args[0] as number)
}

const builtins: Readonly<Record<string, ValidationRule>> = Object.freeze({
  required: { validate: required, skipEmpty: false, message: 'a value is required' },
  string: { validate: string, message: 'expected a string value' },
  number: { validate: number, message: 'expected a finite number' },
  boolean: { validate: boolean, message: 'expected a boolean value' },
  object: { validate: object, message: 'expected an object value' },
  min: { validate: min, validateArgs: numericArgument, message: ({ args }) => `must be at least ${args[0]}` },
  max: { validate: max, validateArgs: numericArgument, message: ({ args }) => `must be at most ${args[0]}` },
  maxLength: { validate: maxLength, validateArgs: lengthArgument, message: ({ args }) => `must contain at most ${args[0]} UTF-16 code units` },
})

const registered = new Map<string, ValidationRule>()

function assertHandler(name: string, rule: ValidationRule): void {
  if (typeof name !== 'string' || !ruleName.test(name) || Object.hasOwn(builtins, name) || ['__proto__', 'constructor', 'prototype'].includes(name)) {
    throw new SyntaxError(`Validation rule name is reserved or invalid: ${name}`)
  }
  if (!rule || typeof rule.validate !== 'function' || rule.validateArgs !== undefined && typeof rule.validateArgs !== 'function'
      || rule.skipEmpty !== undefined && typeof rule.skipEmpty !== 'boolean'
      || rule.message !== undefined && typeof rule.message !== 'string' && typeof rule.message !== 'function') {
    throw new SyntaxError(`Invalid validation handler: ${name}`)
  }
}

/** Register once at application startup; per-read rules override registered rules. */
export function registerValidationRule(name: string, rule: ValidationRule): void {
  assertHandler(name, rule)
  if (registered.has(name)) {
    throw new SyntaxError(`Validation rule is already registered: ${name}`)
  }
  registered.set(name, { ...rule })
}

/** Form reading owns its handlers; rendering only carries rule definitions. */
export function createValidation(options: ValidationOptions = {}) {
  const rules = new Map([...Object.entries(builtins), ...registered])
  for (const [name, rule] of Object.entries(options.validationRules ?? {})) {
    assertHandler(name, rule)
    rules.set(name, { ...rule })
  }
  for (const message of Object.values(options.validationMessages ?? {})) {
    if (typeof message !== 'string' && typeof message !== 'function') {
      throw new SyntaxError('Runtime validation messages must be strings or functions')
    }
  }
  return function prepare(validation: Validation = [], messages: Readonly<Record<string, string>> = {}): ValidateValue {
    const chain = parseValidation(validation).map((use, index) => {
      const handler = rules.get(use.rule)
      if (!handler) {
        throw new SyntaxError(`Unknown validation rule: ${use.rule}`)
      }
      const args = use.args ?? []
      if ((handler.validateArgs ? handler.validateArgs(args) : args.length === 0) !== true) {
        throw new SyntaxError(`Invalid arguments for validation rule: ${use.rule}`)
      }
      return { use, args, handler, index }
    })
    return function validate(value, context) {
      for (const { use, args, handler, index } of chain) {
        if (handler.skipEmpty !== false && isBlank(value)) {
          continue
        }
        try {
          const valid = handler.validate(value, args, context)
          if (typeof valid !== 'boolean') {
            throw new TypeError('expected a synchronous boolean')
          }
          if (valid) {
            continue
          }
          const fallback = Object.hasOwn(messages, use.rule) ? messages[use.rule] : undefined
          const runtime = Object.hasOwn(options.validationMessages ?? {}, use.rule) ? options.validationMessages![use.rule] : undefined
          const message = use.message ?? fallback ?? runtime ?? handler.message ?? `failed validation: ${use.rule}`
          const details = { ...context, value, rule: use.rule, args, index }
          const text = typeof message === 'function' ? message(details) : message
          if (typeof text !== 'string') {
            throw new TypeError('expected a string message')
          }
          return { code: use.rule, rule: use.rule, args: structuredClone(args), index, message: text }
        }
        catch (error) {
          throw new ValidationExecutionError(use.rule, context.path, error)
        }
      }
      return undefined
    }
  }
}

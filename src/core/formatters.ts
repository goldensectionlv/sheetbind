import { assertJson, isDataObject } from './json'
import type { JsonValue } from './json'
import { parseRuleArgument, splitRuleText } from './rule-syntax'
import type { TemplateValue } from './template'
import { isBlank } from './validation'

export interface FormatterUse {
  readonly formatter: string
  readonly args?: readonly JsonValue[]
}
export type Formatting = string | readonly FormatterUse[]
/** Formatters run synchronously when issuing a form or rendering a report, never on read. */
export type Formatter = (value: TemplateValue, args: readonly JsonValue[]) => TemplateValue

const namePattern = /^[A-Za-z_][A-Za-z0-9_.-]*$/
function validName(name: string): boolean {
  return typeof name === 'string' && namePattern.test(name) && !['__proto__', 'constructor', 'prototype'].includes(name)
}

export function parseFormatting(value: unknown): readonly FormatterUse[] {
  if (typeof value === 'string') {
    if (!value.trim()) {
      return []
    }
    return parseFormatting(splitRuleText(value, '|').map(text => {
      const [formatter, ...parts] = splitRuleText(text, ':')
      if (parts.length > 1) {
        throw new SyntaxError('A formatter has one colon followed by comma-separated arguments')
      }
      return { formatter, ...(parts.length ? { args: splitRuleText(parts[0], ',').map(parseRuleArgument) } : {}) }
    }))
  }
  if (!Array.isArray(value)) {
    throw new SyntaxError('Formatting must be a pipeline or an ordered formatter array')
  }
  assertJson(value)
  return value.map(use => {
    if (!isDataObject(use) || Object.keys(use).some(key => !['formatter', 'args'].includes(key))
      || typeof use.formatter !== 'string' || !validName(use.formatter) || use.args !== undefined && !Array.isArray(use.args)) {
      throw new SyntaxError('Invalid formatter occurrence')
    }
    return { formatter: use.formatter, ...(use.args?.length ? { args: structuredClone(use.args) as JsonValue[] } : {}) }
  })
}

const builtins: Readonly<Record<string, Formatter>> = {
  float(value) {
    if (isBlank(value)) {
      return value
    }
    const number = Number(value)
    return Number.isFinite(number) ? number : value
  },
  bool_replace(value, args) {
    return isBlank(value) ? value : args[value ? 0 : 1] as string
  },
  format_date(value, args) {
    if (value === null || typeof value === 'boolean' || isBlank(value)) {
      return value
    }
    const date = new Date(value)
    if (Number.isNaN(date.getTime())) {
      return value
    }
    const pad = (part: number) => String(part).padStart(2, '0')
    const tokens: Record<string, string> = {
      YYYY: String(date.getFullYear()), YY: pad(date.getFullYear() % 100),
      MM: pad(date.getMonth() + 1), M: String(date.getMonth() + 1), DD: pad(date.getDate()), D: String(date.getDate()),
      HH: pad(date.getHours()), H: String(date.getHours()), mm: pad(date.getMinutes()), m: String(date.getMinutes()),
      ss: pad(date.getSeconds()), s: String(date.getSeconds()),
    }
    return ((args[0] as string | undefined) ?? 'DD.MM.YYYY').replace(/YYYY|YY|MM|M|DD|D|HH|H|mm|m|ss|s/g, token => tokens[token])
  },
}
const registered = new Map<string, Formatter>()

/** Register once at application startup; built-in names cannot be replaced. */
export function registerFormatter(name: string, formatter: Formatter): void {
  if (!validName(name) || Object.hasOwn(builtins, name)) {
    throw new SyntaxError(`Formatter name is reserved or invalid: ${name}`)
  }
  if (typeof formatter !== 'function') {
    throw new TypeError(`Invalid formatter: ${name}`)
  }
  registered.set(name, formatter)
}

export function prepareFormatting(formatting: Formatting): (value: TemplateValue) => TemplateValue {
  const chain = (typeof formatting === 'string' ? parseFormatting(formatting) : formatting).map(use => {
    const formatter = registered.get(use.formatter) ?? (Object.hasOwn(builtins, use.formatter) ? builtins[use.formatter] : undefined)
    if (!formatter) {
      throw new SyntaxError(`Unknown formatter: ${use.formatter}`)
    }
    const args = use.args ?? []
    if (use.formatter === 'float' && args.length) {
      throw new SyntaxError('float accepts no arguments')
    }
    if (use.formatter === 'bool_replace' && (args.length !== 2 || args.some(arg => typeof arg !== 'string'))) {
      throw new SyntaxError('bool_replace requires two text labels')
    }
    if (use.formatter === 'format_date' && (args.length > 1 || args.length === 1 && typeof args[0] !== 'string')) {
      throw new SyntaxError('format_date accepts one text mask')
    }
    return { ...use, formatter }
  })
  return value => chain.reduce((result, use) => {
    const formatted = use.formatter(result, structuredClone(use.args ?? []))
    if (formatted !== null && typeof formatted !== 'string' && typeof formatted !== 'boolean'
      && !(typeof formatted === 'number' && Number.isFinite(formatted))) {
      throw new TypeError('A formatter must return a synchronous finite scalar')
    }
    return formatted
  }, value)
}

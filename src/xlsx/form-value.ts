import type { Cell } from 'exceljs'
import type { TemplateValue } from '../core/template'

/** Read the submitted cell, including its current Excel number format, before validation. */
export function readFormValue(cell: Cell): TemplateValue | undefined {
  let value = cell.value
  if (value === null || value === undefined || value === '') {
    return null
  }
  if (value instanceof Date) {
    if (/\[(?:h+|m+|s+)\]/i.test(numberPattern(cell.numFmt || '', 0))) {
      return value.getTime() / 86400000 + 25569 - (cell.worksheet.workbook.properties.date1904 ? 1462 : 0)
    }
    return Number.isFinite(value.getTime()) ? value.toISOString() : undefined
  }
  if (typeof value === 'object') {
    if ('hyperlink' in value) {
      value = value.text
    }
    else if ('richText' in value) {
      value = value.richText.map(part => part.text).join('')
    }
    else {
      // Formula results may be stale; a formula is not a submitted scalar.
      return undefined
    }
  }
  if (typeof value === 'boolean') {
    return value
  }
  const number = typeof value === 'number' ? value : numericText(value)
  const pattern = numberPattern(cell.numFmt || 'General', number ?? 0)
  if (/^general$/i.test(pattern.trim()) && typeof value === 'string' && /^[+-]?0\d/.test(value.trim())) {
    return value
  }
  if (pattern.includes('@') && !/[0#?]/.test(pattern)) {
    return String(value) || null
  }
  if (number === undefined) {
    return value || null
  }
  return Number.isFinite(number) ? roundToFormat(number, pattern) : undefined
}

function numericText(value: string): number | undefined {
  const text = value.trim()
  // A single comma means a decimal separator; grouping punctuation is not removed.
  if (!/^[+-]?(?:\d+(?:[.,]\d*)?|[.,]\d+)(?:e[+-]?\d+)?$/i.test(text)) {
    return undefined
  }
  const number = Number(text.replace(',', '.'))
  return Number.isFinite(number) ? number : undefined
}

interface Section {
  pattern: string
  condition?: readonly [string, number]
}

/** Strip literals, colours and locale hints without mistaking their digits for precision. */
function numberPattern(format: string, value: number): string {
  const sections: Section[] = [{ pattern: '' }]
  for (const token of format.match(/"[^"]*"|\\.|_.|\*.|\[[^\]]*\]|[^]/g) ?? []) {
    if (token === ';') {
      sections.push({ pattern: '' })
      continue
    }
    const section = sections[sections.length - 1]
    const condition = /^\[(<=|>=|<>|=|<|>)([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)\]$/i.exec(token)
    if (condition) {
      section.condition = [condition[1], Number(condition[2])]
    }
    else if (/^\[(h+|m+|s+)\]$/i.test(token)) {
      section.pattern += token
    }
    else if (!['"', '\\', '_', '*', '['].includes(token[0])) {
      section.pattern += token
    }
  }
  const numeric = sections.slice(0, 3)
  if (numeric.some(section => section.condition)) {
    return numeric.find(section => !section.condition || matches(value, ...section.condition))?.pattern ?? ''
  }
  return numeric[value < 0 && numeric.length > 1 ? 1 : value === 0 && numeric.length > 2 ? 2 : 0].pattern
}

function matches(value: number, operator: string, limit: number): boolean {
  switch (operator) {
    case '<': return value < limit
    case '<=': return value <= limit
    case '>': return value > limit
    case '>=': return value >= limit
    case '=': return value === limit
    default: return value !== limit
  }
}

function roundToFormat(value: number, pattern: string): number {
  // General, dates/times and fractions do not declare a decimal precision.
  if (/general|[ymdhsa/]/i.test(pattern) || !/[0#?]/.test(pattern)) {
    return value
  }
  const scientific = /[eE][+-]?[0#?]+/.test(pattern)
  const mantissa = pattern.split(/[eE]/)[0]
  const decimals = /\.([0#?]+)/.exec(mantissa)?.[1].length ?? 0
  const percent = (mantissa.match(/%/g) ?? []).length
  const scaling = /[0#?](,+)(?=[^0#?]*$)/.exec(mantissa)?.[1].length ?? 0
  const exponent = scientific && value !== 0 ? Math.floor(Math.log10(Math.abs(value))) : 0
  const places = decimals + percent * 2 - scaling * 3 - exponent
  // Shift the decimal exponent instead of multiplying: 10.075 must round to 10.08.
  const [coefficient, power = '0'] = Math.abs(value).toString().split('e')
  const shifted = Number(`${coefficient}e${Number(power) + places}`)
  if (!Number.isFinite(shifted)) {
    return value
  }
  const [rounded, roundedPower = '0'] = Math.round(shifted).toString().split('e')
  const result = Number(`${rounded}e${Number(roundedPower) - places}`) * Math.sign(value)
  return Number.isFinite(result) ? result || 0 : value
}

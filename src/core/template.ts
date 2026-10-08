import type { JsonValue } from './json'
import { isDataObject } from './json'

/** Workbook binding values, data paths and diagnostics. */
export type TemplateValue = string | number | boolean | null
export type FieldValue = TemplateValue | Readonly<Record<string, unknown>>
export type ValueExpression = { readonly literal: TemplateValue }
  | (DataReference & { readonly optional?: boolean })

export interface Origin {
  readonly nodeId: string
  /** A concrete source path, not a stable record key for editable forms. */
  readonly dataPath: string
  readonly iterations: readonly { readonly nodeId: string, readonly index: number }[]
}
export interface TemplateIssue {
  readonly rule?: string
  readonly args?: readonly JsonValue[]
  readonly index?: number
  readonly phase: 'template' | 'data'
  readonly code: string
  readonly path: string
  readonly nodeId: string
  readonly message: string
}
export class TemplateError extends Error {
  constructor(readonly issues: readonly TemplateIssue[]) {
    super(issues.map(issue => `${issue.path}: ${issue.message}`).join('\n'))
    this.name = 'TemplateError'
  }
}
export interface DataReference {
  readonly path: string
  readonly from?: 'current' | 'root'
}

export function isDataPath(value: unknown): value is string {
  return typeof value === 'string' && /^[\p{ID_Start}_]\p{ID_Continue}*(?:\.[\p{ID_Start}_]\p{ID_Continue}*)*$/u.test(value)
    && !value.split('.').some(part => ['__proto__', 'prototype', 'constructor'].includes(part))
}

/** Dotted references traverse own properties of data objects, never collection indexes. */
export function readDataPath(value: unknown, path: string): unknown {
  return compileDataPath(path)(value)
}
export function compileDataPath(path: string): (value: unknown) => unknown {
  const parts = path.split('.')
  return value => {
    for (const part of parts) {
      value = isDataObject(value) && Object.hasOwn(value, part) ? value[part] : undefined
    }
    return value
  }
}
/** Human-readable references; stored definitions keep path and origin separate. */
export function parseDataReference(text: string): DataReference {
  const from = text.startsWith('$root.') ? 'root' : 'current'
  const path = text.startsWith('$root.') ? text.slice(6) : text.startsWith('.') ? text.slice(1) : text
  if (!isDataPath(path)) {
    throw new SyntaxError('Use a dotted path such as customer.name, or .name for the current item')
  }
  return { path, from }
}

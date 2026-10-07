import type { JsonValue } from './json'
import type { DataReference } from './reference'
export type { DataReference } from './reference'

/** Template semantics, independent of coordinates, document formats and I/O. */
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

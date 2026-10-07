import { TemplateError } from '../core/template'
import type { TemplateIssue } from '../core/template'
import { workbookCells, workbookRegions } from '../grid/workbook'
import type { WorkbookDefinition } from '../grid/workbook'

export interface TaggedXlsxIssue extends TemplateIssue {
  readonly sheetName?: string
  readonly address?: string
}

export class TaggedXlsxError extends TemplateError {
  declare readonly issues: readonly TaggedXlsxIssue[]

  constructor(issues: readonly (Omit<TaggedXlsxIssue, 'nodeId'> & { readonly nodeId?: string })[], options?: ErrorOptions) {
    const located = issues.map(issue => ({ ...issue, nodeId: issue.nodeId ?? '' }))
    super(located)
    this.issues = located
    this.message = located.map(issue => {
      const location = issue.sheetName && issue.address ? `${issue.sheetName}!${issue.address}` : undefined
      return `${location ? location + (issue.address === issue.path ? '' : ': ' + issue.path) : issue.path}: ${issue.message}`
    }).join('\n')
    this.name = 'TaggedXlsxError'
    if (options) {
      this.cause = options.cause
    }
  }
}

/** Locate execution errors in the authored template using their cell and region IDs. */
export function withTemplateLocations<T>(template: WorkbookDefinition, run: () => T): T {
  try {
    return run()
  }
  catch (error) {
    if (!(error instanceof TemplateError)) {
      throw error
    }
    const locations = new Map(template.sheets.flatMap(sheet => [
      ...workbookCells(sheet), ...workbookRegions(sheet),
    ].map(node => [node.id, { sheetName: sheet.name, address: node.xlsx.address }] as const)))
    throw new TaggedXlsxError(error.issues.map(issue => ({ ...issue, ...locations.get(issue.nodeId) })), { cause: error })
  }
}

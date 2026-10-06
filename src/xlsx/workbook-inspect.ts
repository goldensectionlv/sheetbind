import { prepareFormatting } from '../core/formatters'
import { workbookCells } from '../grid/workbook'
import type { TaggedXlsxIssue } from './tagged-template'
import { WorkbookTemplate } from './template'

export interface WorkbookTemplateFinding extends TaggedXlsxIssue {
  readonly severity: 'error' | 'warning'
}

/** Check render dependencies without data, formatter execution or value validation. */
export function inspectWorkbookTemplate(template: WorkbookTemplate, options: { readonly dictionaries?: readonly string[] } = {}): readonly WorkbookTemplateFinding[] {
  const findings: WorkbookTemplateFinding[] = []
  const known = options.dictionaries && new Set(options.dictionaries)
  for (const sheet of WorkbookTemplate.content(template).definition.sheets) {
    for (const cell of workbookCells(sheet)) {
      const location = { phase: 'template' as const, nodeId: cell.id, path: cell.xlsx?.address ?? cell.id, sheetName: sheet.name, address: cell.xlsx?.address }
      if (cell.rules?.format) {
        try {
          prepareFormatting(cell.rules.format)
        }
        catch (error) {
          findings.push({ ...location, code: 'invalid-format', severity: 'error', message: (error as Error).message })
        }
      }
      const choice = cell.rules?.choice?.source
      const dependencies = new Set([cell.rules?.list, choice && 'dictionary' in choice ? choice.dictionary : undefined])
      for (const name of dependencies) {
        if (name && known && !known.has(name)) {
          findings.push({ ...location, code: 'unknown-dict', severity: 'warning', message: `Unknown dictionary: ${name}` })
        }
      }
    }
  }
  return findings
}

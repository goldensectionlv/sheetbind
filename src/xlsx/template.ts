import type { WorkbookDefinition } from '../grid/workbook'

/** An imported XLSX and its compiled bindings, owned by the library. */
export class WorkbookTemplate {
  readonly #content: { readonly definition: WorkbookDefinition, readonly source: Uint8Array }

  constructor(definition: WorkbookDefinition, source: Uint8Array) {
    this.#content = { definition, source: Uint8Array.from(source) }
  }

  static content(template: WorkbookTemplate) {
    if (!(template instanceof WorkbookTemplate)) {
      throw new TypeError('Load a tagged XLSX with importWorkbookXlsx first')
    }
    return template.#content
  }
}

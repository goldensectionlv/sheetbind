import type { WorkbookDefinition } from '../grid/workbook'
import type { WorkbookResourceSource } from './workbook-resources'

/** An imported XLSX and its compiled bindings, owned by the library. */
export class WorkbookTemplate {
  readonly #content: { readonly definition: WorkbookDefinition, readonly source: Uint8Array, readonly resources: WorkbookResourceSource }

  constructor(definition: WorkbookDefinition, source: Uint8Array, resources: WorkbookResourceSource) {
    this.#content = { definition, source: Uint8Array.from(source), resources }
  }

  static content(template: WorkbookTemplate) {
    if (!(template instanceof WorkbookTemplate)) {
      throw new TypeError('Load a tagged XLSX with importWorkbookXlsx first')
    }
    return template.#content
  }
}

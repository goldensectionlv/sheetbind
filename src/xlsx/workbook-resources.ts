import type JSZip from 'jszip'
import { decodeXml, resolvePart, xmlAttributes, xmlElements } from './xml'
import { decodeXstring } from './report-text'

/** Relationship IDs share a namespace within one .rels part. */
export function workbookRelationshipId(relations: string, preferred: string): string {
  const ids = new Set(xmlElements(relations, 'Relationship').map(node => xmlAttributes(node).Id))
  while (ids.has(preferred)) {
    preferred += '_'
  }
  return preferred
}

/** Package paths and logical sheet names share the same decoded identity at every boundary. */
export async function workbookParts(zip: JSZip) {
  const workbook = await zip.file('xl/workbook.xml')!.async('string')
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
  const relationships = new Map(xmlElements(rels, 'Relationship').map(node => {
    const attr = xmlAttributes(node)
    return [attr.Id, attr]
  }))
  return new Map(xmlElements(workbook, 'sheet').map(node => {
    const attr = xmlAttributes(node)
    const relationship = relationships.get(attr['r:id'])!
    const name = decodeXstring(attr.name)
    return [name, { name, part: resolvePart('xl/workbook.xml', relationship.Target), node, relationship }]
  }))
}

export interface WorkbookResourceSource {
  readonly names: readonly string[]
  readonly references: ReadonlySet<string>
}

/** Inspect authored resources before any generated writer allocates names or formula ranges. */
export async function readWorkbookResources(zip: JSZip): Promise<WorkbookResourceSource> {
  const workbook = await zip.file('xl/workbook.xml')!.async('string')
  const names = xmlElements(workbook, 'definedName').map(node => decodeXstring(xmlAttributes(node.split('>')[0]).name))
  const references = new Set<string>()
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir || !path.startsWith('xl/') || !path.endsWith('.xml')) {
      continue
    }
    const xml = await file.async('string')
    for (const match of xml.matchAll(/<(f|formula|formula1|formula2|definedName|calculatedColumnFormula|totalsRowFormula)\b[^>]*>([^<]*)<\/\1>/g)) {
      for (const name of decodeXstring(decodeXml(match[2])).match(/_sb_ref_[a-z0-9_]+/gi) ?? []) {
        references.add(name.toLowerCase())
      }
    }
  }
  return { names, references }
}

/** One allocation scope per render, seeded from every authored name, including local names. */
export function createWorkbookResources(source: WorkbookResourceSource = { names: [], references: new Set() }) {
  const names = new Set(source.names.map(name => name.toLowerCase()))
  return {
    references: source.references,
    allocate(preferred: string, fixed = false): string {
      let name = preferred
      const numbered = /^(.*_)(\d+)$/.exec(preferred)
      const prefix = numbered ? numbered[1] : preferred + '_'
      let suffix = numbered ? Number(numbered[2]) : 1
      while (names.has(name.toLowerCase())) {
        if (fixed) {
          throw new RangeError(`Generated choice range conflicts with an authored name: ${preferred}`)
        }
        name = `${prefix}${++suffix}`
      }
      names.add(name.toLowerCase())
      return name
    },
  }
}
export type WorkbookResources = ReturnType<typeof createWorkbookResources>

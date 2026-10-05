import type JSZip from 'jszip'
import type { GridRange } from '../grid/geometry'
import { formatAddress, parseAddress, parseRange } from './addresses'
import { sourceFormula } from './source-coordinates'
import type { SourceCoordinates } from './source-coordinates'
import { decodeXml, encodeXml, resolvePart, setXmlAttributes, xmlAttributes, xmlElements } from './xml'

/** An absent range removes the filter or sort operation, including its dependent settings. */
export function relocateFilterRanges(xml: string, map: SourceCoordinates): string {
  function relocate(element: string): string {
    const ref = xmlAttributes(element.split('>')[0]).ref
    if (!ref) {
      return element
    }
    const range = map.range(ref)
    return range ? setXmlAttributes(element, { ref: range }) : ''
  }
  for (const filter of xmlElements(xml, 'autoFilter')) {
    xml = xml.replace(filter, () => relocate(filter))
  }
  for (const sort of xmlElements(xml, 'sortState')) {
    let placed = relocate(sort)
    if (placed) {
      const conditions = xmlElements(sort, 'sortCondition')
      let remaining = conditions.length
      for (const condition of conditions) {
        const mapped = relocate(condition)
        if (!mapped) {
          remaining--
        }
        placed = placed.replace(condition, () => mapped)
      }
      if (conditions.length && !remaining) {
        placed = ''
      }
    }
    xml = xml.replace(sort, () => placed)
  }
  return xml
}

/** Follow only local package relationships; binary payloads remain opaque. */
async function visitSourceParts(zip: JSZip, part: string, visit: (part: string, xml: string) => void, visited = new Set<string>()): Promise<void> {
  if (visited.has(part)) {
    return
  }
  visited.add(part)
  if (/\.(xml|vml)$/.test(part)) {
    const entry = zip.file(part)
    if (entry) {
      visit(part, await entry.async('string'))
    }
  }
  const slash = part.lastIndexOf('/')
  const rels = `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`
  const relations = await zip.file(rels)?.async('string')
  for (const relation of xmlElements(relations ?? '', 'Relationship')) {
    const attr = xmlAttributes(relation)
    if (attr.TargetMode === 'External') {
      continue
    }
    const target = resolvePart(part, attr.Target)
    const entry = zip.file(target)
    if (!entry) {
      continue
    }
    await visitSourceParts(zip, target, visit, visited)
  }
}

/** Native objects reserve geometry even when Excel stores no cell value or style there. */
export async function readSourceContent(zip: JSZip, part: string): Promise<GridRange[]> {
  const ranges = new Map<string, GridRange>()
  const add = (ref: string) => {
    if (!ranges.has(ref)) {
      ranges.set(ref, parseRange(ref))
    }
  }
  await visitSourceParts(zip, part, (_part, xml) => {
    for (const node of xml.matchAll(/<(?:dataValidation|conditionalFormatting|hyperlink|comment)\b[^>]*>/g)) {
      const attr = xmlAttributes(node[0])
      for (const ref of (attr.sqref ?? attr.ref ?? '').split(/\s+/).filter(Boolean)) {
        add(ref)
      }
    }
    for (const from of xml.matchAll(/<(?:\w+:)?from>([\s\S]*?)<\/(?:\w+:)?from>/g)) {
      const row = /<(?:\w+:)?row>(\d+)</.exec(from[1])
      const column = /<(?:\w+:)?col>(\d+)</.exec(from[1])
      if (row && column) {
        add(formatAddress({ row: Number(row[1]) + 1, column: Number(column[1]) + 1 }))
      }
    }
  })
  return [...ranges.values()]
}

/** Follow local package relationships. Unrelated/binary parts pass through without interpretation. */
export async function relocateSourceMetadata(zip: JSZip, maps: ReadonlyMap<string, SourceCoordinates>): Promise<void> {
  const visited = new Set<string>()
  for (const map of maps.values()) {
    if (map.part) {
      await visitSourceParts(zip, map.part, (part, xml) => {
        if (part !== map.part) {
          zip.file(part, relocateMetadata(xml, map, maps))
        }
      }, visited)
    }
  }
}

function relocateMetadata(xml: string, map: SourceCoordinates, maps: ReadonlyMap<string, SourceCoordinates>): string {
  xml = xml.replace(/<comment\b[^>]*>[\s\S]*?<\/comment>/g, node => {
    const ref = xmlAttributes(node.split('>')[0]).ref
    return ref ? map.points(ref).map(ref => setXmlAttributes(node, { ref })).join('') : node
  })
  let shapeId = Math.max(1024, ...[...xml.matchAll(/id="_x0000_s(\d+)"/g)].map(match => Number(match[1])))
  xml = xml.replace(/<v:shape\b[^>]*>[\s\S]*?<\/v:shape>/g, node => {
    const row = /<x:Row>(\d+)<\/x:Row>/.exec(node)
    const column = /<x:Column>(\d+)<\/x:Column>/.exec(node)
    if (!row || !column) {
      return node
    }
    const sourceRow = Number(row[1]) + 1
    const sourceColumn = Number(column[1]) + 1
    return map.points(formatAddress({ row: sourceRow, column: sourceColumn })).map(address => {
      const { row: targetRow, column: targetColumn } = parseAddress(address)
      const rowDelta = targetRow - sourceRow
      const columnDelta = targetColumn - sourceColumn
      return setXmlAttributes(node, { id: `_x0000_s${++shapeId}` })
        .replace(/<x:Row>\d+<\/x:Row>/, `<x:Row>${targetRow - 1}</x:Row>`)
        .replace(/<x:Column>\d+<\/x:Column>/, `<x:Column>${targetColumn - 1}</x:Column>`)
        .replace(/<x:Anchor>([^<]*)<\/x:Anchor>/, (_, text: string) => {
          const values = text.split(',').map(Number)
          for (const index of [0, 4]) {
            values[index] += columnDelta
          }
          for (const index of [2, 6]) {
            values[index] += rowDelta
          }
          return `<x:Anchor>${values.join(', ')}</x:Anchor>`
        })
    }).join('')
  })
  let drawingId = Math.max(0, ...[...xml.matchAll(/<(?:\w+:)?cNvPr\b[^>]*\bid="(\d+)"/g)].map(match => Number(match[1])))
  xml = xml.replace(/<((?:\w+:)?(?:oneCellAnchor|twoCellAnchor))\b[^>]*>[\s\S]*?<\/\1>/g, node => {
    const from = /<(?:\w+:)?from>([\s\S]*?)<\/(?:\w+:)?from>/.exec(node)?.[1]
    if (!from) {
      return node
    }
    const row = Number(/<(?:\w+:)?row>(\d+)</.exec(from)?.[1]) + 1
    const column = Number(/<(?:\w+:)?col>(\d+)</.exec(from)?.[1]) + 1
    const points = map.points(formatAddress({ row, column }))
    return points.map((address, index) => {
      const target = parseAddress(address)
      const rowDelta = target.row - row
      const columnDelta = target.column - column
      let copy = node.replace(/<(?:\w+:)?(?:from|to)>[\s\S]*?<\/(?:\w+:)?(?:from|to)>/g, anchor => anchor
        .replace(/<((?:\w+:)?row)>(\d+)<\/\1>/g, (_, tag: string, value: string) => `<${tag}>${Number(value) + rowDelta}</${tag}>`)
        .replace(/<((?:\w+:)?col)>(\d+)<\/\1>/g, (_, tag: string, value: string) => `<${tag}>${Number(value) + columnDelta}</${tag}>`))
      if (index) {
        copy = copy.replace(/<(?:\w+:)?cNvPr\b[^>]*>/g, tag => setXmlAttributes(tag, { id: ++drawingId, name: `${xmlAttributes(tag).name} (${drawingId})` }))
      }
      return copy
    }).join('')
  })
  xml = relocateFilterRanges(xml, map)
  xml = xml.replace(/<table\b[^>]*>/g, node => {
    const ref = xmlAttributes(node).ref
    return ref ? setXmlAttributes(node, { ref: map.range(ref) ?? '' }) : node
  })
  xml = xml.replace(/<((?:\w+:)?(?:f|formula|calculatedColumnFormula|totalsRowFormula))\b([^>]*)>([^<]*)<\/\1>/g,
    (_, tag: string, attr: string, value: string) => `<${tag}${attr}>${encodeXml(sourceFormula(decodeXml(value), map, maps))}</${tag}>`)
  return xml
}

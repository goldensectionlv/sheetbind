import JSZip from 'jszip'
import type { Workbook } from 'exceljs'
import { prepareWorkbookCells, writeWorkbookStrings } from './workbook-cells'
import type { WorkbookPlan } from '../grid/workbook-layout'
import type { FormulaRows } from '../grid/workbook-formula'
import { sourceCoordinates, sourceFormula } from './source-coordinates'
import type { SourceCoordinates } from './source-coordinates'
import { sourceStyles } from './source-styles'
import { offsetWorkbookStrings, sourceWorksheet } from './source-worksheet'
import { relocateSourceMetadata } from './source-metadata'
import type { WorkbookCells } from './workbook-cells'
import { appendXmlChildren, decodeXml, encodeXml, setXmlAttributes, setXmlElement, xmlAttributes, xmlElements } from './xml'
import { decodeXstring, protect, escapeReportText } from './report-text'
import { workbookParts } from './workbook-resources'

export interface WorkbookSource {
  readonly source: Uint8Array
  readonly plan: WorkbookPlan
  readonly form?: boolean
  readonly carriers?: ReadonlyMap<string, FormulaRows>
}
interface GeneratedWorkbook {
  readonly zip: JSZip
  readonly parts: Awaited<ReturnType<typeof workbookParts>>
  readonly cells: WorkbookCells
}

/** Bind source sheets and their coordinate maps once, before writing any package parts. */
function sourceSheets(source: WorkbookSource, parts: GeneratedWorkbook['parts']) {
  const originals = new Map([...parts.values()].map(part => [part.part, part]))
  return new Map(source.plan.sheets.map(plan => {
    const target = plan.sheet
    const original = originals.get(plan.definition.xlsx!.part)!
    const coordinates = sourceCoordinates(plan, source.carriers?.get(target.name.toLowerCase()))
    return [target.name, { original, output: target, coordinates }] as const
  }))
}

/** Overlay rendered sheets; strings are committed by the package writer after all cells are emitted. */
async function preserveWorkbookSource(zip: JSZip, generated: GeneratedWorkbook, source: WorkbookSource): Promise<void> {
  const bindings = sourceSheets(source, await workbookParts(zip))
  const maps = new Map([...bindings.values()].map(binding => [binding.original.name.toLowerCase(), binding.coordinates]))
  const styles = await sourceStyles(zip, generated.zip)
  let workbook = await zip.file('xl/workbook.xml')!.async('string')
  let relations = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
  let types = await zip.file('[Content_Types].xml')!.async('string')
  let sheetId = Math.max(0, ...xmlElements(workbook, 'sheet').map(node => Number(xmlAttributes(node).sheetId)))
  let partId = Math.max(0, ...Object.keys(zip.files).flatMap(name => /^xl\/worksheets\/sheet(\d+)\.xml$/.exec(name)?.slice(1).map(Number) ?? []))
  for (const [name, part] of generated.parts) {
    let fresh = offsetWorkbookStrings(await generated.zip.file(part.part)!.async('string'), generated.cells.stringOffset)
    const binding = bindings.get(name)
    if (binding) {
      const original = await zip.file(binding.original.part)!.async('string')
      const xml = sourceWorksheet(original, fresh, { output: binding.output, map: binding.coordinates, maps, styles, writer: generated.cells, form: source.form })
      zip.file(binding.original.part, new TextEncoder().encode(xml))
      continue
    }
    fresh = fresh.replace(/<c\b[^>]*>/g, node => setXmlAttributes(node, { s: styles.generated(Number(xmlAttributes(node).s ?? 0)) }))
    const target = `xl/worksheets/sheet${++partId}.xml`
    let relationId = `sheetbind${partId}`
    while (relations.includes(`Id="${relationId}"`)) {
      relationId += '_'
    }
    workbook = appendXmlChildren(workbook, 'sheets', [setXmlAttributes(part.node, { sheetId: ++sheetId, 'r:id': relationId })])
    relations = appendXmlChildren(relations, 'Relationships', [`<Relationship Id="${relationId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${partId}.xml"/>`])
    types = appendXmlChildren(types, 'Types', [`<Override PartName="/${target}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`])
    zip.file(target, new TextEncoder().encode(fresh))
  }
  const freshWorkbook = await generated.zip.file('xl/workbook.xml')!.async('string')
  workbook = mergeWorkbookNames(workbook, freshWorkbook, maps)
  workbook = setXmlElement(workbook, 'calcPr', '<calcPr fullCalcOnLoad="1"/>')
  const names = xmlElements(workbook, 'definedNames')[0]
  if (names) {
    workbook = workbook.replace(names, '').replace(/<calcPr\b/, () => names + '<calcPr')
  }
  zip.remove('xl/calcChain.xml')
  relations = relations.replace(/<Relationship\b[^>]*Type="[^"]*\/calcChain"[^>]*\/>/g, '')
  types = types.replace(/<Override\b[^>]*PartName="\/xl\/calcChain.xml"[^>]*\/>/g, '')
  zip.file('xl/workbook.xml', workbook)
  zip.file('xl/_rels/workbook.xml.rels', relations)
  zip.file('[Content_Types].xml', types)
  styles.save()
  await relocateSourceMetadata(zip, maps)
}

function mergeWorkbookNames(workbook: string, generated: string, maps: ReadonlyMap<string, SourceCoordinates>): string {
  const sheetNames = xmlElements(workbook, 'sheet').map(node => decodeXstring(xmlAttributes(node).name))
  const generatedNames = xmlElements(generated, 'sheet').map(node => decodeXstring(xmlAttributes(node).name))
  const localMaps = new Map([...maps.values()].map(map => [map.name.toLowerCase(), map]))
  const names: { name: string, scope?: string, xml: string }[] = xmlElements(workbook, 'definedName').map(node => {
    const attributes = xmlAttributes(node.split('>')[0])
    const sheetName = sheetNames[Number(attributes.localSheetId)]
    const map = localMaps.get((sheetName ?? '').toLowerCase()) ?? maps.values().next().value
    const xml = map
      ? node.replace(/>([^<]*)<\/definedName>$/, (_, value: string) => `>${encodeXml(protect(sourceFormula(decodeXstring(decodeXml(value)), map, maps)))}</definedName>`)
      : node
    return { name: decodeXstring(attributes.name).toLowerCase(), scope: attributes.localSheetId, xml }
  })
  for (const node of xmlElements(generated, 'definedName')) {
    const attributes = xmlAttributes(node.split('>')[0])
    const local = attributes.localSheetId === undefined ? undefined : generatedNames[Number(attributes.localSheetId)]
    const targetIndex = local === undefined ? undefined : sheetNames.indexOf(local)
    let xml = setXmlAttributes(node, { localSheetId: targetIndex })
    if (local !== undefined && ['_xlnm.Print_Area', '_xlnm.Print_Titles'].includes(attributes.name)) {
      // ExcelJS writes this sheet prefix without escaping its internal apostrophes.
      // Coordinates are already final; repair only the prefix using the sheet identity.
      xml = xml.replace(/>([^<]*)<\/definedName>$/, (_, value: string) =>
        `>${encodeXml(protect(decodeXstring(decodeXml(value)).split(`'${local}'!`).join(`'${local.replace(/'/g, "''")}'!`)))}</definedName>`)
    }
    const scope = targetIndex === undefined ? undefined : String(targetIndex)
    const name = decodeXstring(attributes.name).toLowerCase()
    const identity = names.findIndex(entry => entry.name === name && entry.scope === scope)
    const entry = { name, scope, xml }
    if (identity < 0) {
      names.push(entry)
    }
    else {
      names[identity] = entry
    }
  }
  return setXmlElement(workbook, 'definedNames', names.length ? `<definedNames>${names.map(entry => entry.xml).join('')}</definedNames>` : '')
}

/** Apply resolved values and geometry to the source before the only compression pass. */
export async function writeWorkbookPackage(workbook: Workbook, source: WorkbookSource): Promise<Buffer> {
  const generated = await JSZip.loadAsync(await workbook.xlsx.writeBuffer({ zip: { compression: 'STORE' } }))
  await escapeReportText(generated)
  const original = await JSZip.loadAsync(source.source)
  const inheritedStrings = xmlElements(await original.file('xl/sharedStrings.xml')?.async('string') ?? '', 'si')
  const parts = await workbookParts(generated)
  const cells = await prepareWorkbookCells(generated, inheritedStrings)
  await preserveWorkbookSource(original, { zip: generated, parts, cells }, source)
  await writeWorkbookStrings(original, cells.strings)
  return original.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

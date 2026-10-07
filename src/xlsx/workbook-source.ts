import JSZip from 'jszip'
import type { Dictionaries } from '../core/dictionaries'
import type { WorkbookPlan } from '../grid/workbook-layout'
import type { FormulaRows } from '../grid/workbook-formula'
import { prepareWorkbookCells } from './workbook-cells'
import { sourceCoordinates, sourceFormula } from './source-coordinates'
import type { SourceCoordinates } from './source-coordinates'
import { sourceWorksheet } from './source-worksheet'
import { relocateSourceMetadata } from './source-metadata'
import { appendXmlChildren, decodeXml, encodeXml, setXmlElement, xmlAttributes, xmlElements } from './xml'
import { decodeXstring, protect } from './report-text'
import { createWorkbookResources, workbookParts } from './workbook-resources'
import type { WorkbookResourceSource } from './workbook-resources'
import { workbookLists } from './workbook-lists'
import type { FormMarkers } from './workbook-form-markers'
import { workbookPrintNames } from './workbook-print'

interface WorkbookSource {
  readonly source: Uint8Array
  readonly plan: WorkbookPlan
  readonly dictionaries: Dictionaries
  readonly resources: WorkbookResourceSource
  readonly form?: {
    readonly markers: ReadonlyMap<string, FormMarkers>
    readonly rows: ReadonlyMap<string, FormulaRows>
    readonly choices?: string
  }
}

/** Write the plan into its source package; no intermediate workbook or package is built. */
export async function writeWorkbookPackage(source: WorkbookSource): Promise<Buffer> {
  const zip = await JSZip.loadAsync(source.source)
  const parts = await workbookParts(zip)
  const originals = new Map([...parts.values()].map(part => [part.part, part]))
  const bindings = source.plan.sheets.map(plan => ({
    original: originals.get(plan.definition.xlsx!.part)!,
    output: plan.sheet,
    coordinates: sourceCoordinates(plan, source.form?.rows.get(plan.sheet.name.toLowerCase())),
  }))
  const maps = new Map(bindings.map(binding => [binding.original.name.toLowerCase(), binding.coordinates]))
  const cells = await prepareWorkbookCells(zip)
  const lists = workbookLists(source.plan, source.dictionaries, createWorkbookResources(source.resources), cells, source.form?.choices)
  for (const { original, output, coordinates } of bindings) {
    const xml = sourceWorksheet(await zip.file(original.part)!.async('string'), {
      output, map: coordinates, maps, writer: cells, form: !!source.form,
      markers: source.form?.markers.get(output.name), validations: lists.validations.get(output.name)!,
    })
    zip.file(original.part, new TextEncoder().encode(xml))
  }
  let workbook = await zip.file('xl/workbook.xml')!.async('string')
  let relations = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
  let types = await zip.file('[Content_Types].xml')!.async('string')
  if (lists.xml) {
    const sheetId = Math.max(0, ...xmlElements(workbook, 'sheet').map(node => Number(xmlAttributes(node).sheetId))) + 1
    const partId = Math.max(0, ...Object.keys(zip.files).flatMap(name => /^xl\/worksheets\/sheet(\d+)\.xml$/.exec(name)?.slice(1).map(Number) ?? [])) + 1
    const target = `xl/worksheets/sheet${partId}.xml`
    let relationId = `sheetbind${partId}`
    while (relations.includes(`Id="${relationId}"`)) {
      relationId += '_'
    }
    workbook = appendXmlChildren(workbook, 'sheets', [`<sheet name="${lists.name}" sheetId="${sheetId}" state="veryHidden" r:id="${relationId}"/>`])
    relations = appendXmlChildren(relations, 'Relationships', [`<Relationship Id="${relationId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${partId}.xml"/>`])
    types = appendXmlChildren(types, 'Types', [`<Override PartName="/${target}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`])
    zip.file(target, new TextEncoder().encode(lists.xml))
  }
  const sheetNames = [...parts.keys()]
  const generatedNames = [...source.plan.sheets.flatMap(({ sheet }) => workbookPrintNames(sheet.name, sheetNames.indexOf(sheet.name), sheet.print)), ...lists.names]
  workbook = mergeWorkbookNames(workbook, generatedNames, maps)
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
  await cells.save()
  await relocateSourceMetadata(zip, maps)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

function mergeWorkbookNames(workbook: string, generated: readonly string[], maps: ReadonlyMap<string, SourceCoordinates>): string {
  const sheetNames = xmlElements(workbook, 'sheet').map(node => decodeXstring(xmlAttributes(node).name))
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
  for (const xml of generated) {
    const attributes = xmlAttributes(xml.split('>')[0])
    const name = decodeXstring(attributes.name).toLowerCase()
    const scope = attributes.localSheetId
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

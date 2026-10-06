import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { parseAddress, parseRange, XLSX_MAX_ROW } from './addresses'
import { decodeXstring } from './report-text'
import { workbookParts } from './workbook-resources'
import { xmlAttributes, xmlElements } from './xml'

/** One XLSX decoder for authored templates and returned forms. Source bytes stay untouched. */
export async function loadWorkbook(bytes: Uint8Array, source?: JSZip): Promise<ExcelJS.Workbook> {
  const zip = source ?? await JSZip.loadAsync(bytes)
  let changed = false
  const emptyFormulaResults = new Map<string, string[]>()
  for (const [path, part] of Object.entries(zip.files)) {
    if (part.dir || path !== 'xl/sharedStrings.xml' && !/^xl\/worksheets\/sheet\d+\.xml$/.test(path)) {
      continue
    }
    const xml = await part.async('string')
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(path)) {
      for (const row of xml.matchAll(/<row\b[^>]*\br=["'](\d+)["']/g)) {
        if (Number(row[1]) > XLSX_MAX_ROW) {
          throw new RangeError('Row exceeds the XLSX format')
        }
      }
      for (const cell of xml.matchAll(/<c\b[^>]*\br=["']([^"']+)["']/g)) {
        parseAddress(cell[1])
      }
      for (const merge of xml.matchAll(/<mergeCell\b[^>]*\bref=["']([^"']+)["']/g)) {
        parseRange(merge[1])
      }
      // ExcelJS loses an explicitly cached empty string; a missing numeric cache stays an error.
      const empty = xmlElements(xml, 'c').flatMap(node => {
        const attr = xmlAttributes(node.split('>')[0])
        return attr.t === 'str' && /<f\b/.test(node) && /<v\s*\/>|<v\s*><\/v>/.test(node) ? [attr.r] : []
      })
      if (empty.length) {
        emptyFormulaResults.set(path, empty)
      }
    }
    const normalized = normalizeFormText(xml)
    if (normalized !== xml) {
      changed = true
      zip.file(path, normalized)
    }
  }
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(Uint8Array.from(changed ? await zip.generateAsync({ type: 'uint8array' }) : bytes).buffer)
  // These ST_Xstring attributes are not decoded by ExcelJS at all.
  workbook.definedNames.model = workbook.definedNames.model.map(name => ({ name: decodeXstring(name.name), ranges: name.ranges.map(decodeXstring) }))
  const styles = new WeakSet<object>()
  function decodeStyle(style: Partial<ExcelJS.Style> | undefined) {
    if (style && !styles.has(style)) {
      styles.add(style)
      if (style.numFmt) {
        style.numFmt = decodeXstring(style.numFmt)
      }
    }
  }
  for (const sheet of workbook.worksheets) {
    sheet.name = decodeXstring(sheet.name)
    sheet.columns?.forEach(column => decodeStyle(column.style))
    sheet.eachRow({ includeEmpty: true }, row => {
      decodeStyle(row.model?.style)
      row.eachCell({ includeEmpty: true }, cell => decodeStyle(cell.style))
    })
  }
  if (emptyFormulaResults.size) {
    for (const [name, part] of await workbookParts(zip)) {
      for (const address of emptyFormulaResults.get(part.part) ?? []) {
        const cell = workbook.getWorksheet(name)!.getCell(address)
        cell.value = { ...cell.value as ExcelJS.CellFormulaValue, result: '' }
      }
    }
  }
  return workbook
}

// ExcelJS 4.4 recognizes only uppercase hex digits. Normalize encoded text
// without recursively decoding protected literals.
function normalizeFormText(xml: string): string {
  return xml.replace(/(<t(?:\s[^>]*)?>)([\s\S]*?)(<\/t>)/g, (_, start, text, end) =>
    start + text.replace(/_x([0-9a-fA-F]{4})_/g, (_: string, hex: string) => `_x${hex.toUpperCase()}_`) + end)
}

import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { parseAddress, parseRange, XLSX_MAX_ROW } from './addresses'

export async function loadFormWorkbook(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const zip = await JSZip.loadAsync(bytes)
  let changed = false
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
  workbook.definedNames.model = workbook.definedNames.model.map(name => ({ name: name.name, ranges: name.ranges.map(decodeXstring) }))
  for (const sheet of workbook.worksheets) {
    sheet.name = decodeXstring(sheet.name)
  }
  return workbook
}

// ExcelJS 4.4 recognizes only uppercase hex digits. Normalize encoded text
// without recursively decoding protected literals.
function normalizeFormText(xml: string): string {
  return xml.replace(/(<t(?:\s[^>]*)?>)([\s\S]*?)(<\/t>)/g, (_, start, text, end) =>
    start + text.replace(/_x([0-9a-fA-F]{4})_/g, (_: string, hex: string) => `_x${hex.toUpperCase()}_`) + end)
}

function decodeXstring(value: string): string {
  return value.replace(/_x([0-9a-fA-F]{4})_/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
}

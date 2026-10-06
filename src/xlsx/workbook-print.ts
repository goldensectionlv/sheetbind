import type ExcelJS from 'exceljs'
import type JSZip from 'jszip'
import type { WorkbookPrint } from '../grid/workbook-print'
import { formatAddress, parseRange } from './addresses'
import { decodeXml, xmlAttributes } from './xml'
import { decodeXstring } from './report-text'

/** Read coordinate-bearing print settings; page formatting stays in the source XLSX. */
export async function readWorkbookPrint(zip: JSZip, book: ExcelJS.Workbook): Promise<Map<string, WorkbookPrint>> {
  const settings: { area?: WorkbookPrint['area'], repeatRows?: WorkbookPrint['repeatRows'] }[] = book.worksheets.map(() => ({}))
  const xml = await zip.file('xl/workbook.xml')!.async('string')
  for (const match of xml.matchAll(/<definedName\b([^>]*)>([^<]*)<\/definedName>/g)) {
    const attr = xmlAttributes(match[1])
    const index = Number(attr.localSheetId)
    const sheet = book.worksheets[index]
    if (!/^\d+$/.test(attr.localSheetId ?? '') || !sheet || !['_xlnm.Print_Area', '_xlnm.Print_Titles'].includes(attr.name)) {
      continue
    }
    const value = decodeXstring(decodeXml(match[2]))
    const prefix = value.match(/^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/)
    if (!prefix || (prefix[1]?.replace(/''/g, "'") ?? prefix[2]) !== sheet.name) {
      continue
    }
    if (attr.name === '_xlnm.Print_Area' && /^\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$/i.test(prefix[3])) {
      settings[index].area = parseRange(prefix[3])
    }
    if (attr.name === '_xlnm.Print_Titles') {
      const rows = /^\$?([1-9]\d*):\$?([1-9]\d*)$/.exec(prefix[3])
      if (rows) {
        settings[index].repeatRows = { start: Number(rows[1]), end: Number(rows[2]) }
      }
    }
  }
  return new Map(book.worksheets.flatMap((sheet, index) => Object.keys(settings[index]).length ? [[sheet.name, settings[index]] as const] : []))
}

export function writeWorkbookPrint(sheet: ExcelJS.Worksheet, print: WorkbookPrint | undefined): void {
  if (print?.area) {
    sheet.pageSetup.printArea = `${formatAddress(print.area.start)}:${formatAddress(print.area.end)}`
  }
  if (print?.repeatRows) {
    sheet.pageSetup.printTitlesRow = `${print.repeatRows.start}:${print.repeatRows.end}`
  }
}

import JSZip from 'jszip'
import type { Workbook } from 'exceljs'
import { escapeReportText } from './report-text'
import { preserveWorkbookSource } from './workbook-source'
import { workbookParts } from './workbook-resources'
import type { WorkbookSource } from './workbook-source'
import { prepareWorkbookCells, writeWorkbookStrings } from './workbook-cells'
import { xmlElements } from './xml'

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

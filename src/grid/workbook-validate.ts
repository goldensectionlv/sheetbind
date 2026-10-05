import { intersects, regionBounds, workbookIssue, WorkbookAxis, WORKBOOK_LIMITS } from './workbook'
import type { WorkbookBody, WorkbookDefinition, WorkbookRegion } from './workbook'

/** Geometric band compatibility is a template invariant, including empty repeats. */
function assertWorkbookBands(sheet: WorkbookBody): void {
  for (const axis of [WorkbookAxis.Rows, WorkbookAxis.Columns]) {
    const bands = new Map<string, { start: number, length: number }[]>()
    function visit(body: WorkbookBody, offset: number, context: string): void {
      for (const region of body.regions ?? []) {
        const start = offset + (axis === WorkbookAxis.Rows ? region.row : region.column ?? 1)
        const length = axis === WorkbookAxis.Rows ? region.height : region.width ?? WORKBOOK_LIMITS.columns
        if (region.type !== 'repeat' || (region.axis ?? WorkbookAxis.Rows) !== axis) {
          visit(region, start - 1, context)
          continue
        }
        const peers = bands.get(context) ?? []
        if (peers.some(peer => start < peer.start + peer.length && peer.start < start + length && (start !== peer.start || length !== peer.length))) {
          workbookIssue('growth-band-overlap', region.id, `Overlapping ${axis} repeats must reserve the same band`)
        }
        peers.push({ start, length })
        bands.set(context, peers)
        visit(region, 0, `${context}/${start}:${length}`)
      }
    }
    visit(sheet, 0, '')
  }
}

/** Validate ownership and repeat geometry after the XLSX adapter has parsed the tags. */
export function validateWorkbookDefinition(template: WorkbookDefinition): WorkbookDefinition {
  for (const sheet of template.sheets) {
    validateBody(sheet, WORKBOOK_LIMITS.rows, WORKBOOK_LIMITS.columns)
    assertWorkbookBands(sheet)
  }
  return template
}

function validateBody(body: WorkbookBody, height: number, width: number, owner?: WorkbookRegion): void {
  const regions = body.regions ?? []
  const bounds = (region: WorkbookRegion) => regionBounds({ ...region, width: region.width ?? width })
  for (const [index, region] of regions.entries()) {
    if (region.height < 1) {
      workbookIssue('empty-region', region.id, 'A region needs at least one body row')
    }
    const column = region.column ?? 1
    const extent = region.width ?? width
    if (region.row + region.height - 1 > height || column + extent - 1 > width) {
      workbookIssue('region-boundary', region.id, 'Region exceeds its parent body')
    }
    if (regions.slice(0, index).some(other => intersects(bounds(region), bounds(other)))) {
      workbookIssue('region-overlap', region.id, 'Sibling regions cannot overlap; use a child region for nesting')
    }
    validateBody(region, region.height, extent, region)
  }
  const occupied = new Set<number>()
  for (const cell of body.cells) {
    if (cell.at.row + cell.size.rows - 1 > height || cell.at.column + cell.size.columns - 1 > width) {
      workbookIssue('region-boundary', cell.id, owner ? 'Cell or merge crosses its region body' : 'Cell exceeds the XLSX grid')
    }
    if (regions.some(region => intersects(cell, bounds(region)))) {
      workbookIssue('region-boundary', cell.id, 'Cell or merge intersects a region without belonging to its body')
    }
    for (let row = cell.at.row; row < cell.at.row + cell.size.rows; row++) {
      for (let column = cell.at.column; column < cell.at.column + cell.size.columns; column++) {
        const key = row * WORKBOOK_LIMITS.columns + column
        if (occupied.has(key)) {
          workbookIssue('cell-overlap', cell.id, 'Cells or merges overlap')
        }
        occupied.add(key)
      }
    }
  }
}

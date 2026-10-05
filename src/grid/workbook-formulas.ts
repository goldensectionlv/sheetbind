import { WorkbookAxis } from './workbook'
import type { WorkbookCell } from './workbook'
import type { WorkbookPlacedCell, WorkbookPlacement } from './workbook-layout'
import { FormulaEdge, compileWorkbookFormula } from './workbook-formula'
import type { FormulaAddress, FormulaReference, FormulaRows } from './workbook-formula'
import { mapAxis, sharedAxisContext, workbookAxes } from './workbook-axis'
import type { WorkbookAxes } from './workbook-axis'

interface FormulaSource { readonly axes: WorkbookAxes, readonly authored: ReadonlyMap<string, WorkbookCell> }

/** Formulas use the same two axis plans as cells, dimensions and print ranges. */
export function resolveWorkbookFormulas(layout: WorkbookPlacement, sources: ReadonlyMap<string, FormulaSource>, rows?: ReadonlyMap<string, FormulaRows>): WorkbookPlacement {
  const formulas = new Map<string, ReturnType<typeof compileWorkbookFormula>>()
  return { sheets: layout.sheets.map(sheet => ({ ...sheet, cells: sheet.cells.map(cell => {
    if (!('formula' in cell.value)) {
      return cell
    }
    let formula = formulas.get(cell.value.formula)
    if (!formula) {
      formula = compileWorkbookFormula(cell.value.formula)
      formulas.set(cell.value.formula, formula)
    }
    return resolveCellFormula(cell, sheet.name, sources, formula, rows)
  }) })) }
}

function resolveCellFormula(cell: WorkbookPlacedCell, sheetName: string, sources: ReadonlyMap<string, FormulaSource>, formula: ReturnType<typeof compileWorkbookFormula>, rows?: ReadonlyMap<string, FormulaRows>): WorkbookPlacedCell {
  const source = sources.get(sheetName.toLowerCase())!
  const authored = source.authored.get(cell.definitionId)!
  const indexes = new Map(cell.origin.iterations.map(item => [item.nodeId, item.index]))
  const at = { row: mapAxis(source.axes.rows, authored.at.row, FormulaEdge.Cell, indexes)!, column: mapAxis(source.axes.columns, authored.at.column, FormulaEdge.Cell, indexes)! }
  function resolveReference(ref: FormulaReference): FormulaReference | undefined {
    const target = sources.get((ref.sheet ?? sheetName).toLowerCase())
    if (!target) {
      throw new SyntaxError('Formula references an unknown worksheet: ' + ref.sheet)
    }
    const axes = target.axes
    function locate(point: FormulaAddress, edge: FormulaEdge): FormulaAddress | undefined {
      const result = { ...point }
      for (const axis of [WorkbookAxis.Rows, WorkbookAxis.Columns]) {
        const coordinate = workbookAxes[axis].coordinate
        const absolute = axis === WorkbookAxis.Rows ? point.absoluteRow : point.absoluteColumn
        const context = !absolute && target === source ? sharedAxisContext(source.axes[axis], authored.at[coordinate], point[coordinate], indexes) : new Map<string, number>()
        const position = mapAxis(axes[axis], point[coordinate], edge, context)
        if (position === undefined) {
          return undefined
        }
        const base = absolute ? at[coordinate] : mapAxis(source.axes[axis], authored.at[coordinate], FormulaEdge.Cell, context)!
        result[coordinate] = position + at[coordinate] - base
      }
      result.row = rows?.get((ref.sheet ?? sheetName).toLowerCase())?.(result.row, edge) ?? result.row
      return result
    }
    const start = locate(ref.start, ref.end ? FormulaEdge.Start : FormulaEdge.Cell)
    const end = ref.end ? locate(ref.end, FormulaEdge.End) : undefined
    return !start || ref.end && !end ? undefined : { ...ref, start, ...(end ? { end } : {}) }
  }
  return { ...cell, value: { formula: formula(resolveReference) } }
}

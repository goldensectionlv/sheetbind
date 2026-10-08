import { WorkbookAxis } from './workbook'
import type { WorkbookPlacedCell, WorkbookPlan, WorkbookSheetPlan } from './workbook-layout'
import { FormulaEdge, compileWorkbookFormula } from './workbook-formula'
import type { FormulaAddress, FormulaReference, FormulaRows } from './workbook-formula'
import { mapAxis, sharedAxisContext, workbookAxes } from './workbook-axis'

/** Formulas use the same two axis plans as cells, dimensions and print ranges. */
export function resolveWorkbookFormulas(plan: WorkbookPlan, rows?: ReadonlyMap<string, FormulaRows>): WorkbookPlan {
  const sources = new Map(plan.sheets.map(source => [source.sheet.name.toLowerCase(), source]))
  const formulas = new Map<string, ReturnType<typeof compileWorkbookFormula>>()
  return { sheets: plan.sheets.map(source => ({ ...source, sheet: { ...source.sheet, cells: source.sheet.cells.map(cell => {
    if (!('formula' in cell.value)) {
      return cell
    }
    let formula = formulas.get(cell.value.formula)
    if (!formula) {
      formula = compileWorkbookFormula(cell.value.formula)
      formulas.set(cell.value.formula, formula)
    }
    return resolveCellFormula(cell, source, sources, formula, rows)
  }) } })) }
}

function resolveCellFormula(cell: WorkbookPlacedCell, source: WorkbookSheetPlan, sources: ReadonlyMap<string, WorkbookSheetPlan>, formula: ReturnType<typeof compileWorkbookFormula>, rows?: ReadonlyMap<string, FormulaRows>): WorkbookPlacedCell {
  const sheetName = source.sheet.name
  const position = source.authored.get(cell.definitionId)!
  const indexes = new Map(cell.origin.iterations.map(item => [item.nodeId, item.index]))
  const at = { row: mapAxis(source.axes.rows, position.row, FormulaEdge.Cell, indexes)!, column: mapAxis(source.axes.columns, position.column, FormulaEdge.Cell, indexes)! }
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
        const context = !absolute && target === source ? sharedAxisContext(source.axes[axis], position[coordinate], point[coordinate], indexes) : new Map<string, number>()
        const mapped = mapAxis(axes[axis], point[coordinate], edge, context)
        if (mapped === undefined) {
          return undefined
        }
        const base = absolute ? at[coordinate] : mapAxis(source.axes[axis], position[coordinate], FormulaEdge.Cell, context)!
        result[coordinate] = mapped + at[coordinate] - base
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

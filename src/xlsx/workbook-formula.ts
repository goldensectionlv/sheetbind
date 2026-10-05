import type { Cell } from 'exceljs'
import { copyWorkbookFormula, parseWorkbookFormula } from '../grid/workbook-formula'

/** Materialize shared formulas using string-aware reference offsets. */
export function readWorkbookFormula(cell: Cell): string {
  const value = cell.value
  if (!value || typeof value !== 'object' || 'shareType' in value && value.shareType !== 'shared') {
    throw new SyntaxError('Expected an ordinary or shared formula')
  }
  if ('formula' in value) {
    return parseWorkbookFormula(value.formula)
  }
  if ('sharedFormula' in value && typeof value.sharedFormula === 'string') {
    const master = cell.worksheet.getCell(value.sharedFormula)
    const source = master.value
    if (source && typeof source === 'object' && 'formula' in source) {
      return copyWorkbookFormula(parseWorkbookFormula(source.formula), Number(cell.row) - Number(master.row), Number(cell.col) - Number(master.col))
    }
  }
  throw new SyntaxError('Shared formula master is missing')
}

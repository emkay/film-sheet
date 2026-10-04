/**
 * film-sheet/engine — the spreadsheet formula engine on its own. No DOM;
 * runs in browsers, Node, Deno, Bun and workers.
 *
 * ```ts
 * import { Workbook } from '@mk/film-sheet/engine'
 * const wb = new Workbook()
 * const id = wb.sheets[0].id
 * wb.setCell(id, 0, 0, '2')
 * wb.setCell(id, 1, 0, '=A1^10')
 * wb.getValue(id, 1, 0) // 1024
 * ```
 */
export { Workbook, Sheet, Cell, validateWorkbookJSON } from './workbook.js'
export type { Style, CellPatch, SheetJSON, WorkbookJSON, RecalcStats } from './workbook.js'

export { MAX_ROWS, MAX_COLS, colToName, nameToCol, formatAddress, formatRange, parseAddress, parseRangeAddress, normalizeRange } from './address.js'
export type { CellAddress, RangeAddress } from './address.js'

export { FORMATS, formatValue, parseInput, dateToSerial, serialToDate } from './format.js'
export type { ParsedInput } from './format.js'

export { Errors, FormulaError, isError, isRange, isArray, toNumber, toText, toBoolean, toScalar, compareScalars } from './values.js'
export type { ErrorCode, Scalar, Value, RangeValue, ArrayValue } from './values.js'

export { registerFunction, FUNCTION_NAMES } from './functions.js'
export type { CustomFunction, FnContext, Thunk } from './functions.js'

export { shiftFormula, normalizeFormula } from './refs.js'

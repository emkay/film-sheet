/** Spreadsheet value model and the coercion rules formulas use. */

export type ErrorCode =
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#N/A'
  | '#NUM!'
  | '#NULL!'
  | '#CIRC!'
  | '#ERROR!'

export class FormulaError {
  constructor (readonly code: ErrorCode, readonly message = '') {}
  toString (): string {
    return this.code
  }
}

/** Shared instances so the hot path never allocates for common errors. */
export const Errors = {
  DIV0: new FormulaError('#DIV/0!', 'Division by zero'),
  VALUE: new FormulaError('#VALUE!', 'Wrong type of argument'),
  REF: new FormulaError('#REF!', 'Invalid cell reference'),
  NAME: new FormulaError('#NAME?', 'Unknown function or name'),
  NA: new FormulaError('#N/A', 'Value not available'),
  NUM: new FormulaError('#NUM!', 'Invalid numeric value'),
  NULL: new FormulaError('#NULL!', 'Empty intersection'),
  CIRC: new FormulaError('#CIRC!', 'Circular reference'),
  PARSE: new FormulaError('#ERROR!', 'Formula parse error')
} as const

const ERROR_BY_CODE = new Map<string, FormulaError>(Object.values(Errors).map((e) => [e.code, e]))
export function errorFromCode (code: string): FormulaError | undefined {
  return ERROR_BY_CODE.get(code.toUpperCase())
}

/** A scalar cell value. `null` is an empty cell. */
export type Scalar = number | string | boolean | FormulaError | null

/** A rectangular reference that functions can iterate lazily. */
export interface RangeValue {
  readonly kind: 'range'
  readonly sheet: number
  readonly r1: number
  readonly c1: number
  readonly r2: number
  readonly c2: number
  /** Value at an offset inside the range (0-based). */
  get (dr: number, dc: number): Scalar
  /** Visit every non-empty cell, in no particular order (fast path for SUM, COUNT…). */
  forEach (fn: (v: Scalar, r: number, c: number) => void): void
  /** Visit every non-empty cell row by row, left to right. */
  forEachOrdered (fn: (v: Scalar, r: number, c: number) => void): void
}

/** An in-memory 2D array (array constants such as {1,2;3,4}). */
export interface ArrayValue {
  readonly kind: 'array'
  readonly rows: Scalar[][]
}

export type Value = Scalar | RangeValue | ArrayValue

export function isError (v: unknown): v is FormulaError {
  return v instanceof FormulaError
}

export function isRange (v: Value): v is RangeValue {
  return typeof v === 'object' && v !== null && (v as RangeValue).kind === 'range'
}

export function isArray (v: Value): v is ArrayValue {
  return typeof v === 'object' && v !== null && (v as ArrayValue).kind === 'array'
}

const NUMERIC_RE = /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i

/** Parse a string as a number the way a formula would, or return undefined. */
export function parseNumber (s: string): number | undefined {
  const t = s.trim()
  if (t === '') return undefined
  if (NUMERIC_RE.test(t)) return Number(t)
  // Percentages and thousands separators, e.g. "12%" or "1,234.5"
  if (/^[+-]?[\d,]*\.?\d+%$/.test(t)) return Number(t.slice(0, -1).replace(/,/g, '')) / 100
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) return Number(t.replace(/,/g, ''))
  return undefined
}

/** Reduce a range/array to its single value (implicit intersection light). */
export function toScalar (v: Value): Scalar {
  if (isRange(v)) {
    if (v.r1 === v.r2 && v.c1 === v.c2) return v.get(0, 0)
    return Errors.VALUE
  }
  if (isArray(v)) return v.rows[0]?.[0] ?? null
  return v
}

export function toNumber (v: Value): number | FormulaError {
  const s = toScalar(v)
  if (typeof s === 'number') return s
  if (s === null) return 0
  if (typeof s === 'boolean') return s ? 1 : 0
  if (s instanceof FormulaError) return s
  const n = parseNumber(s)
  return n === undefined ? Errors.VALUE : n
}

export function toText (v: Value): string | FormulaError {
  const s = toScalar(v)
  if (s === null) return ''
  if (typeof s === 'string') return s
  if (typeof s === 'boolean') return s ? 'TRUE' : 'FALSE'
  if (typeof s === 'number') return formatGeneral(s)
  return s
}

export function toBoolean (v: Value): boolean | FormulaError {
  const s = toScalar(v)
  if (typeof s === 'boolean') return s
  if (typeof s === 'number') return s !== 0
  if (s === null) return false
  if (s instanceof FormulaError) return s
  const u = s.toUpperCase()
  if (u === 'TRUE') return true
  if (u === 'FALSE') return false
  return Errors.VALUE
}

/** Excel's "General" number rendering: up to ~11 significant digits. */
export function formatGeneral (n: number): string {
  if (!Number.isFinite(n)) return '#NUM!'
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n)
  const abs = Math.abs(n)
  // toPrecision itself switches to exponent form below 1e-6, so go there first.
  if (abs !== 0 && (abs >= 1e11 || abs < 1e-5)) {
    const [mant, exp] = n.toExponential(5).split('e')
    const e = Number(exp)
    return mant.replace(/\.?0+$/, '') + 'E' + (e < 0 ? '-' : '+') + String(Math.abs(e)).padStart(2, '0')
  }
  const s = n.toPrecision(11)
  return s.includes('.') ? s.replace(/\.?0+$/, '') : s
}

/** Compare two scalars with spreadsheet ordering: numbers < text < booleans. */
export function compareScalars (a: Scalar, b: Scalar): number {
  const ta = typeRank(a)
  const tb = typeRank(b)
  if (a === null && b === null) return 0
  // Blank compares equal to 0, "" and FALSE.
  if (a === null) return compareScalars(blankLike(b), b)
  if (b === null) return compareScalars(a, blankLike(a))
  if (ta !== tb) return ta - tb
  if (typeof a === 'number') return a - (b as number)
  if (typeof a === 'string') {
    const x = a.toLowerCase()
    const y = (b as string).toLowerCase()
    return x < y ? -1 : x > y ? 1 : 0
  }
  if (typeof a === 'boolean') return (a ? 1 : 0) - ((b as boolean) ? 1 : 0)
  return 0
}

function blankLike (v: Scalar): Scalar {
  if (typeof v === 'string') return ''
  if (typeof v === 'boolean') return false
  return 0
}

function typeRank (v: Scalar): number {
  if (typeof v === 'number') return 0
  if (typeof v === 'string') return 1
  if (typeof v === 'boolean') return 2
  return 3
}

import { dateToSerial, serialToDate, textFormat } from './format.js'
import {
  compareScalars,
  Errors,
  FormulaError,
  isArray,
  isError,
  isRange,
  parseNumber,
  toBoolean,
  toNumber,
  toScalar,
  toText,
  type RangeValue,
  type Scalar,
  type Value
} from './values.js'

/** What a function can see about where it is being evaluated. */
export interface FnContext {
  row: number
  col: number
  sheetId: number
  /** Build a range on the given sheet (clamped to the grid). */
  makeRange (sheetId: number, r1: number, c1: number, r2: number, c2: number): RangeValue | FormulaError
  /** Resolve an A1 reference string (for INDIRECT). */
  rangeFromText (text: string, sheetId: number): RangeValue | FormulaError
}

export type Thunk = () => Value

export interface FunctionDef {
  /** Eager functions receive evaluated arguments. */
  fn?: (args: Value[], ctx: FnContext) => Value
  /** Lazy functions receive thunks (IF, IFERROR, AND short-circuit…). */
  lazy?: (args: Thunk[], ctx: FnContext) => Value
  min: number
  max: number
  volatile?: boolean
  /** Signature for the editor's hint, e.g. "SUM(number1, [number2], …)". */
  sig: string
  desc: string
}

// ---------------------------------------------------------------------------
// Argument helpers

/**
 * Visit every value in an argument, flattening ranges and arrays. `ordered`
 * walks ranges row by row, for functions whose result depends on order.
 */
function forEachValue (v: Value, fn: (s: Scalar, fromRef: boolean) => void, ordered = false): void {
  if (isRange(v)) {
    if (ordered) v.forEachOrdered((s) => fn(s, true))
    else v.forEach((s) => fn(s, true))
  } else if (isArray(v)) for (const row of v.rows) for (const s of row) fn(s, true)
  else fn(v, false)
}

/**
 * Collect numbers following Excel's aggregate rules: numbers in references
 * count, text and booleans in references are skipped, direct arguments are
 * coerced, and errors propagate.
 */
function collectNumbers (args: Value[], ordered = false): number[] | FormulaError {
  const out: number[] = []
  let err: FormulaError | undefined
  for (const a of args) {
    forEachValue(a, (s, fromRef) => {
      if (err) return
      if (typeof s === 'number') out.push(s)
      else if (isError(s)) err = s
      else if (!fromRef && s !== null) {
        const n = toNumber(s)
        if (isError(n)) err = n
        else out.push(n)
      }
    }, ordered)
    if (err) return err
  }
  return out
}

/** Sum without materialising an array — the hottest aggregate. */
function sumArgs (args: Value[]): number | FormulaError {
  let total = 0
  let err: FormulaError | undefined
  for (const a of args) {
    if (isRange(a)) {
      a.forEach((s) => {
        if (typeof s === 'number') total += s
        else if (err === undefined && isError(s)) err = s
      })
    } else if (isArray(a)) {
      for (const row of a.rows) for (const s of row) {
        if (typeof s === 'number') total += s
        else if (err === undefined && isError(s)) err = s
      }
    } else if (a !== null) {
      const n = toNumber(a)
      if (isError(n)) return n
      total += n
    }
    if (err) return err
  }
  return total
}

function num (v: Value): number | FormulaError {
  return toNumber(v)
}

function int (v: Value): number | FormulaError {
  const n = toNumber(v)
  return isError(n) ? n : Math.trunc(n)
}

function str (v: Value): string | FormulaError {
  return toText(v)
}

function numberResult (n: number): number | FormulaError {
  return Number.isFinite(n) ? n : Errors.NUM
}

/** Wrap a numeric function of fixed arity with coercion and error checks. */
function math (f: (...n: number[]) => number, min = 1, max = min) {
  return {
    min,
    max,
    fn: (args: Value[]): Value => {
      const ns: number[] = []
      for (const a of args) {
        const n = num(a)
        if (isError(n)) return n
        ns.push(n)
      }
      return numberResult(f(...ns))
    }
  }
}

function dims (v: Value): { rows: number; cols: number; get: (r: number, c: number) => Scalar } {
  if (isRange(v)) {
    return { rows: v.r2 - v.r1 + 1, cols: v.c2 - v.c1 + 1, get: (r, c) => v.get(r, c) }
  }
  if (isArray(v)) return { rows: v.rows.length, cols: v.rows[0]?.length ?? 0, get: (r, c) => v.rows[r]?.[c] ?? null }
  const s = v
  return { rows: 1, cols: 1, get: () => s }
}

/** Rows that actually hold data, so whole-column lookups stay cheap. */
function usedRows (v: Value): number {
  if (isRange(v)) return Math.min(v.r2 - v.r1 + 1, Math.max(0, usedRowsOf(v)))
  return dims(v).rows
}
function usedCols (v: Value): number {
  if (isRange(v)) return Math.min(v.c2 - v.c1 + 1, Math.max(0, usedColsOf(v)))
  return dims(v).cols
}
function usedRowsOf (v: RangeValue): number {
  return (v as RangeValue & { usedRows?: () => number }).usedRows?.() ?? v.r2 - v.r1 + 1
}
function usedColsOf (v: RangeValue): number {
  return (v as RangeValue & { usedCols?: () => number }).usedCols?.() ?? v.c2 - v.c1 + 1
}

// ---------------------------------------------------------------------------
// Criteria (COUNTIF, SUMIF…)

export function makeCriteria (crit: Scalar): (v: Scalar) => boolean {
  if (typeof crit === 'number') return (v) => typeof v === 'number' ? v === crit : typeof v === 'string' && parseNumber(v) === crit
  if (typeof crit === 'boolean') return (v) => v === crit
  if (crit === null) return (v) => v === null || v === ''
  if (isError(crit)) return (v) => v === crit
  const m = /^(<=|>=|<>|=|<|>)?(.*)$/s.exec(crit)!
  const op = m[1] ?? '='
  const operand = m[2]
  const n = parseNumber(operand)
  if (n !== undefined) {
    switch (op) {
      // "=5" matches the number 5 and the text "5", like a bare 5 does.
      case '=': return (v) => (typeof v === 'number' ? v === n : typeof v === 'string' && parseNumber(v) === n)
      case '<>': return (v) => !(typeof v === 'number' ? v === n : typeof v === 'string' && parseNumber(v) === n)
      case '<': return (v) => typeof v === 'number' && v < n
      case '>': return (v) => typeof v === 'number' && v > n
      case '<=': return (v) => typeof v === 'number' && v <= n
      case '>=': return (v) => typeof v === 'number' && v >= n
    }
  }
  const upper = operand.toUpperCase()
  if (upper === 'TRUE' || upper === 'FALSE') {
    const b = upper === 'TRUE'
    return op === '<>' ? (v) => v !== b : (v) => v === b
  }
  if (operand === '' && op === '=') return (v) => v === null || v === ''
  if (operand === '' && op === '<>') return (v) => v !== null && v !== ''
  if (op === '=' || op === '<>') {
    const re = wildcardRegex(operand)
    const test = re ? (v: Scalar) => typeof v === 'string' && re.test(v) : (v: Scalar) => typeof v === 'string' && v.toLowerCase() === operand.toLowerCase()
    return op === '=' ? test : (v) => !test(v)
  }
  const lower = operand.toLowerCase()
  return (v) => {
    if (typeof v !== 'string') return false
    const c = v.toLowerCase() < lower ? -1 : v.toLowerCase() > lower ? 1 : 0
    return op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0
  }
}

function wildcardRegex (pattern: string): RegExp | undefined {
  if (!/[*?]/.test(pattern)) return undefined
  let src = ''
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]
    if (ch === '~' && i + 1 < pattern.length) src += escapeRe(pattern[++i])
    else if (ch === '*') src += '.*'
    else if (ch === '?') src += '.'
    else src += escapeRe(ch)
  }
  return new RegExp('^' + src + '$', 'is')
}

function escapeRe (s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/**
 * Shared engine for the *IFS family: walk criteria ranges in lockstep and
 * call `hit` with the offset of each row/col that satisfies every criterion.
 *
 * Only the area that holds data (in any criteria range or the `target`) is
 * scanned. Beyond it every cell is blank, so when all criteria accept blanks
 * those cells are counted in `blankHits` rather than visited one by one —
 * `COUNTIF(A:A, "<>5")` costs the used rows, not a million.
 */
function ifsScan (pairs: Value[], hit: (dr: number, dc: number) => void, target?: Value): { err?: FormulaError; blankHits: number } {
  if (pairs.length === 0 || pairs.length % 2 !== 0) return { err: Errors.VALUE, blankHits: 0 }
  const ranges = []
  for (let i = 0; i < pairs.length; i += 2) {
    const d = dims(pairs[i])
    const crit = toScalar(pairs[i + 1])
    ranges.push({ d, test: makeCriteria(crit), v: pairs[i] })
  }
  const first = ranges[0]
  for (const r of ranges) if (r.d.rows !== first.d.rows || r.d.cols !== first.d.cols) return { err: Errors.VALUE, blankHits: 0 }
  if (target !== undefined) {
    const t = dims(target)
    if (t.rows !== first.d.rows || t.cols !== first.d.cols) return { err: Errors.VALUE, blankHits: 0 }
  }
  const all = target !== undefined ? [...ranges.map((r) => r.v), target] : ranges.map((r) => r.v)
  const R = Math.min(first.d.rows, Math.max(...all.map(usedRows)))
  const C = Math.min(first.d.cols, Math.max(...all.map(usedCols)))
  for (let r = 0; r < R; r++) {
    for (let c = 0; c < C; c++) {
      let ok = true
      for (const x of ranges) {
        if (!x.test(x.d.get(r, c))) {
          ok = false
          break
        }
      }
      if (ok) hit(r, c)
    }
  }
  const allMatchBlank = ranges.every((r) => r.test(null))
  return { blankHits: allMatchBlank ? first.d.rows * first.d.cols - R * C : 0 }
}

/** Call `fn` with each number in `target` whose row/col meets every criterion. */
function eachMatch (target: Value, pairs: Value[], fn: (v: number) => void): FormulaError | undefined {
  const t = dims(target)
  return ifsScan(pairs, (r, c) => {
    const v = t.get(r, c)
    if (typeof v === 'number') fn(v)
  }, target).err
}

// ---------------------------------------------------------------------------
// Lookup helpers

function lookupMatch (needle: Scalar, get: (i: number) => Scalar, n: number, mode: number, wildcards = true): number {
  if (needle === null) return -1
  if (mode === 0) {
    if (typeof needle === 'string') {
      const re = wildcards ? wildcardRegex(needle) : undefined
      const lower = needle.toLowerCase()
      for (let i = 0; i < n; i++) {
        const v = get(i)
        if (typeof v === 'string' && (re ? re.test(v) : v.toLowerCase() === lower)) return i
      }
      return -1
    }
    for (let i = 0; i < n; i++) {
      const v = get(i)
      if (v === needle || (v !== null && needle !== null && compareScalars(v, needle) === 0 && typeof v === typeof needle)) return i
    }
    return -1
  }
  // Binary search on sorted data (mode 1: largest ≤ needle; -1: smallest ≥ needle, descending).
  let lo = 0
  let hi = n - 1
  let best = -1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    const v = get(mid)
    const c = v === null ? 1 : compareScalars(v, needle)
    if (mode === 1 ? c <= 0 : c >= 0) {
      best = mid
      lo = mid + 1
    } else hi = mid - 1
  }
  return best
}

// ---------------------------------------------------------------------------
// Statistics helpers

function mean (ns: number[]): number {
  let s = 0
  for (const n of ns) s += n
  return s / ns.length
}

function variance (ns: number[], sample: boolean): number | FormulaError {
  if (ns.length < (sample ? 2 : 1)) return Errors.DIV0
  const m = mean(ns)
  let s = 0
  for (const n of ns) s += (n - m) * (n - m)
  return s / (ns.length - (sample ? 1 : 0))
}

/** `aggregate(f)`, or `aggregate(true, f)` when f needs values in sheet order. */
function aggregate (f: ((ns: number[]) => number | FormulaError) | true, g?: (ns: number[]) => number | FormulaError) {
  const ordered = f === true
  const fn = f === true ? g! : f
  return (args: Value[]): Value => {
    const ns = collectNumbers(args, ordered)
    return isError(ns) ? ns : fn(ns)
  }
}

function roundTo (n: number, digits: number, mode: 'round' | 'up' | 'down'): number {
  // Doubles carry ~15–17 significant digits; rounding finer than that is a no-op
  // (and 10^digits would overflow for huge numbers).
  if (digits > 15) return n
  if (digits < -308) return 0
  const f = Math.pow(10, digits)
  if (!Number.isFinite(n * f)) return n
  const x = n * f
  // Nudge by epsilon so 2.675 rounds the way people expect, not the way IEEE stores it.
  const eps = Number.EPSILON * Math.abs(x) * 4
  let r: number
  if (mode === 'round') r = Math.sign(x) * Math.round(Math.abs(x) + eps)
  else if (mode === 'up') r = Math.sign(x) * Math.ceil(Math.abs(x) - eps)
  else r = Math.sign(x) * Math.floor(Math.abs(x) + eps)
  return r / f
}

function truthy (v: Value): boolean | FormulaError {
  return toBoolean(v)
}

function logical (args: Value[], combine: 'and' | 'or' | 'xor'): Value {
  let acc = combine === 'and'
  let count = 0
  let seen = false
  for (const a of args) {
    let err: FormulaError | undefined
    forEachValue(a, (s, fromRef) => {
      if (err) return
      if (fromRef && (s === null || typeof s === 'string')) return
      const b = toBoolean(s)
      if (isError(b)) {
        err = b
        return
      }
      seen = true
      if (combine === 'and') acc = acc && b
      else if (combine === 'or') acc = acc || b
      else if (b) count++
    })
    if (err) return err
  }
  if (!seen) return Errors.VALUE
  return combine === 'xor' ? count % 2 === 1 : acc
}

function today (): number {
  const d = new Date()
  return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate())
}

function now (): number {
  const d = new Date()
  return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds())
}

function datePart (v: Value, part: (d: Date) => number): Value {
  const n = num(v)
  if (isError(n)) return n
  if (n < 0) return Errors.NUM
  return part(serialToDate(n))
}

// ---------------------------------------------------------------------------
// The function table

export const FUNCTIONS: Record<string, FunctionDef> = {
  // Math & aggregates ------------------------------------------------------
  SUM: { min: 1, max: 255, fn: sumArgs, sig: 'SUM(number1, [number2], …)', desc: 'Adds all the numbers' },
  PRODUCT: { min: 1, max: 255, fn: aggregate((ns) => (ns.length ? ns.reduce((a, b) => a * b, 1) : 0)), sig: 'PRODUCT(number1, [number2], …)', desc: 'Multiplies all the numbers' },
  AVERAGE: { min: 1, max: 255, fn: aggregate((ns) => (ns.length ? mean(ns) : Errors.DIV0)), sig: 'AVERAGE(number1, [number2], …)', desc: 'Arithmetic mean of the numbers' },
  MIN: { min: 1, max: 255, fn: aggregate((ns) => (ns.length ? ns.reduce((a, b) => (b < a ? b : a)) : 0)), sig: 'MIN(number1, [number2], …)', desc: 'Smallest number' },
  MAX: { min: 1, max: 255, fn: aggregate((ns) => (ns.length ? ns.reduce((a, b) => (b > a ? b : a)) : 0)), sig: 'MAX(number1, [number2], …)', desc: 'Largest number' },
  MEDIAN: {
    min: 1,
    max: 255,
    fn: aggregate((ns) => {
      if (!ns.length) return Errors.NUM
      const s = [...ns].sort((a, b) => a - b)
      const m = s.length >> 1
      return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2
    }),
    sig: 'MEDIAN(number1, [number2], …)',
    desc: 'Middle value of the numbers'
  },
  MODE: {
    min: 1,
    max: 255,
    fn: aggregate(true, (ns) => {
      const counts = new Map<number, number>()
      let best = 0
      let bestN: number | undefined
      for (const n of ns) {
        const c = (counts.get(n) ?? 0) + 1
        counts.set(n, c)
        if (c > best) {
          best = c
          bestN = n
        }
      }
      return best > 1 && bestN !== undefined ? bestN : Errors.NA
    }),
    sig: 'MODE(number1, [number2], …)',
    desc: 'Most frequently occurring number'
  },
  COUNT: {
    min: 1,
    max: 255,
    fn: (args) => {
      let n = 0
      for (const a of args) forEachValue(a, (s, fromRef) => {
        if (typeof s === 'number') n++
        else if (!fromRef && (typeof s === 'boolean' || (typeof s === 'string' && parseNumber(s) !== undefined))) n++
      })
      return n
    },
    sig: 'COUNT(value1, [value2], …)',
    desc: 'Counts cells containing numbers'
  },
  COUNTA: {
    min: 1,
    max: 255,
    fn: (args) => {
      let n = 0
      for (const a of args) forEachValue(a, (s) => {
        if (s !== null) n++
      })
      return n
    },
    sig: 'COUNTA(value1, [value2], …)',
    desc: 'Counts non-empty cells'
  },
  COUNTBLANK: {
    min: 1,
    max: 1,
    fn: ([r]) => {
      const d = dims(r)
      let filled = 0
      forEachValue(r, (s) => {
        if (s !== null && s !== '') filled++
      })
      return d.rows * d.cols - filled
    },
    sig: 'COUNTBLANK(range)',
    desc: 'Counts empty cells'
  },
  ABS: { ...math(Math.abs), sig: 'ABS(number)', desc: 'Absolute value' },
  SQRT: { ...math((n) => (n < 0 ? NaN : Math.sqrt(n))), sig: 'SQRT(number)', desc: 'Square root' },
  EXP: { ...math(Math.exp), sig: 'EXP(number)', desc: 'e raised to a power' },
  LN: { ...math((n) => (n <= 0 ? NaN : Math.log(n))), sig: 'LN(number)', desc: 'Natural logarithm' },
  LOG10: { ...math((n) => (n <= 0 ? NaN : Math.log10(n))), sig: 'LOG10(number)', desc: 'Base-10 logarithm' },
  LOG: { ...math((n, b = 10) => (n <= 0 || b <= 0 || b === 1 ? NaN : Math.log(n) / Math.log(b)), 1, 2), sig: 'LOG(number, [base])', desc: 'Logarithm to a base' },
  POWER: { ...math((a, b) => Math.pow(a, b), 2), sig: 'POWER(number, power)', desc: 'Number raised to a power' },
  PI: { min: 0, max: 0, fn: () => Math.PI, sig: 'PI()', desc: 'The value of π' },
  SIN: { ...math(Math.sin), sig: 'SIN(angle)', desc: 'Sine' },
  COS: { ...math(Math.cos), sig: 'COS(angle)', desc: 'Cosine' },
  TAN: { ...math(Math.tan), sig: 'TAN(angle)', desc: 'Tangent' },
  ASIN: { ...math(Math.asin), sig: 'ASIN(number)', desc: 'Arcsine' },
  ACOS: { ...math(Math.acos), sig: 'ACOS(number)', desc: 'Arccosine' },
  ATAN: { ...math(Math.atan), sig: 'ATAN(number)', desc: 'Arctangent' },
  ATAN2: { ...math((x, y) => Math.atan2(y, x), 2), sig: 'ATAN2(x, y)', desc: 'Arctangent from coordinates' },
  DEGREES: { ...math((r) => (r * 180) / Math.PI), sig: 'DEGREES(radians)', desc: 'Radians to degrees' },
  RADIANS: { ...math((d) => (d * Math.PI) / 180), sig: 'RADIANS(degrees)', desc: 'Degrees to radians' },
  SIGN: { ...math(Math.sign), sig: 'SIGN(number)', desc: 'Sign of a number' },
  INT: { ...math(Math.floor), sig: 'INT(number)', desc: 'Rounds down to an integer' },
  TRUNC: { ...math((n, d = 0) => roundTo(n, d, 'down'), 1, 2), sig: 'TRUNC(number, [digits])', desc: 'Truncates a number' },
  ROUND: { ...math((n, d = 0) => roundTo(n, Math.trunc(d), 'round'), 1, 2), sig: 'ROUND(number, [digits])', desc: 'Rounds to a number of digits' },
  ROUNDUP: { ...math((n, d = 0) => roundTo(n, Math.trunc(d), 'up'), 1, 2), sig: 'ROUNDUP(number, [digits])', desc: 'Rounds away from zero' },
  ROUNDDOWN: { ...math((n, d = 0) => roundTo(n, Math.trunc(d), 'down'), 1, 2), sig: 'ROUNDDOWN(number, [digits])', desc: 'Rounds toward zero' },
  CEILING: { ...math((n, s = 1) => (s === 0 ? 0 : Math.ceil(n / s) * s), 1, 2), sig: 'CEILING(number, [significance])', desc: 'Rounds up to a multiple' },
  FLOOR: { ...math((n, s = 1) => (s === 0 ? NaN : Math.floor(n / s) * s), 1, 2), sig: 'FLOOR(number, [significance])', desc: 'Rounds down to a multiple' },
  MROUND: { ...math((n, m) => (m === 0 ? 0 : Math.round(n / m) * m), 2), sig: 'MROUND(number, multiple)', desc: 'Rounds to a multiple' },
  MOD: {
    min: 2,
    max: 2,
    fn: ([a, b]) => {
      const x = num(a)
      const y = num(b)
      if (isError(x)) return x
      if (isError(y)) return y
      if (y === 0) return Errors.DIV0
      return x - y * Math.floor(x / y)
    },
    sig: 'MOD(number, divisor)',
    desc: 'Remainder after division'
  },
  QUOTIENT: { ...math((a, b) => (b === 0 ? NaN : Math.trunc(a / b)), 2), sig: 'QUOTIENT(numerator, denominator)', desc: 'Integer portion of a division' },
  FACT: {
    ...math((n) => {
      if (n < 0) return NaN
      let r = 1
      for (let i = 2; i <= Math.floor(n); i++) r *= i
      return r
    }),
    sig: 'FACT(number)',
    desc: 'Factorial'
  },
  GCD: {
    min: 1,
    max: 255,
    fn: aggregate((ns) => ns.map((n) => Math.trunc(Math.abs(n))).reduce((a, b) => { while (b) [a, b] = [b, a % b]; return a }, 0)),
    sig: 'GCD(number1, [number2], …)',
    desc: 'Greatest common divisor'
  },
  RAND: { min: 0, max: 0, volatile: true, fn: () => Math.random(), sig: 'RAND()', desc: 'Random number between 0 and 1' },
  RANDBETWEEN: {
    min: 2,
    max: 2,
    volatile: true,
    fn: ([a, b]) => {
      const lo = num(a)
      const hi = num(b)
      if (isError(lo)) return lo
      if (isError(hi)) return hi
      if (lo > hi) return Errors.NUM
      const l = Math.ceil(lo)
      return l + Math.floor(Math.random() * (Math.floor(hi) - l + 1))
    },
    sig: 'RANDBETWEEN(bottom, top)',
    desc: 'Random integer in a range'
  },
  SUMPRODUCT: {
    min: 1,
    max: 255,
    fn: (args) => {
      const ds = args.map(dims)
      const rows = ds[0].rows
      const cols = ds[0].cols
      if (ds.some((d) => d.rows !== rows || d.cols !== cols)) return Errors.VALUE
      const R = Math.min(...args.map(usedRows))
      const C = Math.min(...args.map(usedCols))
      let total = 0
      for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) {
        let p = 1
        for (const d of ds) {
          const v = d.get(r, c)
          if (isError(v)) return v
          p *= typeof v === 'number' ? v : 0
        }
        total += p
      }
      return total
    },
    sig: 'SUMPRODUCT(array1, [array2], …)',
    desc: 'Sum of products of corresponding entries'
  },
  SUMSQ: { min: 1, max: 255, fn: aggregate((ns) => ns.reduce((a, b) => a + b * b, 0)), sig: 'SUMSQ(number1, [number2], …)', desc: 'Sum of squares' },

  // Conditional aggregates -------------------------------------------------
  // SUMIF/AVERAGEIF's optional sum range is resized to the criteria range's
  // shape by the compiler (as Excel does), so it is a proper precedent here.
  SUMIF: {
    min: 2,
    max: 3,
    fn: ([range, crit, sumRange]) => {
      let total = 0
      const err = eachMatch(sumRange ?? range, [range, crit], (v) => { total += v })
      return err ?? total
    },
    sig: 'SUMIF(range, criterion, [sum_range])',
    desc: 'Sums cells that meet a criterion'
  },
  SUMIFS: {
    min: 3,
    max: 255,
    fn: ([sumRange, ...pairs]) => {
      let total = 0
      const err = eachMatch(sumRange, pairs, (v) => { total += v })
      return err ?? total
    },
    sig: 'SUMIFS(sum_range, criteria_range1, criterion1, …)',
    desc: 'Sums cells that meet several criteria'
  },
  COUNTIF: {
    min: 2,
    max: 2,
    fn: (args) => {
      let n = 0
      const { err, blankHits } = ifsScan(args, () => n++)
      return err ?? n + blankHits
    },
    sig: 'COUNTIF(range, criterion)',
    desc: 'Counts cells that meet a criterion'
  },
  COUNTIFS: {
    min: 2,
    max: 254,
    fn: (args) => {
      let n = 0
      const { err, blankHits } = ifsScan(args, () => n++)
      return err ?? n + blankHits
    },
    sig: 'COUNTIFS(criteria_range1, criterion1, …)',
    desc: 'Counts cells that meet several criteria'
  },
  AVERAGEIF: {
    min: 2,
    max: 3,
    fn: ([range, crit, avgRange]) => {
      let total = 0
      let n = 0
      const err = eachMatch(avgRange ?? range, [range, crit], (v) => { total += v; n++ })
      return err ?? (n ? total / n : Errors.DIV0)
    },
    sig: 'AVERAGEIF(range, criterion, [average_range])',
    desc: 'Averages cells that meet a criterion'
  },
  AVERAGEIFS: {
    min: 3,
    max: 255,
    fn: ([avgRange, ...pairs]) => {
      let total = 0
      let n = 0
      const err = eachMatch(avgRange, pairs, (v) => { total += v; n++ })
      return err ?? (n ? total / n : Errors.DIV0)
    },
    sig: 'AVERAGEIFS(average_range, criteria_range1, criterion1, …)',
    desc: 'Averages cells that meet several criteria'
  },
  MAXIFS: {
    min: 3,
    max: 255,
    fn: ([range, ...pairs]) => {
      let best = -Infinity
      const err = eachMatch(range, pairs, (v) => { if (v > best) best = v })
      return err ?? (best === -Infinity ? 0 : best)
    },
    sig: 'MAXIFS(max_range, criteria_range1, criterion1, …)',
    desc: 'Maximum of cells that meet criteria'
  },
  MINIFS: {
    min: 3,
    max: 255,
    fn: ([range, ...pairs]) => {
      let best = Infinity
      const err = eachMatch(range, pairs, (v) => { if (v < best) best = v })
      return err ?? (best === Infinity ? 0 : best)
    },
    sig: 'MINIFS(min_range, criteria_range1, criterion1, …)',
    desc: 'Minimum of cells that meet criteria'
  },

  // Statistics -------------------------------------------------------------
  STDEV: { min: 1, max: 255, fn: aggregate((ns) => { const v = variance(ns, true); return isError(v) ? v : Math.sqrt(v) }), sig: 'STDEV(number1, [number2], …)', desc: 'Sample standard deviation' },
  'STDEV.S': { min: 1, max: 255, fn: aggregate((ns) => { const v = variance(ns, true); return isError(v) ? v : Math.sqrt(v) }), sig: 'STDEV.S(number1, [number2], …)', desc: 'Sample standard deviation' },
  'STDEV.P': { min: 1, max: 255, fn: aggregate((ns) => { const v = variance(ns, false); return isError(v) ? v : Math.sqrt(v) }), sig: 'STDEV.P(number1, [number2], …)', desc: 'Population standard deviation' },
  VAR: { min: 1, max: 255, fn: aggregate((ns) => variance(ns, true)), sig: 'VAR(number1, [number2], …)', desc: 'Sample variance' },
  'VAR.S': { min: 1, max: 255, fn: aggregate((ns) => variance(ns, true)), sig: 'VAR.S(number1, [number2], …)', desc: 'Sample variance' },
  'VAR.P': { min: 1, max: 255, fn: aggregate((ns) => variance(ns, false)), sig: 'VAR.P(number1, [number2], …)', desc: 'Population variance' },
  LARGE: {
    min: 2,
    max: 2,
    fn: ([arr, k]) => {
      const ns = collectNumbers([arr])
      const kk = int(k)
      if (isError(ns)) return ns
      if (isError(kk)) return kk
      if (kk < 1 || kk > ns.length) return Errors.NUM
      return ns.sort((a, b) => b - a)[kk - 1]
    },
    sig: 'LARGE(array, k)',
    desc: 'k-th largest value'
  },
  SMALL: {
    min: 2,
    max: 2,
    fn: ([arr, k]) => {
      const ns = collectNumbers([arr])
      const kk = int(k)
      if (isError(ns)) return ns
      if (isError(kk)) return kk
      if (kk < 1 || kk > ns.length) return Errors.NUM
      return ns.sort((a, b) => a - b)[kk - 1]
    },
    sig: 'SMALL(array, k)',
    desc: 'k-th smallest value'
  },
  RANK: {
    min: 2,
    max: 3,
    fn: ([n, ref, order]) => {
      const x = num(n)
      if (isError(x)) return x
      const ns = collectNumbers([ref])
      if (isError(ns)) return ns
      const asc = order != null && toBoolean(order) === true
      if (!ns.includes(x)) return Errors.NA
      let rank = 1
      for (const v of ns) if (asc ? v < x : v > x) rank++
      return rank
    },
    sig: 'RANK(number, ref, [order])',
    desc: 'Rank of a number in a list'
  },
  CORREL: {
    min: 2,
    max: 2,
    fn: ([a, b]) => {
      const da = dims(a)
      const db = dims(b)
      const n = da.rows * da.cols
      if (n !== db.rows * db.cols) return Errors.NA
      // Pair values by position in row-major order, so A1:C1 and A2:A4 line up.
      const xs: number[] = []
      const ys: number[] = []
      const limit = Math.min(n, Math.max(usedRows(a) * da.cols, usedRows(b) * db.cols))
      for (let i = 0; i < limit; i++) {
        const x = da.get(Math.floor(i / da.cols), i % da.cols)
        const y = db.get(Math.floor(i / db.cols), i % db.cols)
        if (typeof x === 'number' && typeof y === 'number') {
          xs.push(x)
          ys.push(y)
        }
      }
      if (xs.length < 2) return Errors.DIV0
      const mx = mean(xs)
      const my = mean(ys)
      let sxy = 0
      let sxx = 0
      let syy = 0
      for (let i = 0; i < xs.length; i++) {
        sxy += (xs[i] - mx) * (ys[i] - my)
        sxx += (xs[i] - mx) ** 2
        syy += (ys[i] - my) ** 2
      }
      return sxx && syy ? sxy / Math.sqrt(sxx * syy) : Errors.DIV0
    },
    sig: 'CORREL(array1, array2)',
    desc: 'Correlation coefficient'
  },

  // Logic ------------------------------------------------------------------
  IF: {
    min: 1,
    max: 3,
    lazy: ([cond, a, b]) => {
      const c = truthy(cond())
      if (isError(c)) return c
      if (c) return a ? a() : true
      return b ? b() : false
    },
    sig: 'IF(condition, value_if_true, [value_if_false])',
    desc: 'Returns one value if true, another if false'
  },
  IFS: {
    min: 2,
    max: 254,
    lazy: (args) => {
      for (let i = 0; i + 1 < args.length; i += 2) {
        const c = truthy(args[i]())
        if (isError(c)) return c
        if (c) return args[i + 1]()
      }
      return Errors.NA
    },
    sig: 'IFS(condition1, value1, [condition2, value2], …)',
    desc: 'First value whose condition is true'
  },
  IFERROR: {
    min: 2,
    max: 2,
    lazy: ([a, b]) => {
      const v = toScalar(a())
      return isError(v) ? b() : v
    },
    sig: 'IFERROR(value, value_if_error)',
    desc: 'Fallback value when there is an error'
  },
  IFNA: {
    min: 2,
    max: 2,
    lazy: ([a, b]) => {
      const v = toScalar(a())
      return v === Errors.NA ? b() : v
    },
    sig: 'IFNA(value, value_if_na)',
    desc: 'Fallback value for #N/A'
  },
  SWITCH: {
    min: 3,
    max: 254,
    lazy: ([expr, ...rest]) => {
      const v = toScalar(expr())
      if (isError(v)) return v
      let i = 0
      for (; i + 1 < rest.length; i += 2) {
        const cand = toScalar(rest[i]())
        if (cand !== null && compareScalars(v, cand) === 0 && typeof v === typeof cand) return rest[i + 1]()
      }
      return i < rest.length ? rest[i]() : Errors.NA
    },
    sig: 'SWITCH(expression, case1, value1, …, [default])',
    desc: 'Matches an expression against cases'
  },
  AND: { min: 1, max: 255, fn: (args) => logical(args, 'and'), sig: 'AND(logical1, [logical2], …)', desc: 'True if all arguments are true' },
  OR: { min: 1, max: 255, fn: (args) => logical(args, 'or'), sig: 'OR(logical1, [logical2], …)', desc: 'True if any argument is true' },
  XOR: { min: 1, max: 255, fn: (args) => logical(args, 'xor'), sig: 'XOR(logical1, [logical2], …)', desc: 'True if an odd number of arguments are true' },
  NOT: {
    min: 1,
    max: 1,
    fn: ([a]) => {
      const b = truthy(a)
      return isError(b) ? b : !b
    },
    sig: 'NOT(logical)',
    desc: 'Reverses a logical value'
  },
  TRUE: { min: 0, max: 0, fn: () => true, sig: 'TRUE()', desc: 'The logical value TRUE' },
  FALSE: { min: 0, max: 0, fn: () => false, sig: 'FALSE()', desc: 'The logical value FALSE' },

  // Information ------------------------------------------------------------
  ISBLANK: { min: 1, max: 1, fn: ([a]) => toScalar(a) === null, sig: 'ISBLANK(value)', desc: 'True if the cell is empty' },
  ISNUMBER: { min: 1, max: 1, fn: ([a]) => typeof toScalar(a) === 'number', sig: 'ISNUMBER(value)', desc: 'True if the value is a number' },
  ISTEXT: { min: 1, max: 1, fn: ([a]) => typeof toScalar(a) === 'string', sig: 'ISTEXT(value)', desc: 'True if the value is text' },
  ISNONTEXT: { min: 1, max: 1, fn: ([a]) => typeof toScalar(a) !== 'string', sig: 'ISNONTEXT(value)', desc: 'True if the value is not text' },
  ISLOGICAL: { min: 1, max: 1, fn: ([a]) => typeof toScalar(a) === 'boolean', sig: 'ISLOGICAL(value)', desc: 'True if the value is TRUE/FALSE' },
  ISERROR: { min: 1, max: 1, fn: ([a]) => isError(toScalar(a)), sig: 'ISERROR(value)', desc: 'True if the value is any error' },
  ISERR: { min: 1, max: 1, fn: ([a]) => { const v = toScalar(a); return isError(v) && v !== Errors.NA }, sig: 'ISERR(value)', desc: 'True for errors other than #N/A' },
  ISNA: { min: 1, max: 1, fn: ([a]) => toScalar(a) === Errors.NA, sig: 'ISNA(value)', desc: 'True if the value is #N/A' },
  ISEVEN: { min: 1, max: 1, fn: ([a]) => { const n = int(a); return isError(n) ? n : n % 2 === 0 }, sig: 'ISEVEN(number)', desc: 'True if the number is even' },
  ISODD: { min: 1, max: 1, fn: ([a]) => { const n = int(a); return isError(n) ? n : Math.abs(n % 2) === 1 }, sig: 'ISODD(number)', desc: 'True if the number is odd' },
  NA: { min: 0, max: 0, fn: () => Errors.NA, sig: 'NA()', desc: 'The #N/A error' },
  N: { min: 1, max: 1, fn: ([a]) => { const v = toScalar(a); return typeof v === 'number' ? v : typeof v === 'boolean' ? +v : isError(v) ? v : 0 }, sig: 'N(value)', desc: 'Converts a value to a number' },
  ROW: { min: 0, max: 1, fn: ([r], ctx) => (r && isRange(r) ? r.r1 + 1 : r ? Errors.VALUE : ctx.row + 1), sig: 'ROW([reference])', desc: 'Row number of a reference' },
  COLUMN: { min: 0, max: 1, fn: ([r], ctx) => (r && isRange(r) ? r.c1 + 1 : r ? Errors.VALUE : ctx.col + 1), sig: 'COLUMN([reference])', desc: 'Column number of a reference' },
  ROWS: { min: 1, max: 1, fn: ([r]) => dims(r).rows, sig: 'ROWS(range)', desc: 'Number of rows in a range' },
  COLUMNS: { min: 1, max: 1, fn: ([r]) => dims(r).cols, sig: 'COLUMNS(range)', desc: 'Number of columns in a range' },

  // Lookup -----------------------------------------------------------------
  VLOOKUP: {
    min: 3,
    max: 4,
    fn: ([needle, table, colIndex, approx]) => {
      const key = toScalar(needle)
      if (isError(key)) return key
      const ci = int(colIndex)
      if (isError(ci)) return ci
      const d = dims(table)
      if (ci < 1 || ci > d.cols) return Errors.REF
      const sorted = approx == null ? true : toBoolean(approx)
      if (isError(sorted)) return sorted
      const i = lookupMatch(key, (r) => d.get(r, 0), usedRows(table), sorted ? 1 : 0)
      return i < 0 ? Errors.NA : d.get(i, ci - 1)
    },
    sig: 'VLOOKUP(search_key, range, index, [is_sorted])',
    desc: 'Vertical lookup in the first column of a range'
  },
  HLOOKUP: {
    min: 3,
    max: 4,
    fn: ([needle, table, rowIndex, approx]) => {
      const key = toScalar(needle)
      if (isError(key)) return key
      const ri = int(rowIndex)
      if (isError(ri)) return ri
      const d = dims(table)
      if (ri < 1 || ri > d.rows) return Errors.REF
      const sorted = approx == null ? true : toBoolean(approx)
      if (isError(sorted)) return sorted
      const i = lookupMatch(key, (c) => d.get(0, c), usedCols(table), sorted ? 1 : 0)
      return i < 0 ? Errors.NA : d.get(ri - 1, i)
    },
    sig: 'HLOOKUP(search_key, range, index, [is_sorted])',
    desc: 'Horizontal lookup in the first row of a range'
  },
  MATCH: {
    min: 2,
    max: 3,
    fn: ([needle, range, type]) => {
      const key = toScalar(needle)
      if (isError(key)) return key
      const mode = type == null ? 1 : int(type)
      if (isError(mode)) return mode
      const d = dims(range)
      const vertical = d.cols === 1
      if (!vertical && d.rows !== 1) return Errors.NA
      const n = vertical ? usedRows(range) : usedCols(range)
      const i = lookupMatch(key, (k) => (vertical ? d.get(k, 0) : d.get(0, k)), n, Math.sign(mode))
      return i < 0 ? Errors.NA : i + 1
    },
    sig: 'MATCH(search_key, range, [match_type])',
    desc: 'Position of a value in a range'
  },
  INDEX: {
    min: 2,
    max: 3,
    fn: ([range, row, col], ctx) => {
      const d = dims(range)
      const r = int(row)
      const c = col == null ? (d.cols === 1 ? 1 : d.rows === 1 ? r : 1) : int(col)
      if (isError(r)) return r
      if (isError(c)) return c
      const rr = d.rows === 1 && col == null ? 1 : r
      if (rr < 0 || c < 0 || rr > d.rows || c > d.cols) return Errors.REF
      if (isRange(range) && (rr === 0 || c === 0)) {
        return ctx.makeRange(
          range.sheet,
          rr === 0 ? range.r1 : range.r1 + rr - 1,
          c === 0 ? range.c1 : range.c1 + c - 1,
          rr === 0 ? range.r2 : range.r1 + rr - 1,
          c === 0 ? range.c2 : range.c1 + c - 1
        )
      }
      if (isRange(range)) return ctx.makeRange(range.sheet, range.r1 + rr - 1, range.c1 + c - 1, range.r1 + rr - 1, range.c1 + c - 1)
      return d.get(Math.max(rr, 1) - 1, Math.max(c, 1) - 1)
    },
    sig: 'INDEX(range, row, [column])',
    desc: 'Value at a row and column of a range'
  },
  XLOOKUP: {
    min: 3,
    max: 5,
    fn: ([needle, lookup, ret, notFound, matchMode]) => {
      const key = toScalar(needle)
      if (isError(key)) return key
      const dl = dims(lookup)
      const dr = dims(ret)
      const vertical = dl.cols === 1
      const n = vertical ? usedRows(lookup) : usedCols(lookup)
      const mm = matchMode == null ? 0 : int(matchMode)
      if (isError(mm)) return mm
      let idx = -1
      if (mm === 0) idx = lookupMatch(key, (k) => (vertical ? dl.get(k, 0) : dl.get(0, k)), n, 0, false)
      else {
        // -1: exact or next smaller; 1: exact or next larger (unsorted scan)
        let best: Scalar = null
        for (let k = 0; k < n; k++) {
          const v = vertical ? dl.get(k, 0) : dl.get(0, k)
          if (v === null) continue
          const c = compareScalars(v, key)
          if (c === 0) { idx = k; break }
          if ((mm === -1 && c < 0 && (best === null || compareScalars(v, best) > 0)) || (mm === 1 && c > 0 && (best === null || compareScalars(v, best) < 0))) {
            best = v
            idx = k
          }
        }
      }
      if (idx < 0) return notFound != null ? toScalar(notFound) : Errors.NA
      return vertical ? dr.get(idx, 0) : dr.get(0, idx)
    },
    sig: 'XLOOKUP(search_key, lookup_range, result_range, [if_not_found], [match_mode])',
    desc: 'Looks up a value and returns the matching result'
  },
  CHOOSE: {
    min: 2,
    max: 255,
    lazy: ([index, ...options]) => {
      const i = int(index())
      if (isError(i)) return i
      if (i < 1 || i > options.length) return Errors.VALUE
      return options[i - 1]()
    },
    sig: 'CHOOSE(index, choice1, [choice2], …)',
    desc: 'Picks a value from a list by index'
  },
  INDIRECT: {
    min: 1,
    max: 2,
    volatile: true,
    fn: ([ref], ctx) => {
      const t = str(ref)
      return isError(t) ? t : ctx.rangeFromText(t, ctx.sheetId)
    },
    sig: 'INDIRECT(ref_text)',
    desc: 'Reference given by a text string'
  },
  OFFSET: {
    min: 3,
    max: 5,
    volatile: true,
    fn: ([ref, rows, cols, h, w], ctx) => {
      if (!isRange(ref)) return Errors.VALUE
      const dr = int(rows)
      const dc = int(cols)
      const hh = h == null ? ref.r2 - ref.r1 + 1 : int(h)
      const ww = w == null ? ref.c2 - ref.c1 + 1 : int(w)
      for (const x of [dr, dc, hh, ww]) if (isError(x)) return x
      const r1 = ref.r1 + (dr as number)
      const c1 = ref.c1 + (dc as number)
      if ((hh as number) < 1 || (ww as number) < 1) return Errors.REF
      return ctx.makeRange(ref.sheet, r1, c1, r1 + (hh as number) - 1, c1 + (ww as number) - 1)
    },
    sig: 'OFFSET(reference, rows, cols, [height], [width])',
    desc: 'Reference shifted from a starting point'
  },

  // Text -------------------------------------------------------------------
  CONCAT: {
    min: 1,
    max: 255,
    fn: (args) => {
      let out = ''
      for (const a of args) {
        let err: FormulaError | undefined
        forEachValue(a, (s) => {
          const t = toText(s)
          if (isError(t)) err ??= t
          else out += t
        }, true)
        if (err) return err
      }
      return out
    },
    sig: 'CONCAT(text1, [text2], …)',
    desc: 'Joins text together'
  },
  CONCATENATE: {
    min: 1,
    max: 255,
    fn: (args) => {
      let out = ''
      for (const a of args) {
        const t = str(a)
        if (isError(t)) return t
        out += t
      }
      return out
    },
    sig: 'CONCATENATE(text1, [text2], …)',
    desc: 'Joins text together'
  },
  TEXTJOIN: {
    min: 3,
    max: 255,
    fn: ([delim, ignoreEmpty, ...args]) => {
      const d = str(delim)
      if (isError(d)) return d
      const skip = toBoolean(ignoreEmpty)
      if (isError(skip)) return skip
      const parts: string[] = []
      for (const a of args) {
        let err: FormulaError | undefined
        forEachValue(a, (s) => {
          const t = toText(s)
          if (isError(t)) err ??= t
          else if (!(skip && t === '')) parts.push(t)
        }, true)
        if (err) return err
      }
      return parts.join(d)
    },
    sig: 'TEXTJOIN(delimiter, ignore_empty, text1, …)',
    desc: 'Joins text with a delimiter'
  },
  LEN: { min: 1, max: 1, fn: ([a]) => { const t = str(a); return isError(t) ? t : t.length }, sig: 'LEN(text)', desc: 'Length of text' },
  UPPER: { min: 1, max: 1, fn: ([a]) => { const t = str(a); return isError(t) ? t : t.toUpperCase() }, sig: 'UPPER(text)', desc: 'Converts to upper case' },
  LOWER: { min: 1, max: 1, fn: ([a]) => { const t = str(a); return isError(t) ? t : t.toLowerCase() }, sig: 'LOWER(text)', desc: 'Converts to lower case' },
  PROPER: {
    min: 1,
    max: 1,
    fn: ([a]) => {
      const t = str(a)
      return isError(t) ? t : t.toLowerCase().replace(/(^|[^a-z])([a-z])/g, (_, p, c: string) => p + c.toUpperCase())
    },
    sig: 'PROPER(text)',
    desc: 'Capitalises each word'
  },
  TRIM: { min: 1, max: 1, fn: ([a]) => { const t = str(a); return isError(t) ? t : t.trim().replace(/ {2,}/g, ' ') }, sig: 'TRIM(text)', desc: 'Removes extra spaces' },
  LEFT: {
    min: 1,
    max: 2,
    fn: ([a, n]) => {
      const t = str(a)
      const k = n == null ? 1 : int(n)
      if (isError(t)) return t
      if (isError(k)) return k
      return k < 0 ? Errors.VALUE : t.slice(0, k)
    },
    sig: 'LEFT(text, [num_chars])',
    desc: 'Leftmost characters'
  },
  RIGHT: {
    min: 1,
    max: 2,
    fn: ([a, n]) => {
      const t = str(a)
      const k = n == null ? 1 : int(n)
      if (isError(t)) return t
      if (isError(k)) return k
      return k < 0 ? Errors.VALUE : k === 0 ? '' : t.slice(-k)
    },
    sig: 'RIGHT(text, [num_chars])',
    desc: 'Rightmost characters'
  },
  MID: {
    min: 3,
    max: 3,
    fn: ([a, s, n]) => {
      const t = str(a)
      const st = int(s)
      const k = int(n)
      if (isError(t)) return t
      if (isError(st)) return st
      if (isError(k)) return k
      if (st < 1 || k < 0) return Errors.VALUE
      return t.substr(st - 1, k)
    },
    sig: 'MID(text, start, num_chars)',
    desc: 'Characters from the middle of text'
  },
  FIND: {
    min: 2,
    max: 3,
    fn: ([needle, hay, start]) => {
      const n = str(needle)
      const h = str(hay)
      const s = start == null ? 1 : int(start)
      if (isError(n)) return n
      if (isError(h)) return h
      if (isError(s)) return s
      if (s < 1 || s > h.length + 1) return Errors.VALUE
      const i = h.indexOf(n, s - 1)
      return i < 0 ? Errors.VALUE : i + 1
    },
    sig: 'FIND(search_for, text, [start])',
    desc: 'Case-sensitive position of text'
  },
  SEARCH: {
    min: 2,
    max: 3,
    fn: ([needle, hay, start]) => {
      const n = str(needle)
      const h = str(hay)
      const s = start == null ? 1 : int(start)
      if (isError(n)) return n
      if (isError(h)) return h
      if (isError(s)) return s
      if (s < 1 || s > h.length + 1) return Errors.VALUE
      const re = wildcardRegex(n)
      if (re) {
        const src = re.source.slice(1, -1)
        const m = new RegExp(src, 'is').exec(h.slice(s - 1))
        return m ? m.index + s : Errors.VALUE
      }
      const i = h.toLowerCase().indexOf(n.toLowerCase(), s - 1)
      return i < 0 ? Errors.VALUE : i + 1
    },
    sig: 'SEARCH(search_for, text, [start])',
    desc: 'Case-insensitive position of text'
  },
  SUBSTITUTE: {
    min: 3,
    max: 4,
    fn: ([text, from, to, which]) => {
      const t = str(text)
      const f = str(from)
      const r = str(to)
      if (isError(t)) return t
      if (isError(f)) return f
      if (isError(r)) return r
      if (f === '') return t
      if (which == null) return t.split(f).join(r)
      const w = int(which)
      if (isError(w)) return w
      if (w < 1) return Errors.VALUE
      let idx = -1
      for (let i = 0; i < w; i++) {
        idx = t.indexOf(f, idx + 1)
        if (idx < 0) return t
      }
      return t.slice(0, idx) + r + t.slice(idx + f.length)
    },
    sig: 'SUBSTITUTE(text, search_for, replace_with, [occurrence])',
    desc: 'Replaces existing text with new text'
  },
  REPLACE: {
    min: 4,
    max: 4,
    fn: ([text, start, n, rep]) => {
      const t = str(text)
      const s = int(start)
      const k = int(n)
      const r = str(rep)
      if (isError(t)) return t
      if (isError(s)) return s
      if (isError(k)) return k
      if (isError(r)) return r
      if (s < 1 || k < 0) return Errors.VALUE
      return t.slice(0, s - 1) + r + t.slice(s - 1 + k)
    },
    sig: 'REPLACE(text, position, length, new_text)',
    desc: 'Replaces part of a text string'
  },
  REPT: {
    min: 2,
    max: 2,
    fn: ([text, n]) => {
      const t = str(text)
      const k = int(n)
      if (isError(t)) return t
      if (isError(k)) return k
      return k < 0 || t.length * k > 32767 ? Errors.VALUE : t.repeat(k)
    },
    sig: 'REPT(text, number_times)',
    desc: 'Repeats text'
  },
  EXACT: {
    min: 2,
    max: 2,
    fn: ([a, b]) => {
      const x = str(a)
      const y = str(b)
      if (isError(x)) return x
      if (isError(y)) return y
      return x === y
    },
    sig: 'EXACT(text1, text2)',
    desc: 'True if two strings are identical'
  },
  VALUE: {
    min: 1,
    max: 1,
    fn: ([a]) => {
      const v = toScalar(a)
      if (typeof v === 'number') return v
      if (typeof v !== 'string') return isError(v) ? v : Errors.VALUE
      const n = parseNumber(v)
      return n === undefined ? Errors.VALUE : n
    },
    sig: 'VALUE(text)',
    desc: 'Converts text to a number'
  },
  TEXT: {
    min: 2,
    max: 2,
    fn: ([a, f]) => {
      const fmt = str(f)
      if (isError(fmt)) return fmt
      const v = toScalar(a)
      if (typeof v !== 'number') return toText(v)
      return textFormat(v, fmt)
    },
    sig: 'TEXT(number, format)',
    desc: 'Formats a number as text'
  },
  CHAR: { min: 1, max: 1, fn: ([a]) => { const n = int(a); return isError(n) ? n : n < 1 || n > 255 ? Errors.VALUE : String.fromCharCode(n) }, sig: 'CHAR(number)', desc: 'Character for a code' },
  CODE: { min: 1, max: 1, fn: ([a]) => { const t = str(a); return isError(t) ? t : t.length ? t.charCodeAt(0) : Errors.VALUE }, sig: 'CODE(text)', desc: 'Code of the first character' },
  T: { min: 1, max: 1, fn: ([a]) => { const v = toScalar(a); return typeof v === 'string' || isError(v) ? v : '' }, sig: 'T(value)', desc: 'Text if the value is text' },

  // Date & time ------------------------------------------------------------
  TODAY: { min: 0, max: 0, volatile: true, fn: today, sig: 'TODAY()', desc: "Today's date" },
  NOW: { min: 0, max: 0, volatile: true, fn: now, sig: 'NOW()', desc: 'Current date and time' },
  DATE: {
    min: 3,
    max: 3,
    fn: ([y, m, d]) => {
      const Y = int(y)
      const M = int(m)
      const D = int(d)
      if (isError(Y)) return Y
      if (isError(M)) return M
      if (isError(D)) return D
      const yy = Y < 1900 ? Y + 1900 : Y
      const s = dateToSerial(yy, M, D)
      return s < 0 ? Errors.NUM : s
    },
    sig: 'DATE(year, month, day)',
    desc: 'Date serial number'
  },
  TIME: {
    min: 3,
    max: 3,
    fn: ([h, m, s]) => {
      const H = num(h)
      const M = num(m)
      const S = num(s)
      if (isError(H)) return H
      if (isError(M)) return M
      if (isError(S)) return S
      const t = (Math.trunc(H) * 3600 + Math.trunc(M) * 60 + Math.trunc(S)) / 86400
      if (t < 0) return Errors.NUM
      return t - Math.floor(t)
    },
    sig: 'TIME(hour, minute, second)',
    desc: 'Time as a fraction of a day'
  },
  YEAR: { min: 1, max: 1, fn: ([a]) => datePart(a, (d) => d.getUTCFullYear()), sig: 'YEAR(date)', desc: 'Year of a date' },
  MONTH: { min: 1, max: 1, fn: ([a]) => datePart(a, (d) => d.getUTCMonth() + 1), sig: 'MONTH(date)', desc: 'Month of a date' },
  DAY: { min: 1, max: 1, fn: ([a]) => datePart(a, (d) => d.getUTCDate()), sig: 'DAY(date)', desc: 'Day of a date' },
  HOUR: { min: 1, max: 1, fn: ([a]) => datePart(a, (d) => d.getUTCHours()), sig: 'HOUR(time)', desc: 'Hour of a time' },
  MINUTE: { min: 1, max: 1, fn: ([a]) => datePart(a, (d) => d.getUTCMinutes()), sig: 'MINUTE(time)', desc: 'Minute of a time' },
  SECOND: { min: 1, max: 1, fn: ([a]) => datePart(a, (d) => d.getUTCSeconds()), sig: 'SECOND(time)', desc: 'Second of a time' },
  WEEKDAY: {
    min: 1,
    max: 2,
    fn: ([a, type]) => {
      const t = type == null ? 1 : int(type)
      if (isError(t)) return t
      return datePart(a, (d) => {
        const w = d.getUTCDay()
        return t === 2 ? ((w + 6) % 7) + 1 : t === 3 ? (w + 6) % 7 : w + 1
      })
    },
    sig: 'WEEKDAY(date, [type])',
    desc: 'Day of the week'
  },
  EDATE: {
    min: 2,
    max: 2,
    fn: ([a, m]) => {
      const s = num(a)
      const k = int(m)
      if (isError(s)) return s
      if (isError(k)) return k
      const d = serialToDate(s)
      const target = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + k, 1))
      const last = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate()
      return dateToSerial(target.getUTCFullYear(), target.getUTCMonth() + 1, Math.min(d.getUTCDate(), last))
    },
    sig: 'EDATE(start_date, months)',
    desc: 'Date a number of months away'
  },
  EOMONTH: {
    min: 2,
    max: 2,
    fn: ([a, m]) => {
      const s = num(a)
      const k = int(m)
      if (isError(s)) return s
      if (isError(k)) return k
      const d = serialToDate(s)
      const last = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + k + 1, 0))
      return dateToSerial(last.getUTCFullYear(), last.getUTCMonth() + 1, last.getUTCDate())
    },
    sig: 'EOMONTH(start_date, months)',
    desc: 'Last day of a month a number of months away'
  },
  DAYS: {
    min: 2,
    max: 2,
    fn: ([e, s]) => {
      const a = num(e)
      const b = num(s)
      if (isError(a)) return a
      if (isError(b)) return b
      return Math.floor(a) - Math.floor(b)
    },
    sig: 'DAYS(end_date, start_date)',
    desc: 'Days between two dates'
  },

  // Finance ----------------------------------------------------------------
  PMT: {
    ...math((rate, nper, pv, fv = 0, type = 0) => {
      if (rate === 0) return -(pv + fv) / nper
      const f = Math.pow(1 + rate, nper)
      return (-rate * (pv * f + fv)) / ((1 + rate * type) * (f - 1))
    }, 3, 5),
    sig: 'PMT(rate, nper, pv, [fv], [type])',
    desc: 'Periodic payment for a loan'
  },
  FV: {
    ...math((rate, nper, pmt, pv = 0, type = 0) => {
      if (rate === 0) return -(pv + pmt * nper)
      const f = Math.pow(1 + rate, nper)
      return -(pv * f + (pmt * (1 + rate * type) * (f - 1)) / rate)
    }, 3, 5),
    sig: 'FV(rate, nper, pmt, [pv], [type])',
    desc: 'Future value of an investment'
  },
  PV: {
    ...math((rate, nper, pmt, fv = 0, type = 0) => {
      if (rate === 0) return -(fv + pmt * nper)
      const f = Math.pow(1 + rate, nper)
      return -(fv + (pmt * (1 + rate * type) * (f - 1)) / rate) / f
    }, 3, 5),
    sig: 'PV(rate, nper, pmt, [fv], [type])',
    desc: 'Present value of an investment'
  },
  NPV: {
    min: 2,
    max: 255,
    fn: ([rate, ...vals]) => {
      const r = num(rate)
      if (isError(r)) return r
      const ns = collectNumbers(vals, true)
      if (isError(ns)) return ns
      return ns.reduce((acc, v, i) => acc + v / Math.pow(1 + r, i + 1), 0)
    },
    sig: 'NPV(rate, value1, [value2], …)',
    desc: 'Net present value'
  }
}

// Aliases
FUNCTIONS['CEILING.MATH'] = FUNCTIONS.CEILING
FUNCTIONS['FLOOR.MATH'] = FUNCTIONS.FLOOR
FUNCTIONS['MODE.SNGL'] = FUNCTIONS.MODE
FUNCTIONS['RANK.EQ'] = FUNCTIONS.RANK

/** All function names, sorted (kept up to date by registerFunction). */
export const FUNCTION_NAMES: string[] = Object.keys(FUNCTIONS).sort()

/** A custom function, as passed to {@link registerFunction}. */
export interface CustomFunction {
  /** Called with evaluated arguments (ranges arrive as RangeValue objects). */
  fn?: FunctionDef['fn']
  /** Called with argument thunks instead, for short-circuiting functions. */
  lazy?: FunctionDef['lazy']
  /** Minimum and maximum argument counts (default 0 and 255). */
  min?: number
  max?: number
  /** Recalculate on every change, like RAND() or NOW(). */
  volatile?: boolean
  /** Signature and description shown in autocomplete and the function reference. */
  sig?: string
  desc?: string
}

/**
 * Add (or replace) a spreadsheet function, e.g.
 * `registerFunction('DOUBLE', { min: 1, max: 1, fn: ([x]) => toNumber(x) * 2 })`.
 *
 * Register functions before loading workbooks that use them; formulas
 * already compiled keep returning #NAME? until `workbook.rebuild()`.
 */
export function registerFunction (name: string, def: CustomFunction): void {
  const upper = name.toUpperCase()
  if (!/^[A-Z][A-Z0-9._]*$/.test(upper)) throw new Error(`Invalid function name "${name}"`)
  if (!def.fn === !def.lazy) throw new Error(`${upper}: provide exactly one of fn or lazy`)
  const min = def.min ?? 0
  const max = def.max ?? 255
  if (min < 0 || max < min) throw new Error(`${upper}: invalid argument counts ${min}–${max}`)
  const isNew = !(upper in FUNCTIONS)
  FUNCTIONS[upper] = {
    fn: def.fn,
    lazy: def.lazy,
    min,
    max,
    volatile: def.volatile,
    sig: def.sig ?? `${upper}(…)`,
    desc: def.desc ?? 'Custom function'
  }
  if (isNew) {
    FUNCTION_NAMES.push(upper)
    FUNCTION_NAMES.sort()
  }
}

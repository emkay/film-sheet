import { MAX_COLS, MAX_ROWS } from './address.js'
import { FUNCTIONS, type FnContext, type Thunk } from './functions.js'
import { ParseError, parseFormula, type Node } from './parser.js'
import {
  compareScalars,
  Errors,
  FormulaError,
  errorFromCode,
  isError,
  toNumber,
  toScalar,
  toText,
  type ArrayValue,
  type RangeValue,
  type Scalar,
  type Value
} from './values.js'

/** Where a formula reads from — used to build the dependency graph. */
export interface Precedent {
  sheetId: number
  r1: number
  c1: number
  r2: number
  c2: number
}

export interface CompiledFormula {
  /** Evaluate in the context of the formula's own cell. */
  evaluate: (ctx: FnContext) => Value
  precedents: Precedent[]
  volatile: boolean
}

/** What the compiler needs from the workbook. */
export interface CompileEnv {
  /** Resolve a sheet name (undefined = the formula's own sheet) to an id. */
  resolveSheet (name: string | undefined): number | undefined
  /** Read a single cell's current value (participates in dirty tracking). */
  read (sheetId: number, r: number, c: number): Scalar
  /** A lazily-evaluated range object. */
  range (sheetId: number, r1: number, c1: number, r2: number, c2: number): RangeValue
}

type Fn = (ctx: FnContext) => Value

export class CompileError extends Error {}

export function compileFormula (body: string, env: CompileEnv): CompiledFormula {
  let ast: Node
  try {
    ast = parseFormula(body)
  } catch (e) {
    throw new CompileError((e as Error).message)
  }
  const precedents: Precedent[] = []
  let volatile = false

  function refTarget (node: Extract<Node, { t: 'ref' }>): Precedent | FormulaError {
    const sheetId = env.resolveSheet(node.ref.sheet)
    if (sheetId === undefined) return Errors.REF
    const { r1, c1, r2, c2 } = node.ref
    return {
      sheetId,
      r1: Math.min(r1, r2),
      c1: Math.min(c1, c2),
      r2: Math.max(r1, r2),
      c2: Math.max(c1, c2)
    }
  }

  /** Compile a node where a range is acceptable (function arguments). */
  function compileArg (node: Node): Fn {
    if (node.t === 'ref') {
      const target = refTarget(node)
      if (isError(target)) return () => target
      precedents.push(target)
      const range = env.range(target.sheetId, target.r1, target.c1, target.r2, target.c2)
      return () => range
    }
    return compile(node)
  }

  function compile (node: Node): Fn {
    switch (node.t) {
      case 'num': {
        const v = node.v
        return () => v
      }
      case 'str': {
        const v = node.v
        return () => v
      }
      case 'bool': {
        const v = node.v
        return () => v
      }
      case 'err': {
        const v = errorFromCode(node.v) ?? Errors.VALUE
        return () => v
      }
      case 'empty':
        return () => null
      case 'name':
        return () => Errors.NAME
      case 'ref': {
        const target = refTarget(node)
        if (isError(target)) return () => target
        precedents.push(target)
        const { sheetId, r1, c1, r2, c2 } = target
        if (r1 === r2 && c1 === c2) return () => env.read(sheetId, r1, c1)
        const range = env.range(sheetId, r1, c1, r2, c2)
        return () => range
      }
      case 'array': {
        const rows = node.rows.map((row) => row.map(compile))
        return (ctx): ArrayValue => ({ kind: 'array', rows: rows.map((row) => row.map((f) => toScalar(f(ctx)))) })
      }
      case 'unary': {
        const a = compile(node.a)
        if (node.op === '+') return a
        return (ctx) => {
          const n = toNumber(a(ctx))
          return isError(n) ? n : -n
        }
      }
      case 'pct': {
        const a = compile(node.a)
        return (ctx) => {
          const n = toNumber(a(ctx))
          return isError(n) ? n : n / 100
        }
      }
      case 'bin':
        return compileBinary(node.op, compile(node.a), compile(node.b))
      case 'call':
        return compileCall(node.name, node.args)
    }
  }

  function compileCall (name: string, args: Node[]): Fn {
    const def = FUNCTIONS[name]
    if (!def) return () => Errors.NAME
    if (args.length < def.min || args.length > def.max) {
      throw new CompileError(`${name} expects ${def.min === def.max ? def.min : `${def.min}–${def.max}`} argument${def.max === 1 ? '' : 's'}, got ${args.length}`)
    }
    if (def.volatile) volatile = true
    if ((name === 'SUMIF' || name === 'AVERAGEIF') && args.length === 3) args = [args[0], args[1], resizeLike(args[2], args[0])]
    const compiled = args.map(compileArg)
    if (def.lazy) {
      const lazy = def.lazy
      return (ctx) => lazy(compiled.map((f): Thunk => () => f(ctx)), ctx)
    }
    const fn = def.fn!
    const n = compiled.length
    // Specialise small arities so the hot path avoids map() allocations.
    if (n === 1) {
      const [a] = compiled
      return (ctx) => fn([a(ctx)], ctx)
    }
    if (n === 2) {
      const [a, b] = compiled
      return (ctx) => fn([a(ctx), b(ctx)], ctx)
    }
    return (ctx) => {
      const vals = new Array<Value>(n)
      for (let i = 0; i < n; i++) vals[i] = compiled[i](ctx)
      return fn(vals, ctx)
    }
  }

  const evaluate = compile(ast)
  return { evaluate, precedents, volatile }
}

function arith (op: (x: number, y: number) => number | FormulaError, a: Fn, b: Fn): Fn {
  return (ctx) => {
    const x = toNumber(a(ctx))
    if (isError(x)) return x
    const y = toNumber(b(ctx))
    if (isError(y)) return y
    const r = op(x, y)
    return typeof r === 'number' && !Number.isFinite(r) ? Errors.NUM : r
  }
}

function compileBinary (op: string, a: Fn, b: Fn): Fn {
  switch (op) {
    case '+':
      return arith((x, y) => x + y, a, b)
    case '-':
      return arith((x, y) => x - y, a, b)
    case '*':
      return arith((x, y) => x * y, a, b)
    case '/':
      return arith((x, y) => (y === 0 ? Errors.DIV0 : x / y), a, b)
    case '^':
      return arith((x, y) => (x === 0 && y === 0 ? Errors.NUM : Math.pow(x, y)), a, b)
    case '&':
      return (ctx) => {
        const x = toText(a(ctx))
        if (isError(x)) return x
        const y = toText(b(ctx))
        if (isError(y)) return y
        return x + y
      }
    default: {
      const test = comparator(op)
      return (ctx) => {
        const x = toScalar(a(ctx))
        if (isError(x)) return x
        const y = toScalar(b(ctx))
        if (isError(y)) return y
        return test(compareScalars(x, y))
      }
    }
  }
}

function comparator (op: string): (c: number) => boolean {
  switch (op) {
    case '=': return (c) => c === 0
    case '<>': return (c) => c !== 0
    case '<': return (c) => c < 0
    case '>': return (c) => c > 0
    case '<=': return (c) => c <= 0
    case '>=': return (c) => c >= 0
  }
  throw new ParseError(`Unknown operator ${op}`)
}

/**
 * SUMIF(A1:A10, ">0", B1) sums B1:B10: the sum range takes the criteria
 * range's shape from its top-left corner. Rewriting the reference here makes
 * the full area a real precedent, so edits to B2:B10 recalculate.
 */
function resizeLike (target: Node, like: Node): Node {
  if (target.t !== 'ref' || like.t !== 'ref') return target
  const rows = Math.abs(like.ref.r2 - like.ref.r1)
  const cols = Math.abs(like.ref.c2 - like.ref.c1)
  const r1 = Math.min(target.ref.r1, target.ref.r2)
  const c1 = Math.min(target.ref.c1, target.ref.c2)
  if (!inGrid(r1 + rows, c1 + cols)) return target
  return { t: 'ref', ref: { ...target.ref, kind: 'range', r1, c1, r2: r1 + rows, c2: c1 + cols } }
}

export function inGrid (r: number, c: number): boolean {
  return r >= 0 && c >= 0 && r < MAX_ROWS && c < MAX_COLS
}

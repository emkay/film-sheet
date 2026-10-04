import { colToName, MAX_COLS, MAX_ROWS, quoteSheetName } from './address.js'
import { tokenize, type RefToken } from './tokenizer.js'

/** Serialise a reference token back to A1 text. */
export function formatRefToken (ref: RefToken): string {
  const [ac1, ar1, ac2, ar2] = ref.abs
  const prefix = ref.sheet !== undefined ? quoteSheetName(ref.sheet) + '!' : ''
  const col = (c: number, abs: boolean): string => (abs ? '$' : '') + colToName(c)
  const row = (r: number, abs: boolean): string => (abs ? '$' : '') + (r + 1)
  switch (ref.kind) {
    case 'cell':
      return prefix + col(ref.c1, ac1) + row(ref.r1, ar1)
    case 'range':
      return prefix + col(ref.c1, ac1) + row(ref.r1, ar1) + ':' + col(ref.c2, ac2) + row(ref.r2, ar2)
    case 'cols':
      return prefix + col(ref.c1, ac1) + ':' + col(ref.c2, ac2)
    case 'rows':
      return prefix + row(ref.r1, ar1) + ':' + row(ref.r2, ar2)
  }
}

/**
 * Rewrite every reference in a formula (text including the leading "=").
 * The callback returns a replacement token, or null for #REF!.
 * Text outside references is preserved byte for byte.
 */
export function rewriteFormula (formula: string, fn: (ref: RefToken) => RefToken | null): string {
  if (formula[0] !== '=') return formula
  let tokens
  try {
    tokens = tokenize(formula.slice(1))
  } catch {
    return formula
  }
  let out = '='
  let changed = false
  for (const t of tokens) {
    if (t.type === 'ref') {
      const next = fn(t.ref!)
      if (next === t.ref) {
        out += t.text
        continue
      }
      changed = true
      out += next ? formatRefToken(next) : '#REF!'
    } else if (t.type === 'str') {
      out += '"' + t.text.replace(/"/g, '""') + '"'
    } else out += t.text
  }
  return changed ? out : formula
}

/** Adjust relative references for a copy from one cell to another (fill, paste). */
export function shiftFormula (formula: string, dr: number, dc: number): string {
  if (dr === 0 && dc === 0) return formula
  return rewriteFormula(formula, (ref) => {
    const [ac1, ar1, ac2, ar2] = ref.abs
    const r1 = ar1 ? ref.r1 : ref.r1 + dr
    const r2 = ar2 ? ref.r2 : ref.r2 + dr
    const c1 = ac1 ? ref.c1 : ref.c1 + dc
    const c2 = ac2 ? ref.c2 : ref.c2 + dc
    if (ref.kind !== 'cols' && (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS)) return null
    if (ref.kind !== 'rows' && (c1 < 0 || c2 < 0 || c1 >= MAX_COLS || c2 >= MAX_COLS)) return null
    if (r1 === ref.r1 && r2 === ref.r2 && c1 === ref.c1 && c2 === ref.c2) return ref
    return {
      ...ref,
      r1: ref.kind === 'cols' ? ref.r1 : r1,
      r2: ref.kind === 'cols' ? ref.r2 : r2,
      c1: ref.kind === 'rows' ? ref.c1 : c1,
      c2: ref.kind === 'rows' ? ref.c2 : c2
    }
  })
}

/**
 * Coordinate transform for inserting (count > 0) or deleting (count < 0)
 * rows/columns at `at`. Returns the new [start, end] or null when the span
 * was deleted entirely.
 */
export function transformSpan (start: number, end: number, at: number, count: number): [number, number] | null {
  if (count > 0) {
    return [start >= at ? start + count : start, end >= at ? end + count : end]
  }
  const n = -count
  const delEnd = at + n // exclusive
  const map = (x: number, isEnd: boolean): number => {
    if (x < at) return x
    if (x >= delEnd) return x - n
    return isEnd ? at - 1 : at
  }
  const s = map(start, false)
  const e = map(end, true)
  if (e < s) return null
  return [s, e]
}

/**
 * Rewrite a formula for a structural change on one sheet. `onSheet` decides
 * whether a reference (by its sheet name, or undefined = same sheet as the
 * formula) points at the sheet being changed.
 */
export function transformFormulaForStructure (
  formula: string,
  axis: 'rows' | 'cols',
  at: number,
  count: number,
  onSheet: (sheetName: string | undefined) => boolean
): string {
  return rewriteFormula(formula, (original) => {
    if (!onSheet(original.sheet)) return original
    // Work on the normalised range so a reversed one (A5:A2) maps correctly.
    const ref = normalizeRef(original)
    const unchanged = (r: RefToken): RefToken => (r.r1 === ref.r1 && r.r2 === ref.r2 && r.c1 === ref.c1 && r.c2 === ref.c2 ? original : r)
    if (axis === 'rows') {
      if (ref.kind === 'cols') return original
      const span = transformSpan(ref.r1, ref.r2, at, count)
      if (!span) return null
      if (span[1] >= MAX_ROWS) return null
      return unchanged({ ...ref, r1: span[0], r2: span[1] })
    }
    if (ref.kind === 'rows') return original
    const span = transformSpan(ref.c1, ref.c2, at, count)
    if (!span) return null
    if (span[1] >= MAX_COLS) return null
    return unchanged({ ...ref, c1: span[0], c2: span[1] })
  })
}

/** Put a range's endpoints in top-left → bottom-right order, keeping each $ with its coordinate. */
function normalizeRef (ref: RefToken): RefToken {
  const [ac1, ar1, ac2, ar2] = ref.abs
  const swapR = ref.r1 > ref.r2
  const swapC = ref.c1 > ref.c2
  if (!swapR && !swapC) return ref
  return {
    ...ref,
    r1: swapR ? ref.r2 : ref.r1,
    r2: swapR ? ref.r1 : ref.r2,
    c1: swapC ? ref.c2 : ref.c1,
    c2: swapC ? ref.c1 : ref.c2,
    abs: [swapC ? ac2 : ac1, swapR ? ar2 : ar1, swapC ? ac1 : ac2, swapR ? ar1 : ar2]
  }
}

/** Rename sheet references inside a formula. */
export function renameSheetInFormula (formula: string, from: string, to: string): string {
  const lower = from.toLowerCase()
  return rewriteFormula(formula, (ref) => (ref.sheet !== undefined && ref.sheet.toLowerCase() === lower ? { ...ref, sheet: to } : ref))
}

/**
 * Tidy a typed formula the way Excel does: upper-case function names and
 * references, and close any parentheses left open.
 */
export function normalizeFormula (text: string): string {
  let toks
  try {
    toks = tokenize(text.slice(1))
  } catch {
    return text
  }
  let out = '='
  let depth = 0
  for (const t of toks) {
    if (t.type === 'func') out += t.text.toUpperCase()
    else if (t.type === 'ref') out += t.ref!.sheet !== undefined ? quoteSheetName(t.ref!.sheet) + '!' + t.text.slice(t.text.lastIndexOf('!') + 1).toUpperCase() : t.text.toUpperCase()
    else if (t.type === 'str') out += '"' + t.text.replace(/"/g, '""') + '"'
    else if (t.type === 'name' && /^(true|false)$/i.test(t.text)) out += t.text.toUpperCase()
    else out += t.text
    if (t.type === '(') depth++
    if (t.type === ')') depth--
  }
  while (depth-- > 0) out += ')'
  return out
}

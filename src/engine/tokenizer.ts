import { MAX_COLS, MAX_ROWS, nameToCol } from './address.js'

export type TokenType =
  | 'num'
  | 'str'
  | 'err'
  | 'ref'
  | 'func'
  | 'name'
  | 'op'
  | '('
  | ')'
  | ','
  | ';'
  | '{'
  | '}'
  | 'ws'

export interface RefToken {
  /** Sheet name as written (unquoted), or undefined for the current sheet. */
  sheet?: string
  kind: 'cell' | 'range' | 'cols' | 'rows'
  r1: number
  c1: number
  r2: number
  c2: number
  /** Absolute flags: [$col1, $row1, $col2, $row2] */
  abs: [boolean, boolean, boolean, boolean]
}

export interface Token {
  type: TokenType
  text: string
  start: number
  end: number
  ref?: RefToken
}

const SHEET = String.raw`(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!`
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})`
const COL = String.raw`(\$?)([A-Za-z]{1,3})`
const ROW = String.raw`(\$?)(\d{1,7})`

const CELL_RANGE_RE = new RegExp(`^(?:${SHEET})?${CELL}(?::${CELL})?`)
const COL_RANGE_RE = new RegExp(`^(?:${SHEET})?${COL}:${COL}`)
const ROW_RANGE_RE = new RegExp(`^(?:${SHEET})?${ROW}:${ROW}`)

const NUMBER_RE = /^(\d+\.?\d*|\.\d+)(e[+-]?\d+)?/i
const NAME_RE = /^[A-Za-z_\\][A-Za-z0-9_.]*/
const ERROR_RE = /^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!|CIRC!|ERROR!)/i

/** True when the text that follows would make a ref-looking match part of a name. */
function continuesName (src: string, pos: number): boolean {
  const ch = src[pos]
  if (ch !== undefined && /[A-Za-z0-9_.(!]/.test(ch)) return true
  // "LOG10 (100)" is a function call, not the cell LOG10 followed by a paren.
  let k = pos
  while (src[k] === ' ') k++
  return k > pos && src[k] === '('
}

function sheetOf (m: RegExpExecArray): string | undefined {
  if (m[1] !== undefined) return m[1].replace(/''/g, "'")
  return m[2]
}

function tryRef (src: string, pos: number): Token | undefined {
  const rest = src.slice(pos)
  let m = CELL_RANGE_RE.exec(rest)
  if (m && !continuesName(src, pos + m[0].length)) {
    const c1 = nameToCol(m[4])
    const r1 = Number(m[6]) - 1
    const hasSecond = m[8] !== undefined
    const c2 = hasSecond ? nameToCol(m[8]) : c1
    const r2 = hasSecond ? Number(m[10]) - 1 : r1
    if (c1 < MAX_COLS && c2 < MAX_COLS && r1 >= 0 && r2 >= 0 && r1 < MAX_ROWS && r2 < MAX_ROWS) {
      return {
        type: 'ref',
        text: m[0],
        start: pos,
        end: pos + m[0].length,
        ref: {
          sheet: sheetOf(m),
          kind: hasSecond ? 'range' : 'cell',
          r1,
          c1,
          r2,
          c2,
          abs: [m[3] === '$', m[5] === '$', hasSecond ? m[7] === '$' : m[3] === '$', hasSecond ? m[9] === '$' : m[5] === '$']
        }
      }
    }
  }
  m = COL_RANGE_RE.exec(rest)
  if (m && !continuesName(src, pos + m[0].length)) {
    const c1 = nameToCol(m[4])
    const c2 = nameToCol(m[6])
    if (c1 >= 0 && c2 >= 0 && c1 < MAX_COLS && c2 < MAX_COLS) {
      return {
        type: 'ref',
        text: m[0],
        start: pos,
        end: pos + m[0].length,
        ref: { sheet: sheetOf(m), kind: 'cols', r1: 0, c1, r2: MAX_ROWS - 1, c2, abs: [m[3] === '$', true, m[5] === '$', true] }
      }
    }
  }
  m = ROW_RANGE_RE.exec(rest)
  if (m && !continuesName(src, pos + m[0].length)) {
    const r1 = Number(m[4]) - 1
    const r2 = Number(m[6]) - 1
    if (r1 >= 0 && r2 >= 0 && r1 < MAX_ROWS && r2 < MAX_ROWS) {
      return {
        type: 'ref',
        text: m[0],
        start: pos,
        end: pos + m[0].length,
        ref: { sheet: sheetOf(m), kind: 'rows', r1, c1: 0, r2, c2: MAX_COLS - 1, abs: [true, m[3] === '$', true, m[5] === '$'] }
      }
    }
  }
  return undefined
}

export class TokenizeError extends Error {}

/**
 * Tokenize a formula body (without the leading "="). Whitespace tokens are
 * kept so that reference rewriting can reproduce the original text exactly.
 */
export function tokenize (src: string): Token[] {
  const tokens: Token[] = []
  let i = 0
  const n = src.length
  while (i < n) {
    const ch = src[i]
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      let j = i + 1
      while (j < n && /\s/.test(src[j])) j++
      tokens.push({ type: 'ws', text: src.slice(i, j), start: i, end: j })
      i = j
      continue
    }
    if (ch === '"') {
      let j = i + 1
      let value = ''
      for (;;) {
        if (j >= n) throw new TokenizeError('Unterminated string')
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            value += '"'
            j += 2
            continue
          }
          break
        }
        value += src[j++]
      }
      tokens.push({ type: 'str', text: value, start: i, end: j + 1 })
      i = j + 1
      continue
    }
    if (ch === '#') {
      const m = ERROR_RE.exec(src.slice(i))
      if (!m) throw new TokenizeError('Unknown error literal')
      tokens.push({ type: 'err', text: m[0].toUpperCase(), start: i, end: i + m[0].length })
      i += m[0].length
      continue
    }
    if (ch === "'" || ch === '$' || /[A-Za-z0-9_]/.test(ch)) {
      const ref = tryRef(src, i)
      if (ref) {
        tokens.push(ref)
        i = ref.end
        continue
      }
    }
    if (/[0-9.]/.test(ch)) {
      const m = NUMBER_RE.exec(src.slice(i))
      if (m) {
        tokens.push({ type: 'num', text: m[0], start: i, end: i + m[0].length })
        i += m[0].length
        continue
      }
    }
    if (/[A-Za-z_\\]/.test(ch)) {
      const m = NAME_RE.exec(src.slice(i))!
      const j = i + m[0].length
      let k = j
      while (k < n && src[k] === ' ') k++
      const isFunc = src[k] === '('
      tokens.push({ type: isFunc ? 'func' : 'name', text: m[0], start: i, end: j })
      i = j
      continue
    }
    const two = src.slice(i, i + 2)
    if (two === '<=' || two === '>=' || two === '<>') {
      tokens.push({ type: 'op', text: two, start: i, end: i + 2 })
      i += 2
      continue
    }
    if ('+-*/^&=<>%'.includes(ch)) {
      tokens.push({ type: 'op', text: ch, start: i, end: i + 1 })
      i++
      continue
    }
    if ('(),;{}'.includes(ch)) {
      tokens.push({ type: ch as TokenType, text: ch, start: i, end: i + 1 })
      i++
      continue
    }
    throw new TokenizeError(`Unexpected character "${ch}"`)
  }
  return tokens
}

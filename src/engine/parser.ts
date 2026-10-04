import { tokenize, type RefToken, type Token } from './tokenizer.js'

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: string }
  | { t: 'ref'; ref: RefToken }
  | { t: 'name'; name: string }
  | { t: 'unary'; op: '-' | '+'; a: Node }
  | { t: 'pct'; a: Node }
  | { t: 'bin'; op: string; a: Node; b: Node }
  | { t: 'call'; name: string; args: Node[] }
  | { t: 'array'; rows: Node[][] }
  | { t: 'empty' }

export class ParseError extends Error {}

const BINARY_PRECEDENCE: Record<string, number> = {
  '=': 1,
  '<>': 1,
  '<': 1,
  '>': 1,
  '<=': 1,
  '>=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
  '^': 5
}
const UNARY_PRECEDENCE = 6

/** Parse a formula body (text after "="). */
export function parseFormula (src: string): Node {
  const tokens = tokenize(src).filter((t) => t.type !== 'ws')
  let pos = 0

  const peek = (): Token | undefined => tokens[pos]
  const next = (): Token => {
    const t = tokens[pos++]
    if (!t) throw new ParseError('Unexpected end of formula')
    return t
  }
  const expect = (type: string): Token => {
    const t = next()
    if (t.type !== type) throw new ParseError(`Expected "${type}" but found "${t.text}"`)
    return t
  }

  function parseExpr (minPrec: number): Node {
    let left = parsePrefix()
    for (;;) {
      const t = peek()
      if (!t || t.type !== 'op') break
      if (t.text === '%') {
        pos++
        left = { t: 'pct', a: left }
        continue
      }
      const prec = BINARY_PRECEDENCE[t.text]
      if (prec === undefined || prec < minPrec) break
      pos++
      // ^ is left-associative in Excel, like everything else.
      const right = parseExpr(prec + 1)
      left = { t: 'bin', op: t.text, a: left, b: right }
    }
    return left
  }

  function parsePrefix (): Node {
    const t = next()
    switch (t.type) {
      case 'num':
        return { t: 'num', v: Number(t.text) }
      case 'str':
        return { t: 'str', v: t.text }
      case 'err':
        return { t: 'err', v: t.text }
      case 'ref':
        return { t: 'ref', ref: t.ref! }
      case 'name': {
        const u = t.text.toUpperCase()
        if (u === 'TRUE') return { t: 'bool', v: true }
        if (u === 'FALSE') return { t: 'bool', v: false }
        return { t: 'name', name: t.text }
      }
      case 'func':
        return parseCall(t.text)
      case '(': {
        const e = parseExpr(0)
        expect(')')
        return e
      }
      case '{':
        return parseArray()
      case 'op':
        if (t.text === '-' || t.text === '+') {
          const a = parseExpr(UNARY_PRECEDENCE)
          return { t: 'unary', op: t.text, a }
        }
        break
    }
    throw new ParseError(`Unexpected "${t.text}"`)
  }

  function parseCall (name: string): Node {
    expect('(')
    const args: Node[] = []
    if (peek()?.type === ')') {
      pos++
      return { t: 'call', name: name.toUpperCase(), args }
    }
    for (;;) {
      const t = peek()
      if (t && (t.type === ',' || t.type === ')')) args.push({ t: 'empty' })
      else args.push(parseExpr(0))
      const sep = next()
      if (sep.type === ')') break
      if (sep.type !== ',') throw new ParseError(`Expected "," or ")" but found "${sep.text}"`)
      if (peek()?.type === ')') {
        args.push({ t: 'empty' })
        pos++
        break
      }
    }
    return { t: 'call', name: name.toUpperCase(), args }
  }

  function parseArray (): Node {
    const rows: Node[][] = [[]]
    for (;;) {
      const e = parseExpr(0)
      rows[rows.length - 1].push(e)
      const t = next()
      if (t.type === '}') break
      if (t.type === ';') rows.push([])
      else if (t.type !== ',') throw new ParseError('Bad array constant')
    }
    return { t: 'array', rows }
  }

  if (tokens.length === 0) throw new ParseError('Empty formula')
  const root = parseExpr(0)
  if (pos < tokens.length) throw new ParseError(`Unexpected "${tokens[pos].text}"`)
  return root
}

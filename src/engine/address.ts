/** A1-style addressing helpers. Rows and columns are 0-based internally. */

export const MAX_ROWS = 1_048_576
export const MAX_COLS = 16_384

const colNameCache: string[] = []

export function colToName (c: number): string {
  const cached = colNameCache[c]
  if (cached !== undefined) return cached
  let n = c + 1
  let s = ''
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  if (c < 2048) colNameCache[c] = s
  return s
}

/** "A" → 0, "AA" → 26. Returns -1 for invalid names. */
export function nameToCol (name: string): number {
  let n = 0
  for (let i = 0; i < name.length; i++) {
    const code = name.charCodeAt(i) | 0x20 // lowercase
    if (code < 97 || code > 122) return -1
    n = n * 26 + (code - 96)
  }
  return n - 1
}

export interface CellAddress {
  row: number
  col: number
}

export interface RangeAddress {
  r1: number
  c1: number
  r2: number
  c2: number
}

export function formatAddress (row: number, col: number): string {
  return colToName(col) + (row + 1)
}

export function formatRange (r: RangeAddress): string {
  const a = formatAddress(r.r1, r.c1)
  if (r.r1 === r.r2 && r.c1 === r.c2) return a
  return a + ':' + formatAddress(r.r2, r.c2)
}

const A1_RE = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/

export function parseAddress (s: string): CellAddress | undefined {
  const m = A1_RE.exec(s.trim())
  if (!m) return undefined
  const col = nameToCol(m[1])
  const row = Number(m[2]) - 1
  if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) return undefined
  return { row, col }
}

/** Parse "B3" or "A1:C10" (as typed into the name box). */
export function parseRangeAddress (s: string): RangeAddress | undefined {
  const parts = s.trim().split(':')
  if (parts.length === 1) {
    const a = parseAddress(parts[0])
    return a && { r1: a.row, c1: a.col, r2: a.row, c2: a.col }
  }
  if (parts.length !== 2) return undefined
  const a = parseAddress(parts[0])
  const b = parseAddress(parts[1])
  if (!a || !b) return undefined
  return normalizeRange({ r1: a.row, c1: a.col, r2: b.row, c2: b.col })
}

export function normalizeRange (r: RangeAddress): RangeAddress {
  return {
    r1: Math.min(r.r1, r.r2),
    c1: Math.min(r.c1, r.c2),
    r2: Math.max(r.r1, r.r2),
    c2: Math.max(r.c1, r.c2)
  }
}

/** Quote a sheet name for use in a formula when needed. */
export function quoteSheetName (name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
    ? name
    : "'" + name.replace(/'/g, "''") + "'"
}

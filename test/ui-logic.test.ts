import { describe, expect, it } from 'vitest'
import { parseAddress } from '../src/engine/address'
import { valueToRaw } from '../src/engine/format'
import { normalizeFormula } from '../src/engine/refs'
import { Errors } from '../src/engine/values'
import { Workbook } from '../src/engine/workbook'
import { innermostCall } from '../src/ui/editor'
import { fillPatches } from '../src/ui/fill'
import { History } from '../src/ui/history'
import { csvSafe, detectDelimiter, escapeHtml, parseDelimited, toDelimited } from '../src/ui/io'

function sheetWith (cells: Record<string, string>, styles: Record<string, { fmt: string }> = {}) {
  const wb = new Workbook()
  const sheet = wb.sheets[0]
  wb.batch(() => {
    for (const [addr, raw] of Object.entries(cells)) {
      const a = parseAddress(addr)!
      wb.applyPatches(sheet.id, [{ r: a.row, c: a.col, raw, style: styles[addr] ?? null }])
    }
  })
  return { wb, sheet }
}

const rawsOf = (patches: { r: number; c: number; raw?: string }[]) => patches.map((p) => p.raw)

describe('fillPatches', () => {
  it('continues a linear series down', () => {
    const { sheet } = sheetWith({ A1: '1', A2: '3' })
    expect(rawsOf(fillPatches(sheet, { r1: 0, c1: 0, r2: 1, c2: 0 }, { r1: 0, c1: 0, r2: 4, c2: 0 }))).toEqual(['5', '7', '9'])
  })

  it('continues a series upward and to the left', () => {
    const { sheet } = sheetWith({ A5: '10', A6: '20', E1: '5', F1: '6' })
    expect(rawsOf(fillPatches(sheet, { r1: 4, c1: 0, r2: 5, c2: 0 }, { r1: 2, c1: 0, r2: 5, c2: 0 }))).toEqual(['-10', '0'])
    expect(rawsOf(fillPatches(sheet, { r1: 0, c1: 4, r2: 0, c2: 5 }, { r1: 0, c1: 2, r2: 0, c2: 5 }))).toEqual(['3', '4'])
  })

  it('copies a single number but counts dates by day', () => {
    const { sheet } = sheetWith({ A1: '7', B1: '2024-01-31' }, { B1: { fmt: 'yyyy-mm-dd' } })
    expect(rawsOf(fillPatches(sheet, { r1: 0, c1: 0, r2: 0, c2: 0 }, { r1: 0, c1: 0, r2: 2, c2: 0 }))).toEqual(['7', '7'])
    expect(rawsOf(fillPatches(sheet, { r1: 0, c1: 1, r2: 0, c2: 1 }, { r1: 0, c1: 1, r2: 2, c2: 1 }))).toEqual(['45323', '45324'])
  })

  it('counts "Item 7"-style text and repeats plain text patterns', () => {
    const { sheet } = sheetWith({ A1: 'Q1', B1: 'x', B2: 'y' })
    expect(rawsOf(fillPatches(sheet, { r1: 0, c1: 0, r2: 0, c2: 0 }, { r1: 0, c1: 0, r2: 2, c2: 0 }))).toEqual(['Q2', 'Q3'])
    expect(rawsOf(fillPatches(sheet, { r1: 0, c1: 1, r2: 1, c2: 1 }, { r1: 0, c1: 1, r2: 4, c2: 1 }))).toEqual(['x', 'y', 'x'])
  })

  it('shifts relative references in formulas', () => {
    const { sheet } = sheetWith({ A1: '1', B1: '=A1*$C$1' })
    expect(rawsOf(fillPatches(sheet, { r1: 0, c1: 1, r2: 0, c2: 1 }, { r1: 0, c1: 1, r2: 2, c2: 1 }))).toEqual(['=A2*$C$1', '=A3*$C$1'])
  })

  it('copies without series detection when asked (Ctrl+D)', () => {
    const { sheet } = sheetWith({ A1: '1', A2: '2' })
    expect(rawsOf(fillPatches(sheet, { r1: 0, c1: 0, r2: 0, c2: 0 }, { r1: 0, c1: 0, r2: 2, c2: 0 }, true))).toEqual(['1', '1'])
  })

  it('carries styles along', () => {
    const { sheet } = sheetWith({ A1: '1' }, { A1: { fmt: '0.00' } })
    expect(fillPatches(sheet, { r1: 0, c1: 0, r2: 0, c2: 0 }, { r1: 0, c1: 0, r2: 1, c2: 0 })[0].style).toEqual({ fmt: '0.00' })
  })
})

describe('delimited text', () => {
  it('parses quotes, escaped quotes, embedded newlines and CRLF', () => {
    expect(parseDelimited('a,"b,c","say ""hi"""\r\n"multi\nline",x', ',')).toEqual([['a', 'b,c', 'say "hi"'], ['multi\nline', 'x']])
  })

  it('treats an unclosed quote as an ordinary character', () => {
    expect(parseDelimited('"oops\tb\nc', '\t')).toEqual([['"oops', 'b'], ['c']])
  })

  it('keeps empty trailing fields and a lone "" row', () => {
    expect(parseDelimited('a,,\n""', ',')).toEqual([['a', '', ''], ['']])
  })

  it('detects tabs, semicolons and commas', () => {
    expect(detectDelimiter('a\tb\tc\n1\t2\t3')).toBe('\t')
    expect(detectDelimiter('a;b;c')).toBe(';')
    expect(detectDelimiter('a,b')).toBe(',')
    expect(detectDelimiter('a'.repeat(5000) + ',b')).toBe(',')
  })

  it('round-trips through toDelimited', () => {
    const rows = [['plain', 'with,comma', 'with "quote"'], ['line\nbreak', '', 'x']]
    expect(parseDelimited(toDelimited(rows, ','), ',')).toEqual(rows)
    expect(parseDelimited(toDelimited(rows, '\t'), '\t')).toEqual(rows)
  })

  it('neutralises formula-like text for CSV export but keeps numbers', () => {
    expect(csvSafe('=HYPERLINK("x")')).toBe("'=HYPERLINK(\"x\")")
    expect(csvSafe('@SUM(A1)')).toBe("'@SUM(A1)")
    expect(csvSafe('-5')).toBe('-5')
    expect(csvSafe('+1.5')).toBe('+1.5')
    expect(csvSafe('hello')).toBe('hello')
  })

  it('escapes HTML', () => {
    expect(escapeHtml('<a href="x">&</a>')).toBe('&lt;a href=&quot;x&quot;&gt;&amp;&lt;/a&gt;')
  })
})

describe('History', () => {
  const entry = (n: number) => ({ kind: 'size' as const, sheetId: 1, axis: 'cols' as const, index: n, before: 1, after: 2 })

  it('undoes and redoes in order, and a new push clears redo', () => {
    const h = new History()
    h.push(entry(1))
    h.push(entry(2))
    expect(h.undo()).toMatchObject({ index: 2 })
    expect(h.canRedo).toBe(true)
    expect(h.redo()).toMatchObject({ index: 2 })
    h.undo()
    h.push(entry(3))
    expect(h.canRedo).toBe(false)
    expect(h.undo()).toMatchObject({ index: 3 })
    expect(h.undo()).toMatchObject({ index: 1 })
    expect(h.undo()).toBeUndefined()
  })

  it('keeps at most 200 entries', () => {
    const h = new History()
    for (let i = 0; i < 250; i++) h.push(entry(i))
    let n = 0
    while (h.undo()) n++
    expect(n).toBe(200)
  })

  it('groups the last entries into one undo step', () => {
    const h = new History()
    h.push(entry(1))
    h.push(entry(2))
    h.push(entry(3))
    h.group(2)
    expect(h.undo()).toMatchObject({ kind: 'group', entries: [{ index: 2 }, { index: 3 }] })
    expect(h.undo()).toMatchObject({ index: 1 })
  })
})

describe('formula helpers', () => {
  it('normalises typed formulas like Excel', () => {
    expect(normalizeFormula('=sum(a1:b2')).toBe('=SUM(A1:B2)')
    expect(normalizeFormula('=if(true,"keep case",c$3)')).toBe('=IF(TRUE,"keep case",C$3)')
    expect(normalizeFormula("='my sheet'!a1+1")).toBe("='my sheet'!A1+1")
    expect(normalizeFormula('=round( (1+2) , 1')).toBe('=ROUND( (1+2) , 1)')
    expect(normalizeFormula('="unterminated')).toBe('="unterminated')
  })

  it('finds the innermost call and argument under the caret', () => {
    expect(innermostCall('=SUM(')).toEqual({ name: 'SUM', arg: 0 })
    expect(innermostCall('=IF(A1>0, ROUND(B1, ')).toEqual({ name: 'ROUND', arg: 1 })
    expect(innermostCall('=IF(A1, MAX(1,2), ')).toEqual({ name: 'IF', arg: 2 })
    expect(innermostCall('=CONCAT("(", ')).toEqual({ name: 'CONCAT', arg: 1 })
    expect(innermostCall('=1+2')).toBeNull()
  })

  it('turns values back into raw input that reproduces them', () => {
    expect(valueToRaw(null)).toBe('')
    expect(valueToRaw(1.5)).toBe('1.5')
    expect(valueToRaw(true)).toBe('TRUE')
    expect(valueToRaw(Errors.NA)).toBe('#N/A')
    expect(valueToRaw('hello')).toBe('hello')
    expect(valueToRaw('=not a formula')).toBe("'=not a formula")
    expect(valueToRaw('42')).toBe("'42")
  })
})

describe('workbook API', () => {
  it('returns inverse patches that undo an edit exactly', () => {
    const { wb, sheet } = sheetWith({ A1: '1', B1: '=A1*2' })
    const undo = wb.applyPatches(sheet.id, [{ r: 0, c: 0, raw: '5', style: { bold: true } }])
    expect(wb.getValue(sheet.id, 0, 1)).toBe(10)
    wb.applyPatches(sheet.id, undo)
    expect(wb.getCell(sheet.id, 0, 0)!.style).toBeNull()
    expect(wb.getValue(sheet.id, 0, 1)).toBe(2)
  })

  it('sorts rows and moves formulas with them', () => {
    const { wb, sheet } = sheetWith({ A1: '3', A2: '1', A3: '2', B1: '=A1*10', B2: '=A2*10', B3: '=A3*10' })
    wb.applyPatches(sheet.id, wb.sortPatches(sheet.id, 0, 0, 2, 1, 0, true))
    expect([0, 1, 2].map((r) => wb.getValue(sheet.id, r, 1))).toEqual([10, 20, 30])
    wb.applyPatches(sheet.id, wb.sortPatches(sheet.id, 0, 0, 2, 1, 0, false))
    expect([0, 1, 2].map((r) => wb.getValue(sheet.id, r, 0))).toEqual([3, 2, 1])
  })

  it('treats text-formatted cells as text, even if they look like formulas', () => {
    const { wb, sheet } = sheetWith({ A1: '=1+1', A2: '007' }, { A1: { fmt: '@' }, A2: { fmt: '@' } })
    expect(wb.getValue(sheet.id, 0, 0)).toBe('=1+1')
    expect(wb.getValue(sheet.id, 1, 0)).toBe('007')
    wb.applyPatches(sheet.id, [{ r: 0, c: 0, style: null }])
    expect(wb.getValue(sheet.id, 0, 0)).toBe(2)
  })

  it('applies typed-format hints only when asked', () => {
    const { wb, sheet } = sheetWith({})
    wb.applyPatches(sheet.id, [{ r: 0, c: 0, raw: '12%' }], true)
    wb.applyPatches(sheet.id, [{ r: 1, c: 0, raw: '12%' }], false)
    expect(wb.getCell(sheet.id, 0, 0)!.style?.fmt).toBe('0%')
    expect(wb.getCell(sheet.id, 1, 0)!.style).toBeNull()
  })

  it('recalculates volatile functions on every recalc', () => {
    const { wb, sheet } = sheetWith({ A1: '=RAND()' })
    const seen = new Set<unknown>()
    for (let i = 0; i < 5; i++) {
      wb.recalc()
      seen.add(wb.getValue(sheet.id, 0, 0))
    }
    expect(seen.size).toBeGreaterThan(1)
  })

  it('turns references to a removed sheet into #REF! and restores them when it returns', () => {
    const wb = new Workbook()
    const main = wb.sheets[0].id
    const data = wb.addSheet('Data')
    wb.setCell(data.id, 0, 0, '7')
    wb.setCell(main, 0, 0, '=Data!A1')
    wb.removeSheet(data.id)
    expect(String(wb.getValue(main, 0, 0))).toBe('#REF!')
    wb.addSheet('Data')
    expect(wb.getValue(main, 0, 0)).toBe(0)
  })

  it('reorders sheets and rejects bad names', () => {
    const wb = new Workbook()
    const b = wb.addSheet('B')
    wb.moveSheet(b.id, 0)
    expect(wb.sheets.map((s) => s.name)).toEqual(['B', 'Sheet1'])
    expect(wb.renameSheet(b.id, 'Sheet1')).toBe(false)
    expect(wb.renameSheet(b.id, 'a/b')).toBe(false)
    expect(wb.renameSheet(b.id, '  ')).toBe(false)
    expect(wb.renameSheet(b.id, 'Ok')).toBe(true)
  })

  it('dedupes sheet names and ids when loading', () => {
    const wb = new Workbook()
    wb.load({ version: 1, sheets: [{ id: 3, name: 'X', cells: [] }, { id: 3, name: 'x', cells: [] }] })
    expect(wb.sheets.map((s) => s.name)).toEqual(['X', 'x (2)'])
    expect(new Set(wb.sheets.map((s) => s.id)).size).toBe(2)
  })

  it('notifies listeners and stops after unsubscribe', () => {
    const wb = new Workbook()
    let n = 0
    const off = wb.onChange(() => n++)
    wb.setCell(wb.sheets[0].id, 0, 0, '1')
    off()
    wb.setCell(wb.sheets[0].id, 0, 0, '2')
    expect(n).toBe(1)
  })

  it('reports INDIRECT errors and resolves sheet-qualified text', () => {
    const wb = new Workbook()
    const main = wb.sheets[0].id
    const other = wb.addSheet('Other')
    wb.setCell(other.id, 1, 1, '9')
    wb.setCell(main, 0, 0, '=INDIRECT("Other!B2")')
    wb.setCell(main, 1, 0, '=INDIRECT("not a ref")')
    wb.setCell(main, 2, 0, '=INDIRECT("Nope!A1")')
    expect(wb.getValue(main, 0, 0)).toBe(9)
    expect(String(wb.getValue(main, 1, 0))).toBe('#REF!')
    expect(String(wb.getValue(main, 2, 0))).toBe('#REF!')
  })
})

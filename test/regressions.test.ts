import { describe, expect, it } from 'vitest'
import { parseAddress } from '../src/engine/address'
import { Workbook } from '../src/engine/workbook'

/** Build a one-sheet workbook from cells entered in the given order. */
function book (cells: [string, string][]) {
  const wb = new Workbook()
  const id = wb.sheets[0].id
  const at = (addr: string) => parseAddress(addr)!
  wb.batch(() => {
    for (const [addr, raw] of cells) wb.setCell(id, at(addr).row, at(addr).col, raw)
  })
  const get = (addr: string) => {
    const v = wb.getValue(id, at(addr).row, at(addr).col)
    return typeof v === 'object' && v !== null ? String(v) : v
  }
  const raw = (addr: string) => wb.getCell(id, at(addr).row, at(addr).col)?.raw
  const set = (addr: string, v: string) => wb.setCell(id, at(addr).row, at(addr).col, v)
  return { wb, id, get, raw, set }
}

describe('regressions', () => {
  it('keeps reversed ranges intact when unrelated rows are deleted', () => {
    const b = book([['A2', '1'], ['A5', '2'], ['B1', '=SUM(A5:A2)']])
    b.wb.deleteRows(b.id, 19, 1)
    expect(b.raw('B1')).toBe('=SUM(A5:A2)')
    b.wb.deleteRows(b.id, 2, 1) // row 3
    expect(b.raw('B1')).toBe('=SUM(A2:A4)')
    expect(b.get('B1')).toBe(3)
  })

  it('turns references pushed off the grid into #REF!', () => {
    const b = book([['A1', '=A1048576']])
    b.wb.insertRows(b.id, 5, 1)
    expect(b.raw('A1')).toBe('=#REF!')
  })

  it('joins ranges in row order regardless of entry order', () => {
    const b = book([['B2', 'd'], ['A2', 'c'], ['B1', 'b'], ['A1', 'a'], ['C1', '=TEXTJOIN(",",TRUE,A1:B2)'], ['C2', '=CONCAT(A1:B2)']])
    expect(b.get('C1')).toBe('a,b,c,d')
    expect(b.get('C2')).toBe('abcd')
  })

  it('computes NPV in sheet order', () => {
    const b = book([['A3', '3'], ['A1', '1'], ['A2', '2'], ['B1', '=NPV(0.1,A1:A3)']])
    expect(b.get('B1')).toBeCloseTo(1 / 1.1 + 2 / 1.1 ** 2 + 3 / 1.1 ** 3, 10)
  })

  it('recalculates SUMIF when cells of an auto-sized sum range change', () => {
    const b = book([['A1', '1'], ['A2', '1'], ['A3', '1'], ['B1', '10'], ['B2', '20'], ['B3', '30'], ['C1', '=SUMIF(A1:A3,">0",B1)']])
    expect(b.get('C1')).toBe(60)
    b.set('B2', '200')
    expect(b.get('C1')).toBe(240)
  })

  it('rejects mismatched shapes in the *IFS family', () => {
    const b = book([['A1', '1'], ['A2', '1'], ['A3', '1'], ['B3', '5'], ['C1', '=SUMIFS(B1:B2,A1:A3,">0")']])
    expect(b.get('C1')).toBe('#VALUE!')
  })

  it('does not report a false cycle from an untaken IF branch', () => {
    const b = book([['D1', '1'], ['A1', '=IF(D1*0+B1,1,C1)'], ['B1', '=D1>0'], ['C1', '=A1+1']])
    expect(b.get('A1')).toBe(1)
    expect(b.get('C1')).toBe(2)
  })

  it('counts blank-matching criteria over whole columns quickly and correctly', () => {
    const b = book([['A1', '5'], ['A2', '6'], ['B1', '=COUNTIF(A:A,"<>5")'], ['B2', '=COUNTIF(A1:A10,"")']])
    expect(b.get('B1')).toBe(1_048_575)
    expect(b.get('B2')).toBe(8)
    const t0 = performance.now()
    for (let i = 3; i < 103; i++) b.set(`C${i}`, '=COUNTIF(A:A,"<>5")')
    expect(performance.now() - t0).toBeLessThan(500)
  })

  it('does not match blank lookup values, and XLOOKUP exact match ignores wildcards', () => {
    const b = book([['A1', 'abc'], ['A2', 'a*'], ['B1', 'hit'], ['B2', 'exact'], ['C1', '=XLOOKUP("a*",A1:A2,B1:B2)'], ['C2', '=MATCH(D9,A1:A3,0)']])
    expect(b.get('C1')).toBe('exact')
    expect(b.get('C2')).toBe('#N/A')
  })

  it('keeps sheet ids stable through a JSON round trip', () => {
    const wb = new Workbook()
    const second = wb.addSheet('Second')
    wb.moveSheet(second.id, 0)
    wb.load(wb.toJSON())
    expect(wb.sheetByName('Second')!.id).toBe(second.id)
  })

  it('leaves the workbook untouched when loading invalid JSON', () => {
    const b = book([['A1', '42']])
    expect(() => b.wb.load({ name: 'package' } as never)).toThrow(/Not a FilmSheet workbook/)
    expect(b.get('A1')).toBe(42)
  })

  it('dirty marking stays linear with many range listeners on one column', () => {
    const wb = new Workbook()
    const id = wb.sheets[0].id
    const n = 16_000
    wb.batch(() => {
      wb.setCell(id, 0, 0, '1')
      for (let r = 1; r < n; r++) wb.setCell(id, r, 0, `=A${r}+1`)
      for (let r = 0; r < n; r++) wb.setCell(id, r, 1, `=ROWS(A1:A${n})`)
    })
    const t0 = performance.now()
    wb.setCell(id, 0, 0, '2')
    expect(performance.now() - t0).toBeLessThan(250)
    expect(wb.getValue(id, n - 1, 0)).toBe(n + 1)
  })
  it('resolves a reference to a sheet added later', () => {
    const wb = new Workbook()
    const s1 = wb.sheets[0].id
    wb.setCell(s1, 0, 0, '=Later!A1+1')
    expect(String(wb.getValue(s1, 0, 0))).toBe('#REF!')
    const later = wb.addSheet('Later')
    wb.setCell(later.id, 0, 0, '41')
    expect(wb.getValue(s1, 0, 0)).toBe(42)
  })
  it('handles function and format edge cases like Excel', () => {
    const b = book([
      ['A1', '5'], ['A2', "'5"], ['A3', '1'], ['A4', '2'], ['A5', '3'], ['B1', '1'], ['B2', '2'], ['B3', '3'],
      ['C1', '=COUNTIF(A1:A2,"5")'], ['C2', '=COUNTIF(A1:A2,"=5")'], ['C3', '=COUNTIF(A1:A2,5)'],
      ['C4', '=CORREL(A3:A5,B1:D1)'], ['C5', '=SUBSTITUTE("abc","b","x",0)'], ['C6', '=FIND("a","abc",0)'],
      ['C7', '=REPLACE("abc",0,1,"x")'], ['C8', '=LOG10 (100)'], ['C9', '=PRODUCT(Z1:Z5)'], ['C10', '=ROUND(1E300,10)'],
      ['C11', '=TIME(-1,0,0)'], ['C12', '=CHAR(300)'], ['C13', '=1/10000000']
    ])
    b.set('D1', '1'); b.set('C1', '=COUNTIF(A1:A2,"5")')
    expect([b.get('C1'), b.get('C2'), b.get('C3')]).toEqual([2, 2, 2])
    expect(b.get('C5')).toBe('#VALUE!')
    expect(b.get('C6')).toBe('#VALUE!')
    expect(b.get('C7')).toBe('#VALUE!')
    expect(b.get('C8')).toBe(2)
    expect(b.get('C9')).toBe(0)
    expect(b.get('C10')).toBe(1e300)
    expect(b.get('C11')).toBe('#NUM!')
    expect(b.get('C12')).toBe('#VALUE!')
    const cell = b.wb.getCell(b.id, 12, 2)!
    expect(b.wb.getDisplay(cell)).toBe('1E-07')
  })

  it('pairs CORREL values by position across shapes', () => {
    const b = book([['A1', '1'], ['B1', '2'], ['C1', '3'], ['A2', '1'], ['A3', '2'], ['A4', '3'], ['D1', '=CORREL(A1:C1,A2:A4)']])
    expect(b.get('D1')).toBeCloseTo(1, 10)
  })
})

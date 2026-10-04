import { describe, expect, it } from 'vitest'
import { Workbook } from '../src/engine/workbook'
import { shiftFormula, transformFormulaForStructure } from '../src/engine/refs'
import { formatValue, parseInput } from '../src/engine/format'
import { parseAddress, colToName, nameToCol } from '../src/engine/address'

function sheet (cells: Record<string, string>) {
  const wb = new Workbook()
  const id = wb.sheets[0].id
  wb.batch(() => {
    for (const [addr, raw] of Object.entries(cells)) {
      const a = parseAddress(addr)!
      wb.setCell(id, a.row, a.col, raw)
    }
  })
  const get = (addr: string) => {
    const a = parseAddress(addr)!
    return wb.getValue(id, a.row, a.col)
  }
  const set = (addr: string, raw: string) => {
    const a = parseAddress(addr)!
    wb.setCell(id, a.row, a.col, raw)
  }
  const display = (addr: string) => {
    const a = parseAddress(addr)!
    const c = wb.getCell(id, a.row, a.col)
    return c ? wb.getDisplay(c) : ''
  }
  return { wb, id, get, set, display }
}

const evalOne = (formula: string, cells: Record<string, string> = {}) => {
  const s = sheet({ ...cells, Z100: formula })
  const v = s.get('Z100')
  return v instanceof Object && 'code' in v ? (v as { code: string }).code : v
}

describe('addresses', () => {
  it('converts columns', () => {
    expect(colToName(0)).toBe('A')
    expect(colToName(25)).toBe('Z')
    expect(colToName(26)).toBe('AA')
    expect(colToName(16383)).toBe('XFD')
    expect(nameToCol('XFD')).toBe(16383)
  })
})

describe('arithmetic and precedence', () => {
  it.each([
    ['=1+2*3', 7],
    ['=(1+2)*3', 9],
    ['=2^3^2', 64],
    ['=-2^2', 4],
    ['=10/4', 2.5],
    ['=50%', 0.5],
    ['=1/0', '#DIV/0!'],
    ['="a"&"b"&1', 'ab1'],
    ['=1<2', true],
    ['="abc"="ABC"', true],
    ['=1+"2"', 3],
    ['=1+"x"', '#VALUE!'],
    ['=FOO(1)', '#NAME?'],
    ['=TRUE+1', 2]
  ])('%s → %s', (f, expected) => {
    expect(evalOne(f)).toBe(expected)
  })
})

describe('functions', () => {
  const data = { A1: '1', A2: '2', A3: '3', A4: 'x', A5: '', B1: 'apple', B2: 'banana', B3: 'cherry', C1: '10', C2: '20', C3: '30' }
  it.each([
    ['=SUM(A1:A5)', 6],
    ['=SUM(A1:A3, 10)', 16],
    ['=AVERAGE(A1:A4)', 2],
    ['=COUNT(A1:A5)', 3],
    ['=COUNTA(A1:A5)', 4],
    ['=MIN(A1:A3)', 1],
    ['=MAX(A:A)', 3],
    ['=IF(A1>0,"pos","neg")', 'pos'],
    ['=IFERROR(1/0, "oops")', 'oops'],
    ['=VLOOKUP("banana", B1:C3, 2, FALSE)', 20],
    ['=INDEX(C1:C3, MATCH("cherry", B1:B3, 0))', 30],
    ['=XLOOKUP("apple", B1:B3, C1:C3)', 10],
    ['=XLOOKUP("kiwi", B1:B3, C1:C3, "none")', 'none'],
    ['=SUMIF(A1:A3, ">1", C1:C3)', 50],
    ['=COUNTIF(B1:B3, "*an*")', 1],
    ['=SUMIFS(C1:C3, A1:A3, ">=2", B1:B3, "<>cherry")', 20],
    ['=ROUND(2.675, 2)', 2.68],
    ['=LEFT("hello", 2)&UPPER(MID("hello",3,3))', 'heLLO'],
    ['=TEXTJOIN(", ", TRUE, B1:B3)', 'apple, banana, cherry'],
    ['=SUMPRODUCT(A1:A3, C1:C3)', 140],
    ['=AND(TRUE, A1=1)', true],
    ['=CHOOSE(2, "a", "b", "c")', 'b'],
    ['=DATE(2024, 1, 31)', 45322],
    ['=YEAR(45322)&"-"&MONTH(45322)', '2024-1'],
    ['=TEXT(0.256, "0.0%")', '25.6%'],
    ['=ROWS(A1:B10)*COLUMNS(A1:B10)', 20],
    ['=SUM(INDIRECT("C1:C2"))', 30],
    ['=SUM(OFFSET(C1, 1, 0, 2, 1))', 50],
    ['=MEDIAN(1, 5, 3, 9)', 4],
    ['=MOD(-3, 2)', 1],
    ['=STDEV.P(2,4,4,4,5,5,7,9)', 2]
  ])('%s → %s', (f, expected) => {
    expect(evalOne(f, data)).toBe(expected)
  })

  it('reports arity errors at compile time', () => {
    const s = sheet({ A1: '=ROUND()' })
    expect(String(s.get('A1'))).toBe('#ERROR!')
    expect(s.wb.getCell(s.id, 0, 0)!.error).toMatch(/ROUND expects/)
  })
})

describe('recalculation', () => {
  it('propagates edits through chains', () => {
    const s = sheet({ A1: '1', A2: '=A1*2', A3: '=A2+A1', A4: '=SUM(A1:A3)' })
    expect(s.get('A4')).toBe(6)
    s.set('A1', '10')
    expect(s.get('A2')).toBe(20)
    expect(s.get('A3')).toBe(30)
    expect(s.get('A4')).toBe(60)
  })

  it('handles formulas entered before their inputs', () => {
    const s = sheet({ A3: '=A2+1', A2: '=A1+1', A1: '5' })
    expect(s.get('A3')).toBe(7)
  })

  it('detects cycles and recovers when they are broken', () => {
    const s = sheet({ A1: '=B1+1', B1: '=A1+1' })
    expect(String(s.get('A1'))).toBe('#CIRC!')
    expect(String(s.get('B1'))).toBe('#CIRC!')
    s.set('B1', '5')
    expect(s.get('A1')).toBe(6)
  })

  it('detects self references', () => {
    const s = sheet({ A1: '=A1' })
    expect(String(s.get('A1'))).toBe('#CIRC!')
  })

  it('updates range dependents when cells inside the range change', () => {
    const s = sheet({ B1: '=SUM(A:A)' })
    expect(s.get('B1')).toBe(0)
    s.set('A500', '7')
    expect(s.get('B1')).toBe(7)
    s.set('A500', '')
    expect(s.get('B1')).toBe(0)
  })

  it('handles a 100k-long dependency chain without recursion', () => {
    const wb = new Workbook()
    const id = wb.sheets[0].id
    const N = 100_000
    wb.batch(() => {
      wb.setCell(id, 0, 0, '1')
      // enter in reverse so evaluation has to discover the order
      for (let r = N - 1; r >= 1; r--) wb.setCell(id, r, 0, `=A${r}+1`)
    })
    expect(wb.getValue(id, N - 1, 0)).toBe(N)
    wb.setCell(id, 0, 0, '10')
    expect(wb.getValue(id, N - 1, 0)).toBe(N + 9)
  })

  it('evaluates cross-sheet references and follows renames', () => {
    const wb = new Workbook()
    const s1 = wb.sheets[0].id
    const s2 = wb.addSheet('Data').id
    wb.setCell(s2, 0, 0, '42')
    wb.setCell(s1, 0, 0, '=Data!A1*2')
    expect(wb.getValue(s1, 0, 0)).toBe(84)
    wb.renameSheet(s2, 'My Data')
    expect(wb.getCell(s1, 0, 0)!.raw).toBe("='My Data'!A1*2")
    wb.setCell(s2, 0, 0, '1')
    expect(wb.getValue(s1, 0, 0)).toBe(2)
  })

  it('round-trips through JSON', () => {
    const s = sheet({ A1: '2', A2: '=A1^10', B1: '12%' })
    const copy = Workbook.fromJSON(JSON.parse(JSON.stringify(s.wb.toJSON())))
    expect(copy.getValue(copy.sheets[0].id, 1, 0)).toBe(1024)
  })
})

describe('structural edits', () => {
  it('shifts references when inserting rows', () => {
    const s = sheet({ A1: '1', A2: '2', A3: '=SUM(A1:A2)', B1: '=A2' })
    s.wb.insertRows(s.id, 1, 2)
    expect(s.wb.getCell(s.id, 4, 0)!.raw).toBe('=SUM(A1:A4)')
    expect(s.wb.getCell(s.id, 0, 1)!.raw).toBe('=A4')
    expect(s.get('A5')).toBe(3)
  })

  it('produces #REF! when a referenced row is deleted', () => {
    const s = sheet({ A1: '1', A2: '2', A3: '=A2*2', A4: '=SUM(A1:A3)' })
    s.wb.deleteRows(s.id, 1, 1)
    expect(s.wb.getCell(s.id, 1, 0)!.raw).toBe('=#REF!*2')
    expect(s.wb.getCell(s.id, 2, 0)!.raw).toBe('=SUM(A1:A2)')
  })

  it('shifts columns', () => {
    const s = sheet({ A1: '1', B1: '2', C1: '=A1+B1' })
    s.wb.insertCols(s.id, 1, 1)
    expect(s.wb.getCell(s.id, 0, 3)!.raw).toBe('=A1+C1')
    expect(s.get('D1')).toBe(3)
  })
})

describe('reference rewriting', () => {
  it('shifts relative references only', () => {
    expect(shiftFormula('=A1+$B$2+C$3+$D4', 1, 1)).toBe('=B2+$B$2+D$3+$D5')
    expect(shiftFormula('=SUM(A1:B2)', 2, 0)).toBe('=SUM(A3:B4)')
    expect(shiftFormula('=A1', -1, 0)).toBe('=#REF!')
    expect(shiftFormula('="A1"&A1', 1, 0)).toBe('="A1"&A2')
  })

  it('expands ranges on insert', () => {
    expect(transformFormulaForStructure('=SUM(A1:A10)', 'rows', 5, 3, () => true)).toBe('=SUM(A1:A13)')
  })
})

describe('input parsing and formatting', () => {
  it('detects typed formats', () => {
    expect(parseInput('12%')).toEqual({ kind: 'value', value: 0.12, format: '0%' })
    expect(parseInput('$1,234.50').format).toBe('$#,##0.00')
    expect(parseInput('2024-03-05').format).toBe('yyyy-mm-dd')
    expect(parseInput("'123").value).toBe('123')
  })

  it('formats numbers', () => {
    expect(formatValue(1234.5, '#,##0.00')).toBe('1,234.50')
    expect(formatValue(-1234.5, '$#,##0.00')).toBe('-$1,234.50')
    expect(formatValue(0.1234, '0.00%')).toBe('12.34%')
    expect(formatValue(45322, 'yyyy-mm-dd')).toBe('2024-01-31')
    expect(formatValue(45322.5, 'yyyy-mm-dd hh:mm')).toBe('2024-01-31 12:00')
    expect(formatValue(12345.678, '0.00E+00')).toBe('1.23E+04')
    expect(formatValue(0.1 + 0.2, undefined)).toBe('0.3')
    expect(formatValue(0, '#,##0.00;(#,##0.00);"-"')).toBe('-')
    expect(formatValue(-5, '#,##0.00;(#,##0.00);"-"')).toBe('(5.00)')
    expect(formatValue(1.5, '0.##')).toBe('1.5')
    expect(formatValue(-5, '0.00;[Red]-0.00')).toBe('-5.00')
    expect(formatValue(1234.5, '#,##0.00_)')).toBe('1,234.50 ')
    expect(formatValue(1.5, 'hh:mm:ss.000')).toBe('12:00:00.000')
    expect(formatValue(45000.99999999, 'yyyy-mm-dd hh:mm:ss')).toBe('2023-03-16 00:00:00')
    expect(parseInput('2024-01-01 25:99').format).toBeUndefined()
  })
})

import { describe, expect, it } from 'vitest'
import { parseAddress } from '../src/engine/address'
import { FUNCTIONS } from '../src/engine/functions'
import { Workbook } from '../src/engine/workbook'

/**
 * Every built-in function, evaluated against one shared fixture. The last
 * test checks that each function in the table is exercised, so adding a
 * function without a test fails CI.
 */
const FIXTURE: Record<string, string> = {
  A1: '1', A2: '2', A3: '3', A4: '4', A5: '5',
  B1: 'apple', B2: 'banana', B3: 'cherry', B4: 'date', B5: 'elder',
  C1: '10', C2: '20', C3: '30', C4: '40', C5: '50',
  D1: 'TRUE', D2: 'x', D4: '=1/0',
  E1: '2', E2: '4', E3: '4', E4: '4', E5: '5',
  F1: '2024-01-31',
  H1: '1', I1: '2', J1: '3', H2: 'a', I2: 'b', J2: 'c'
}

type Expected = number | string | boolean | ((v: unknown) => boolean)

const CASES: [string, Expected][] = [
  // Math & aggregates
  ['=SUM(A1:A5)', 15], ['=PRODUCT(A1:A5)', 120], ['=AVERAGE(A1:A5)', 3], ['=MIN(A1:A5)', 1], ['=MAX(C1:C5)', 50],
  ['=MEDIAN(E1:E5)', 4], ['=MODE(E1:E5)', 4], ['=MODE.SNGL(E1:E5)', 4], ['=COUNT(A1:D5)', 10], ['=COUNTA(D1:D5)', 3],
  ['=COUNTBLANK(D1:D5)', 2], ['=ABS(-3)', 3], ['=SQRT(16)', 4], ['=SQRT(-1)', '#NUM!'], ['=EXP(0)', 1], ['=LN(1)', 0],
  ['=LOG10(1000)', 3], ['=LOG(8,2)', 3], ['=POWER(2,10)', 1024], ['=PI()', Math.PI], ['=SIN(0)', 0], ['=COS(0)', 1],
  ['=TAN(0)', 0], ['=ASIN(1)', Math.PI / 2], ['=ACOS(1)', 0], ['=ATAN(1)', Math.PI / 4], ['=ATAN2(1,1)', Math.PI / 4],
  ['=DEGREES(PI())', 180], ['=RADIANS(180)', Math.PI], ['=SIGN(-5)', -1], ['=INT(-1.5)', -2], ['=TRUNC(-1.57,1)', -1.5],
  ['=ROUND(1.255,2)', 1.26], ['=ROUNDUP(1.21,1)', 1.3], ['=ROUNDDOWN(1.29,1)', 1.2], ['=CEILING(4.3,2)', 6],
  ['=CEILING.MATH(4.3)', 5], ['=FLOOR(4.7,2)', 4], ['=FLOOR.MATH(-4.3)', -5], ['=MROUND(10,3)', 9], ['=MOD(10,3)', 1],
  ['=MOD(1,0)', '#DIV/0!'], ['=QUOTIENT(7,2)', 3], ['=FACT(5)', 120], ['=GCD(12,18)', 6],
  ['=RAND()', (v) => typeof v === 'number' && v >= 0 && v < 1], ['=RANDBETWEEN(5,5)', 5],
  ['=SUMPRODUCT(A1:A5,C1:C5)', 550], ['=SUMSQ(1,2,3)', 14],
  // Conditional aggregates
  ['=SUMIF(A1:A5,">2",C1:C5)', 120], ['=SUMIFS(C1:C5,A1:A5,">1",A1:A5,"<5")', 90], ['=COUNTIF(B1:B5,"*e*")', 4],
  ['=COUNTIFS(A1:A5,">1",C1:C5,"<50")', 3], ['=AVERAGEIF(A1:A5,">3",C1:C5)', 45], ['=AVERAGEIFS(C1:C5,A1:A5,"<=2")', 15],
  ['=MAXIFS(C1:C5,A1:A5,"<4")', 30], ['=MINIFS(C1:C5,A1:A5,">2")', 30],
  // Statistics
  ['=STDEV(2,4,4,4,5,5,7,9)', Math.sqrt(32 / 7)], ['=STDEV.S(2,4,4,4,5,5,7,9)', Math.sqrt(32 / 7)],
  ['=STDEV.P(2,4,4,4,5,5,7,9)', 2], ['=VAR(1,2,3,4)', 5 / 3], ['=VAR.S(1,2,3,4)', 5 / 3], ['=VAR.P(1,2,3,4)', 1.25],
  ['=LARGE(C1:C5,2)', 40], ['=SMALL(C1:C5,2)', 20], ['=RANK(30,C1:C5)', 3], ['=RANK.EQ(30,C1:C5,1)', 3],
  ['=CORREL(A1:A5,C1:C5)', 1],
  // Logic
  ['=IF(A1>0,"y","n")', 'y'], ['=IFS(A1>5,"a",A1>0,"b")', 'b'], ['=IFERROR(D4,"err")', 'err'], ['=IFNA(NA(),"na")', 'na'],
  ['=SWITCH(2,1,"one",2,"two","other")', 'two'], ['=AND(TRUE,A1=1)', true], ['=OR(FALSE,A1=2)', false],
  ['=XOR(TRUE,FALSE)', true], ['=NOT(FALSE)', true], ['=TRUE()', true], ['=FALSE()', false],
  // Information
  ['=ISBLANK(D3)', true], ['=ISNUMBER(A1)', true], ['=ISTEXT(B1)', true], ['=ISNONTEXT(A1)', true], ['=ISLOGICAL(D1)', true],
  ['=ISERROR(D4)', true], ['=ISERR(NA())', false], ['=ISNA(NA())', true], ['=ISEVEN(4)', true], ['=ISODD(3)', true],
  ['=NA()', '#N/A'], ['=N(TRUE)', 1], ['=ROW(C3)', 3], ['=COLUMN(C3)', 3], ['=ROWS(A1:B4)', 4], ['=COLUMNS(A1:B4)', 2],
  // Lookup
  ['=VLOOKUP(3,A1:C5,3,FALSE)', 30], ['=VLOOKUP(3.5,A1:C5,3)', 30], ['=HLOOKUP(2,H1:J2,2,FALSE)', 'b'],
  ['=MATCH("cherry",B1:B5,0)', 3], ['=INDEX(C1:C5,4)', 40], ['=XLOOKUP("date",B1:B5,C1:C5)', 40],
  ['=CHOOSE(3,"a","b","c")', 'c'], ['=INDIRECT("C2")', 20], ['=OFFSET(A1,2,2)', 30],
  // Text
  ['=CONCAT(B1,"-",A1)', 'apple-1'], ['=CONCATENATE("a","b")', 'ab'], ['=TEXTJOIN("/",TRUE,B1:B3)', 'apple/banana/cherry'],
  ['=LEN("hello")', 5], ['=UPPER("abc")', 'ABC'], ['=LOWER("ABC")', 'abc'], ['=PROPER("hello world")', 'Hello World'],
  ['=TRIM("  a  b ")', 'a b'], ['=LEFT("abc",2)', 'ab'], ['=RIGHT("abc",2)', 'bc'], ['=MID("abcdef",2,3)', 'bcd'],
  ['=FIND("c","abcabc",4)', 6], ['=SEARCH("B*","aBc")', 2], ['=SUBSTITUTE("aaa","a","b",2)', 'aba'],
  ['=REPLACE("abcdef",2,3,"X")', 'aXef'], ['=REPT("ab",3)', 'ababab'], ['=EXACT("a","A")', false],
  ['=VALUE("1,234.5")', 1234.5], ['=TEXT(1234.567,"#,##0.00")', '1,234.57'], ['=CHAR(65)', 'A'], ['=CODE("A")', 65],
  ['=T(B1)', 'apple'],
  // Date & time
  ['=TODAY()', (v) => Number.isInteger(v) && (v as number) > 45000], ['=NOW()', (v) => typeof v === 'number' && v > 45000],
  ['=DATE(2024,2,29)', 45351], ['=TIME(12,0,0)', 0.5], ['=YEAR(F1)', 2024], ['=MONTH(F1)', 1], ['=DAY(F1)', 31],
  ['=HOUR(0.75)', 18], ['=MINUTE(TIME(1,2,3))', 2], ['=SECOND(TIME(1,2,3))', 3], ['=WEEKDAY(DATE(2024,1,1))', 2],
  ['=EDATE(F1,1)', 45351], ['=EOMONTH(F1,1)', 45351], ['=DAYS(DATE(2024,3,1),DATE(2024,2,1))', 29],
  // Finance
  ['=PMT(0.05/12,60,10000)', -188.7123364], ['=FV(0.05,10,-100)', 1257.789254], ['=PV(0.05,10,-100)', 772.1734929],
  ['=NPV(0.1,100,100)', 100 / 1.1 + 100 / 1.21],
  // Error paths
  ['=SUM(D1:D5)', '#DIV/0!'], ['=AVERAGE(B1:B2)', '#DIV/0!'], ['=VLOOKUP("zzz",B1:C5,2,FALSE)', '#N/A'],
  ['=INDEX(A1:A5,9)', '#REF!'], ['=LEFT("a",-1)', '#VALUE!'], ['=DATE("x",1,1)', '#VALUE!'], ['=ROUND()', '#ERROR!'],
  ['=ABS(1,2)', '#ERROR!'], ['=NOSUCH(1)', '#NAME?'], ['=LARGE(C1:C5,9)', '#NUM!']
]

function evaluate (formula: string): unknown {
  const wb = new Workbook()
  const id = wb.sheets[0].id
  wb.batch(() => {
    for (const [addr, raw] of Object.entries({ ...FIXTURE, Z100: formula })) {
      const a = parseAddress(addr)!
      wb.setCell(id, a.row, a.col, raw)
    }
  })
  const v = wb.getValue(id, 99, 25)
  return typeof v === 'object' && v !== null ? String(v) : v
}

describe('built-in functions', () => {
  it.each(CASES)('%s', (formula, expected) => {
    const v = evaluate(formula)
    // Numbers compare to 6 decimal places; predicates check ranges (RAND, NOW).
    const ok = typeof expected === 'function'
      ? expected(v)
      : typeof expected === 'number'
        ? typeof v === 'number' && Math.abs(v - expected) < 5e-7
        : v === expected
    expect(ok, `${formula} gave ${String(v)}, expected ${typeof expected === 'function' ? 'a match' : String(expected)}`).toBe(true)
  })

  it('has at least one case for every function', () => {
    const tested = new Set(CASES.flatMap(([f]) => [...f.matchAll(/([A-Z][A-Z0-9.]*)\(/g)].map((m) => m[1])))
    const missing = Object.keys(FUNCTIONS).filter((name) => !tested.has(name))
    expect(missing).toEqual([])
  })
})

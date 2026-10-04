import { describe, expect, it } from 'vitest'
import { formatRange, normalizeRange, parseAddress, parseRangeAddress, quoteSheetName } from '../src/engine/address'
import { ParseError, parseFormula } from '../src/engine/parser'
import { compareScalars, Errors, errorFromCode, formatGeneral, parseNumber, toBoolean, toNumber, toText } from '../src/engine/values'
import { Workbook } from '../src/engine/workbook'

describe('addresses', () => {
  it('parses and formats cells and ranges', () => {
    expect(parseAddress('$B$3')).toEqual({ row: 2, col: 1 })
    expect(parseAddress('XFE1')).toBeUndefined()
    expect(parseAddress('A0')).toBeUndefined()
    expect(parseRangeAddress('c5:a1')).toEqual({ r1: 0, c1: 0, r2: 4, c2: 2 })
    expect(parseRangeAddress('A1:B2:C3')).toBeUndefined()
    expect(parseRangeAddress('A1:nope')).toBeUndefined()
    expect(formatRange(normalizeRange({ r1: 3, c1: 2, r2: 0, c2: 0 }))).toBe('A1:C4')
    expect(formatRange({ r1: 0, c1: 0, r2: 0, c2: 0 })).toBe('A1')
  })

  it('quotes sheet names only when needed', () => {
    expect(quoteSheetName('Data')).toBe('Data')
    expect(quoteSheetName('My Data')).toBe("'My Data'")
    expect(quoteSheetName("Bob's")).toBe("'Bob''s'")
    expect(quoteSheetName('A1')).toBe("'A1'")
  })
})

describe('parser', () => {
  it('parses array constants, unary chains and empty arguments', () => {
    expect(parseFormula('{1,2;3,4}')).toMatchObject({ t: 'array', rows: [[{ v: 1 }, { v: 2 }], [{ v: 3 }, { v: 4 }]] })
    expect(parseFormula('--1')).toMatchObject({ t: 'unary', a: { t: 'unary' } })
    expect(parseFormula('IF(A1,,2)')).toMatchObject({ t: 'call', args: [{ t: 'ref' }, { t: 'empty' }, { t: 'num' }] })
    expect(parseFormula('NOW(1,)')).toMatchObject({ args: [{ t: 'num' }, { t: 'empty' }] })
  })

  it.each([
    ['1+', /end of formula/],
    ['(1', /end of formula/],
    ['SUM(1 2)', /Expected "," or "\)"/],
    ['{1,2', /end of formula/],
    ['1 2', /Unexpected "2"/],
    [')', /Unexpected "\)"/],
    ['#BOGUS!', /Unknown error literal/]
  ])('rejects %s', (src, message) => {
    expect(() => parseFormula(src)).toThrow(message)
  })

  it('throws ParseError for structural mistakes', () => {
    expect(() => parseFormula('(1')).toThrow(ParseError)
  })

  it('evaluates array constants and percent/unary operators', () => {
    const wb = new Workbook()
    const id = wb.sheets[0].id
    wb.setCell(id, 0, 0, '=SUM({1,2;3,4})')
    wb.setCell(id, 1, 0, '=-50%+--1')
    wb.setCell(id, 2, 0, '=2^0.5^2')
    expect(wb.getValue(id, 0, 0)).toBe(10)
    expect(wb.getValue(id, 1, 0)).toBe(0.5)
    expect(wb.getValue(id, 2, 0)).toBeCloseTo(2, 10)
  })
})

describe('values', () => {
  it('coerces between types like a spreadsheet', () => {
    expect(toNumber('1,234')).toBe(1234)
    expect(toNumber(true)).toBe(1)
    expect(toNumber(null)).toBe(0)
    expect(toNumber('abc')).toBe(Errors.VALUE)
    expect(toText(1.5)).toBe('1.5')
    expect(toText(false)).toBe('FALSE')
    expect(toBoolean('true')).toBe(true)
    expect(toBoolean(0)).toBe(false)
    expect(toBoolean('maybe')).toBe(Errors.VALUE)
    expect(parseNumber('12.5%')).toBe(0.125)
    expect(parseNumber('')).toBeUndefined()
    expect(errorFromCode('#n/a')).toBe(Errors.NA)
  })

  it('orders numbers < text < booleans, with blanks equal to zero values', () => {
    expect(compareScalars(5, 'a')).toBeLessThan(0)
    expect(compareScalars('a', true)).toBeLessThan(0)
    expect(compareScalars('Apple', 'apple')).toBe(0)
    expect(compareScalars(null, 0)).toBe(0)
    expect(compareScalars(null, '')).toBe(0)
    expect(compareScalars(false, null)).toBe(0)
    expect(compareScalars(true, false)).toBeGreaterThan(0)
  })

  it('renders General numbers with Excel-like precision', () => {
    expect(formatGeneral(123456789012)).toBe('123456789012')
    expect(formatGeneral(1.5e20)).toBe('1.5E+20')
    expect(formatGeneral(123456789012.5)).toBe('1.23457E+11')
    expect(formatGeneral(1 / 3)).toBe('0.33333333333')
    expect(formatGeneral(-0.000001)).toBe('-1E-06')
    expect(formatGeneral(Infinity)).toBe('#NUM!')
  })
})

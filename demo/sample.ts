import type { SheetJSON, Style, WorkbookJSON } from '../src/engine/index'

const header: Style = { bold: true }
/** Input cells: a fixed light fill needs a fixed dark text colour to stay readable in dark mode. */
const input: Style = { fill: '#fff4c2', color: '#3a2f00' }
const money = '$#,##0.00'

/** A small workbook that shows off formulas, formats and cross-sheet references. */
export function sampleWorkbook (): WorkbookJSON {
  const regions = ['North', 'South', 'East', 'West']
  const products: [string, number][] = [['Widget', 12.5], ['Gadget', 24], ['Gizmo', 7.25], ['Doohickey', 99]]
  const sales: SheetJSON['cells'] = [
    [0, 0, 'Date', header], [0, 1, 'Region', header], [0, 2, 'Product', header], [0, 3, 'Units', header],
    [0, 4, 'Unit price', header], [0, 5, 'Revenue', header], [0, 6, 'Big order?', header]
  ]
  // Deterministic pseudo-random data so the sample looks the same every time.
  let seed = 7
  const rand = (): number => ((seed = (seed * 16807) % 2147483647) / 2147483647)
  const rows = 40
  for (let i = 1; i <= rows; i++) {
    const n = i + 1
    const [product, price] = products[Math.floor(rand() * products.length)]
    const month = 1 + Math.floor((i - 1) / 10)
    const day = 1 + Math.floor(rand() * 27)
    sales.push(
      [i, 0, `2026-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`, { fmt: 'yyyy-mm-dd' }],
      [i, 1, regions[Math.floor(rand() * regions.length)]],
      [i, 2, product],
      [i, 3, String(5 + Math.floor(rand() * 120))],
      [i, 4, String(price), { fmt: money }],
      [i, 5, `=D${n}*E${n}`, { fmt: money }],
      [i, 6, `=IF(F${n}>2000,"Yes","")`, { align: 'center' }]
    )
  }
  const total = rows + 2
  sales.push(
    [total, 0, 'Total', { bold: true }],
    [total, 3, `=SUM(D2:D${rows + 1})`, { bold: true }],
    [total, 5, `=SUM(F2:F${rows + 1})`, { bold: true, fmt: money }],
    [total, 6, `=COUNTIF(G2:G${rows + 1},"Yes")&" big"`, { bold: true, align: 'center' }]
  )

  const summary: SheetJSON['cells'] = [
    [0, 0, 'Revenue by region', { bold: true }],
    [1, 0, 'Region', header], [1, 1, 'Revenue', header], [1, 2, 'Share', header], [1, 3, 'Orders', header], [1, 4, 'Avg order', header]
  ]
  regions.forEach((region, i) => {
    const r = i + 2
    const n = r + 1
    summary.push(
      [r, 0, region],
      [r, 1, `=SUMIF(Sales!B:B,A${n},Sales!F:F)`, { fmt: money }],
      [r, 2, `=B${n}/B$7`, { fmt: '0.0%' }],
      [r, 3, `=COUNTIF(Sales!B:B,A${n})`],
      [r, 4, `=IFERROR(B${n}/D${n},0)`, { fmt: money }]
    )
  })
  summary.push(
    [6, 0, 'All regions', { bold: true }],
    [6, 1, '=SUM(B3:B6)', { bold: true, fmt: money }],
    [6, 2, '=SUM(C3:C6)', { bold: true, fmt: '0.0%' }],
    [6, 3, '=SUM(D3:D6)', { bold: true }],
    [6, 4, '=B7/D7', { bold: true, fmt: money }],

    [8, 0, 'Lookups', { bold: true }],
    [9, 0, 'Product'], [9, 1, 'Gadget', input],
    [10, 0, 'Units sold'], [10, 1, '=SUMIFS(Sales!D:D,Sales!C:C,B10)'],
    [11, 0, 'Best region'], [11, 1, '=INDEX(A3:A6,MATCH(MAX(B3:B6),B3:B6,0))'],
    [12, 0, 'Largest order'], [12, 1, '=MAX(Sales!F2:F41)', { fmt: money }],
    [13, 0, 'Median units'], [13, 1, '=MEDIAN(Sales!D2:D41)'],
    [14, 0, 'Price lookup'], [14, 1, '=XLOOKUP(B10,Sales!C:C,Sales!E:E,"n/a")', { fmt: money }],

    [16, 0, 'Loan calculator', { bold: true }],
    [17, 0, 'Principal'], [17, 1, '25000', { ...input, fmt: money }],
    [18, 0, 'Annual rate'], [18, 1, '6.5%', { ...input, fmt: '0.00%' }],
    [19, 0, 'Years'], [19, 1, '5', input],
    [20, 0, 'Monthly payment'], [20, 1, '=-PMT(B19/12,B20*12,B18)', { fmt: money, bold: true }],
    [21, 0, 'Total interest'], [21, 1, '=B21*B20*12-B18', { fmt: money }],

    [23, 0, 'Tip: edit the yellow cells, or any number on the Sales sheet, and watch everything update.', { italic: true }]
  )

  return {
    version: 1,
    sheets: [
      { name: 'Summary', cells: summary, colWidths: [[0, 150], [1, 120], [2, 80], [3, 80], [4, 100]] },
      { name: 'Sales', cells: sales, colWidths: [[0, 100], [2, 100], [5, 110], [6, 90]] }
    ]
  }
}

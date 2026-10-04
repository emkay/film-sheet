import { toast } from '@mk/film'
import '@mk/film/css/themes/default/index.css'
import '../src/index'
import type { WorkbookJSON, Style } from '../src/index'
import type { FilmSheet } from '../src/ui/film-sheet'
import './demo.css'
import { sampleWorkbook } from './sample'

/**
 * The FilmSheet demo: a page that hosts one <film-sheet> and adds the
 * page-level things a library component shouldn't own — the document title,
 * the Film theme picker, sample data and a performance test.
 */

type ValueEl = HTMLElement & { value: string }
type CheckEl = HTMLElement & { checked: boolean }
type DialogEl = HTMLElement & { show (): void; close (): void }

const PREFS_KEY = 'film-sheet:prefs'
// oxlint-disable-next-line typescript/no-unnecessary-type-parameters
const $ = <T extends HTMLElement = HTMLElement>(sel: string): T => document.querySelector<T>(sel)!

await customElements.whenDefined('film-sheet')
const sheet = $<FilmSheet>('film-sheet')
const title = $<HTMLInputElement>('.doc-title')
await sheet.updateComplete

// First visit (nothing autosaved): start with the sample workbook.
if (!localStorage.getItem(sheet.storageKey)) {
  sheet.load(sampleWorkbook())
  sheet.filename = 'Sample: Quarterly sales'
}

// Document title ↔ the sheet's download name, and the page title.
let activeSheet = sheet.workbook?.sheets[0]?.name ?? ''
const updatePageTitle = (): void => {
  document.title = `${sheet.filename || 'Untitled'} – ${activeSheet} – FilmSheet`
}
title.value = sheet.filename
title.addEventListener('change', () => {
  sheet.filename = title.value
  updatePageTitle()
})
title.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') sheet.focus()
})
sheet.addEventListener('film-sheet-select', (e) => {
  activeSheet = e.detail.sheet
  if (title.value !== sheet.filename) title.value = sheet.filename
  updatePageTitle()
})
updatePageTitle()

// Film theme and colour scheme for the whole page; <film-sheet> follows along.
const themeSelect = $<ValueEl>('.theme-select')
const dark = $<CheckEl>('.dark-switch')
function applyTheme (theme: string, isDark: boolean): void {
  document.documentElement.dataset.filmTheme = theme
  document.documentElement.dataset.theme = isDark ? 'dark' : 'light'
  try {
    localStorage.setItem(PREFS_KEY, JSON.stringify({ theme, dark: isDark }))
  } catch { /* private mode */ }
}
let prefs: { theme?: string; dark?: boolean } = {}
try {
  prefs = JSON.parse(localStorage.getItem(PREFS_KEY) ?? '{}')
} catch { /* ignore */ }
themeSelect.value = prefs.theme ?? 'paper'
dark.checked = prefs.dark ?? matchMedia('(prefers-color-scheme: dark)').matches
applyTheme(themeSelect.value, dark.checked)
themeSelect.addEventListener('change', () => applyTheme(themeSelect.value, dark.checked))
dark.addEventListener('change', () => applyTheme(themeSelect.value, dark.checked))

// Demo actions
$('.load-sample').addEventListener('click', () => {
  sheet.load(sampleWorkbook())
  sheet.filename = title.value = 'Sample: Quarterly sales'
  updatePageTitle()
  sheet.focus()
})
$('.open-stress').addEventListener('click', () => {
  $('.stress-result').textContent = ''
  $<DialogEl>('#stress-dialog').show()
})
$('.stress-close').addEventListener('click', () => $<DialogEl>('#stress-dialog').close())
$('.stress-run').addEventListener('click', () => void stressTest())

/**
 * Build a large sheet (values, two formulas per row and a running-total chain
 * as deep as the sheet is long) and time how long it takes to load.
 */
async function stressTest (): Promise<void> {
  const rows = Number($<ValueEl>('.stress-rows').value)
  const progress = $<HTMLElement & { indeterminate: boolean }>('.stress-progress')
  const result = $('.stress-result')
  progress.hidden = false
  progress.indeterminate = true
  result.textContent = 'Building…'
  await new Promise((r) => setTimeout(r, 50))

  const t0 = performance.now()
  const bold: Style = { bold: true }
  const cells: [number, number, string, Style?][] = [
    [0, 0, 'Value', bold], [0, 1, 'Double', bold], [0, 2, 'Squared', bold], [0, 3, 'Running total', bold],
    [0, 5, 'Summary', bold], [1, 5, 'Rows'], [1, 6, '=COUNT(A:A)'], [2, 5, 'Sum of values'], [2, 6, '=SUM(A:A)', { fmt: '#,##0.00' }],
    [3, 5, 'Average'], [3, 6, '=AVERAGE(A:A)', { fmt: '0.000' }], [4, 5, 'Values > 500'], [4, 6, '=COUNTIF(A:A,">500")'],
    [5, 5, 'Grand total'], [5, 6, `=D${rows + 1}`, { fmt: '#,##0.00' }], [6, 5, 'Max squared'], [6, 6, '=MAX(C:C)', { fmt: '#,##0' }]
  ]
  for (let r = 1; r <= rows; r++) {
    const n = r + 1
    cells.push([r, 0, String(Math.round(Math.random() * 100000) / 100)])
    cells.push([r, 1, `=A${n}*2`])
    cells.push([r, 2, `=A${n}^2`])
    cells.push([r, 3, r === 1 ? `=A${n}` : `=D${r}+A${n}`])
  }
  const json: WorkbookJSON = sheet.toJSON() ?? { version: 1, sheets: [] }
  json.sheets.push({ name: `Perf ${rows.toLocaleString('en-US').replace(/,/g, '_')}`, cells, colWidths: [[2, 120], [3, 140], [5, 110], [6, 140]] })
  const tBuild = performance.now()
  sheet.load(json)
  const tLoad = performance.now()
  progress.hidden = true
  const formulas = rows * 3 + 6
  result.textContent = `${(rows * 4).toLocaleString()} cells (${formulas.toLocaleString()} formulas, including a ${rows.toLocaleString()}-deep running-total chain) compiled and calculated in ${(tLoad - tBuild).toFixed(0)} ms (+${(tBuild - t0).toFixed(0)} ms to generate data). Open the new sheet and type a number in column A: only the cells that depend on it are recalculated, and the time shows in the status bar.`
  toast(`Built ${(rows * 4).toLocaleString()} cells in ${(tLoad - tBuild).toFixed(0)} ms`, { variant: 'success', duration: 6000 })
}

sheet.focus()

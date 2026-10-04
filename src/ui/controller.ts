import { toast } from '@mk/film'
import { colToName, formatAddress, formatRange, MAX_COLS, MAX_ROWS, normalizeRange, parseRangeAddress, type RangeAddress } from '../engine/address.js'
import { FORMATS, valueToRaw } from '../engine/format.js'
import { FUNCTION_NAMES, FUNCTIONS } from '../engine/functions.js'
import { formatRefToken, normalizeFormula, shiftFormula } from '../engine/refs.js'
import { tokenize } from '../engine/tokenizer.js'
import { formatGeneral, type Scalar } from '../engine/values.js'
import { Workbook, type Cell, type CellPatch, type Sheet, type Style, type WorkbookJSON } from '../engine/workbook.js'
import { CellEditor } from './editor.js'
import { fillPatches } from './fill.js'
import { GridView, type RangeHighlight, type Selection } from './grid.js'
import { History } from './history.js'
import { csvSafe, download, escapeHtml, parseDelimited, toDelimited } from './io.js'

const AUTOSAVE_CELL_LIMIT = 250_000
/** Most cells one formatting or Ctrl+Enter action will write individually. */
const STYLE_CELL_LIMIT = 500_000
/** Commands that don't change the workbook, so they stay available when read-only. */
const READONLY_COMMANDS = new Set(['file.save', 'file.exportCsv', 'edit.copy', 'edit.selectAll', 'edit.find', 'view.formulas', 'view.goto', 'help.shortcuts', 'help.functions'])

const REF_COLORS = ['#1a73e8', '#d93025', '#7b1fa2', '#188038', '#e37400', '#0097a7']
const SWATCHES: [string, string][] = [
  ['#000000', 'Black'], ['#434343', 'Dark grey'], ['#666666', 'Grey'], ['#999999', 'Medium grey'], ['#cccccc', 'Light grey'], ['#ffffff', 'White'],
  ['#c0392b', 'Red'], ['#e67e22', 'Orange'], ['#f1c40f', 'Yellow'], ['#27ae60', 'Green'], ['#2980b9', 'Blue'], ['#8e44ad', 'Purple'],
  ['#f5b7b1', 'Light red'], ['#fad7a0', 'Light orange'], ['#fcf3cf', 'Light yellow'], ['#abebc6', 'Light green'], ['#aed6f1', 'Light blue'], ['#d7bde2', 'Light purple']
]

interface Clip {
  sheetId: number
  range: RangeAddress
  cells: ({ raw: string; style: Style | null; value: Scalar } | null)[][]
  text: string
  cut: boolean
}

type Dir = 'up' | 'down' | 'left' | 'right'

// Film's form controls expose `value`/`checked` as properties.
type ValueEl = HTMLElement & { value: string }
type CheckEl = HTMLElement & { checked: boolean }
type DialogEl = HTMLElement & { open: boolean; show (): void; close (): void; label: string }

/** What the controller needs from the <film-sheet> element that owns it. */
export interface SheetHost extends HTMLElement {
  readonly: boolean
  storageKey: string
  filename: string
}

const isMac = /Mac|iPhone|iPad/.test(navigator.platform)

/**
 * Everything behind <film-sheet>: commands, keyboard, clipboard, undo,
 * persistence and the wiring between the engine, grid, editor and chrome.
 * All DOM access goes through `root` (the element's shadow root), so several
 * sheets can live on one page without seeing each other.
 */
export class SheetController {
  wb: Workbook
  sheet: Sheet
  readonly grid: GridView
  readonly editor: CellEditor
  readonly history = new History()
  private clip: Clip | null = null
  private showFormulas = false
  private editOriginal = ''
  /** Where the edit is being typed: the cell overlay or the formula bar. */
  private editSource: 'cell' | 'bar' = 'cell'
  private pointSel: Selection | null = null
  private statsFrame = 0
  private saveTimer = 0
  private promptResolve: ((v: string | null) => void) | null = null

  private readonly nameBox: ValueEl
  private readonly formulaBar: ValueEl
  private readonly tabs: ValueEl & { active: string }
  private readonly statusMode: HTMLElement
  private readonly statusStats: HTMLElement
  private readonly statusPerf: HTMLElement
  private readonly formatSelect: ValueEl
  private readonly saveState: HTMLElement
  private readonly announcer: HTMLElement
  /** Undo for everything this controller attached outside its own subtree. */
  private readonly cleanups: (() => void)[] = []
  private disposed = false
  private announceTimer = 0
  private selectionAnnounceTimer = 0
  private lastToolbarKey = ''

  constructor (
    private readonly host: SheetHost,
    private readonly root: ShadowRoot,
    initial?: WorkbookJSON
  ) {
    this.nameBox = this.$<ValueEl>('.name-box')
    this.formulaBar = this.$<ValueEl>('.formula-input')
    this.tabs = this.$<ValueEl & { active: string }>('.sheet-tabs')
    this.statusMode = this.$('.status-mode')
    this.statusStats = this.$('.status-stats')
    this.statusPerf = this.$('.status-perf')
    this.formatSelect = this.$<ValueEl>('.format-select')
    this.saveState = this.$('.save-state')
    this.announcer = this.$('.announcer')

    this.wb = new Workbook()
    let active: number | null = null
    if (initial) this.wb.load(initial)
    else active = this.restore()
    this.sheet = this.wb.sheets[active ?? 0] ?? this.wb.sheets[0]

    this.grid = new GridView(this.wb, this.sheet)
    this.$('.grid-host').append(this.grid.el)
    // The editor lives beside the grid, not inside its role="grid" element.
    this.editor = new CellEditor(this.$('.grid-host'))
    this.editor.onInput = (text) => this.onEditorInput(text)
    this.editor.onAnnounce = (text) => this.announce(text)

    this.grid.callbacks = {
      selectionChange: () => this.onSelectionChange(),
      cellPointerDown: (row, col) => this.onCellPointerDown(row, col),
      pick: (range) => {
        // Arrow-key pointing continues from where the mouse pointed.
        this.pointSel = { ar: range.r1, ac: range.c1, fr: range.r2, fc: range.c2 }
        this.pointAt(range)
      },
      edit: (row, col) => {
        // A double-click while pointing a formula has already inserted the reference.
        if (this.editor.active) return
        this.grid.setSelection({ ar: row, ac: col, fr: row, fc: col }, false)
        this.startEdit(undefined, 'edit')
      },
      fill: (src, dest) => this.fill(src, dest),
      resized: (axis, index, before, after) => {
        if (before === after) return
        this.history.push({ kind: 'size', sheetId: this.sheet.id, axis, index, before, after })
        this.scheduleSave()
      },
      contextMenu: (x, y) => this.showMenu('#context-menu', x, y),
      otherPointerDown: () => {
        if (this.editor.active) this.commitEdit('none')
      },
      viewportChange: () => {
        if (!this.editor.active) return
        const rect = this.grid.cellRect(this.editor.row, this.editor.col)
        const visible = rect.x + rect.w > this.grid.headerWidth && rect.y + rect.h > 24 && rect.x < this.grid.viewportWidth
        this.editor.reposition(rect, visible)
      }
    }

    this.cleanups.push(this.wb.onChange(() => {
      this.grid.invalidate()
      this.grid.syncA11y()
      this.scheduleStats()
      this.scheduleSave()
      this.dispatch('film-sheet-change', null)
    }))

    this.localizeShortcuts()
    this.bindChrome()
    this.bindKeyboard()
    this.bindClipboard()
    this.renderTabs()
    this.onSelectionChange()
    this.watchTheme()
  }

  /** Element lookup inside this sheet's shadow root. */
  // T lets callers name the element type they expect, like querySelector<T>.
  // oxlint-disable-next-line typescript/no-unnecessary-type-parameters
  private $<T extends Element = HTMLElement> (sel: string, within: ParentNode = this.root): T {
    const el = within.querySelector<T>(sel)
    if (!el) throw new Error(`Missing element ${sel}`)
    return el
  }

  /** Focus that has left the page entirely or fell back to <body>. */
  private focusLost (): boolean {
    const a = document.activeElement
    return this.root.activeElement === null && (a === null || a === document.body || a === this.host)
  }

  /** Tear down everything that would outlive the element. */
  dispose (): void {
    if (this.disposed) return
    this.disposed = true
    clearTimeout(this.saveTimer)
    if (this.host.storageKey) this.save()
    for (const fn of this.cleanups) fn()
    this.grid.dispose()
    this.editor.dispose()
  }

  /** Re-resolve canvas colours when the page's Film theme or colour scheme changes. */
  private watchTheme (): void {
    const refresh = (): number => requestAnimationFrame(() => this.grid.refreshTheme())
    const observer = new MutationObserver(refresh)
    observer.observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme', 'data-film-theme', 'class', 'style'] })
    this.cleanups.push(() => observer.disconnect())
  }

  /** Reflect the host's `readonly` state in the chrome. */
  applyReadonly (): void {
    const ro = this.host.readonly
    if (ro && this.editor.active) this.cancelEdit()
    for (const item of this.root.querySelectorAll<HTMLElement & { disabled: boolean }>('film-menu-item[value]')) {
      item.disabled = ro && !READONLY_COMMANDS.has(item.getAttribute('value')!)
    }
    for (const sub of this.root.querySelectorAll<HTMLElement & { disabled: boolean }>('film-menu-item:not([value])')) {
      sub.disabled = ro && !sub.querySelector('film-menu-item[value]:not([disabled])')
    }
    ;(this.formulaBar as ValueEl & { readonly: boolean }).readonly = ro
    this.grid.readonly = ro
    this.grid.invalidate()
  }

  /** Replace the workbook (public API: <film-sheet>.load()). */
  load (json: WorkbookJSON): void {
    this.loadWorkbook(json)
  }

  /** Notify the host page. */
  private dispatch (name: 'film-sheet-change' | 'film-sheet-select', detail: unknown): void {
    this.host.dispatchEvent(new CustomEvent(name, { detail, bubbles: true, composed: true }))
  }

  // ===========================================================================
  // Sheets

  private renderTabs (): void {
    const tabs = this.wb.sheets.map((s) => {
      const t = document.createElement('film-tab')
      t.setAttribute('slot', 'nav')
      t.setAttribute('panel', String(s.id))
      t.textContent = s.name
      return t
    })
    this.tabs.replaceChildren(...tabs)
    this.tabs.active = String(this.sheet.id)
    // The tablist role lives in film-tabs' shadow DOM; name it there.
    void (this.tabs as unknown as { updateComplete: Promise<unknown> }).updateComplete.then(() => {
      this.tabs.shadowRoot?.querySelector('[role="tablist"]')?.setAttribute('aria-label', 'Sheets')
    })
  }

  private switchSheet (sheet: Sheet): void {
    if (this.editor.active) this.commitEdit('none')
    this.sheet = sheet
    this.grid.setSheet(sheet)
    this.grid.setSelection({ ar: 0, ac: 0, fr: 0, fc: 0 })
    this.tabs.active = String(sheet.id)
    this.announce(`${sheet.name}, sheet ${this.wb.sheets.indexOf(sheet) + 1} of ${this.wb.sheets.length}`)
    this.scheduleSave()
  }

  /**
   * Speak a short status message through the polite live region. Cleared
   * first so repeating the same message is announced again.
   */
  announce (text: string): void {
    clearTimeout(this.announceTimer)
    this.announcer.textContent = ''
    this.announceTimer = window.setTimeout(() => {
      this.announcer.textContent = text
    }, 100)
  }

  /**
   * Run a whole-workbook change with undo. Only the "before" state is stored
   * up front; the "after" state is captured when the entry is first undone,
   * so each entry costs one copy of the workbook rather than two. `fn` can
   * return false to say nothing changed, and no entry is recorded.
   */
  private snapshotAction (fn: () => boolean | void): boolean {
    if (this.host.readonly) return false
    this.cancelCut()
    const before = this.wb.toJSON()
    const sheetBefore = this.sheet.id
    if (fn() === false) return false
    this.history.push({ kind: 'snapshot', before, sheetBefore, sheetAfter: this.sheet.id })
    return true
  }

  /** Format → Column width… / Row height…: the keyboard route to resizing. */
  private async promptSize (axis: 'rows' | 'cols'): Promise<void> {
    const r = this.grid.range
    const ax = axis === 'cols' ? this.grid.cols : this.grid.rows
    const first = axis === 'cols' ? r.c1 : r.r1
    const last = axis === 'cols' ? Math.min(r.c2, r.c1 + 1000) : Math.min(r.r2, r.r1 + 10000)
    const v = await this.prompt(axis === 'cols' ? 'Column width' : 'Row height', 'Size in pixels', String(ax.size(first)))
    if (v === null) return
    const n = Math.round(Number(v))
    if (!Number.isFinite(n) || n < 8 || n > 2000) {
      toast('Enter a size between 8 and 2000 pixels.', { variant: 'warning' })
      return
    }
    for (let i = first; i <= last; i++) this.setSize(axis, i, n)
  }

  private setSize (axis: 'rows' | 'cols', index: number, size: number): void {
    if (this.host.readonly) return
    const ax = axis === 'cols' ? this.grid.cols : this.grid.rows
    const before = ax.size(index)
    if (before === size) return
    ax.setSize(index, size)
    this.grid.sizesChanged()
    this.history.push({ kind: 'size', sheetId: this.sheet.id, axis, index, before, after: size })
    this.scheduleSave()
  }

  private async renameSheet (): Promise<void> {
    const name = await this.prompt('Rename sheet', 'Sheet name', this.sheet.name)
    if (name === null || name === this.sheet.name) return
    if (!this.snapshotAction(() => this.wb.renameSheet(this.sheet.id, name))) {
      toast(`Can't rename to "${name}": the name is taken or contains [ ] : * ? / \\`, { variant: 'danger', duration: 0 })
      return
    }
    this.renderTabs()
  }

  /** A pending cut is cancelled by any other edit, as in Excel. */
  private cancelCut (): void {
    if (!this.clip?.cut) return
    this.clip = null
    this.grid.setCopyRange(null)
  }

  // ===========================================================================
  // Selection & formula bar

  private get active (): { row: number; col: number } {
    return { row: this.grid.sel.ar, col: this.grid.sel.ac }
  }

  private activeCell (): Cell | undefined {
    return this.sheet.get(this.grid.sel.ar, this.grid.sel.ac)
  }

  private onSelectionChange (): void {
    const r = this.grid.range
    this.dispatch('film-sheet-select', {
      sheet: this.sheet.name,
      address: formatAddress(this.grid.sel.ar, this.grid.sel.ac),
      range: r.r1 === r.r2 && r.c1 === r.c2 ? formatAddress(r.r1, r.c1) : formatRange(r)
    })
    const whole = (r.r1 === 0 && r.r2 === MAX_ROWS - 1) || (r.c1 === 0 && r.c2 === MAX_COLS - 1)
    this.nameBox.value = whole ? formatRange(r).replace(/^A1:XFD1048576$/, 'All') : r.r1 === r.r2 && r.c1 === r.c2 ? formatAddress(r.r1, r.c1) : formatRange(r)
    if (!this.editor.active) this.formulaBar.value = this.activeCell()?.raw ?? ''
    // Dragging a selection fires this per cell; the toolbar only follows the active cell.
    const style = this.activeCell()?.style
    const activeKey = `${this.sheet.id}:${this.grid.sel.ar}:${this.grid.sel.ac}:${style ? JSON.stringify(style) : ''}`
    if (activeKey !== this.lastToolbarKey) {
      this.lastToolbarKey = activeKey
      this.updateToolbar()
    }
    this.scheduleStats()
    // Single-cell moves are announced by the focus proxy; ranges need words.
    clearTimeout(this.selectionAnnounceTimer)
    if ((r.r1 !== r.r2 || r.c1 !== r.c2) && this.grid.hasFocus()) {
      this.selectionAnnounceTimer = window.setTimeout(() => {
        const rows = r.r2 - r.r1 + 1
        const cols = r.c2 - r.c1 + 1
        const what = r.r1 === 0 && r.r2 === MAX_ROWS - 1 && r.c1 === 0 && r.c2 === MAX_COLS - 1
          ? 'All cells'
          : r.r1 === 0 && r.r2 === MAX_ROWS - 1 ? `Column${cols > 1 ? 's' : ''} ${colToName(r.c1)}${cols > 1 ? ' to ' + colToName(r.c2) : ''}`
            : r.c1 === 0 && r.c2 === MAX_COLS - 1 ? `Row${rows > 1 ? 's' : ''} ${r.r1 + 1}${rows > 1 ? ' to ' + (r.r2 + 1) : ''}`
              : `${formatAddress(r.r1, r.c1)} to ${formatAddress(r.r2, r.c2)}, ${(rows * cols).toLocaleString()} cells`
        this.announce(`${what} selected. ${this.statusStats.textContent ?? ''}`)
      }, 250)
    }
  }

  private updateToolbar (): void {
    const style = this.activeCell()?.style ?? {}
    // film-icon-button renders its real <button> in shadow DOM, so the
    // pressed state has to be set there for assistive tech to see it.
    const setPressed = (el: HTMLElement, on: boolean): void => {
      el.toggleAttribute('data-on', on)
      el.shadowRoot?.querySelector('button')?.setAttribute('aria-pressed', String(on))
    }
    for (const el of this.root.querySelectorAll<HTMLElement>('[data-toggle]')) setPressed(el, !!style[el.dataset.toggle as 'bold'])
    for (const el of this.root.querySelectorAll<HTMLElement>('[data-align]')) setPressed(el, style.align === el.dataset.align)
    const fmt = style.fmt
    const key = (Object.keys(FORMATS) as (keyof typeof FORMATS)[]).find((k) => FORMATS[k] === fmt)
    this.formatSelect.value = !fmt || fmt === 'General' ? 'general' : key ?? 'custom'
  }

  private scheduleStats (): void {
    if (this.statsFrame) return
    this.statsFrame = requestAnimationFrame(() => {
      this.statsFrame = 0
      this.updateStats()
    })
  }

  /** Sum / average / count of the selection, like Excel's status bar. */
  private updateStats (): void {
    const r = this.grid.range
    if (r.r1 === r.r2 && r.c1 === r.c2) {
      this.statusStats.textContent = ''
      return
    }
    const s = this.sheet
    let count = 0
    let nums = 0
    let sum = 0
    let min = Infinity
    let max = -Infinity
    const r2 = Math.min(r.r2, s.maxRow)
    const visit = (cell: Cell): void => {
      const v = cell.value
      if (v === null) return
      count++
      if (typeof v === 'number') {
        nums++
        sum += v
        if (v < min) min = v
        if (v > max) max = v
      }
    }
    for (const [c, col] of s.cols) {
      if (c < r.c1 || c > r.c2) continue
      if (r2 - r.r1 + 1 > col.size) {
        for (const cell of col.values()) if (cell.row >= r.r1 && cell.row <= r2) visit(cell)
      } else {
        for (let row = r.r1; row <= r2; row++) {
          const cell = col.get(row)
          if (cell) visit(cell)
        }
      }
    }
    const fmt = (n: number): string => formatGeneral(Math.round(n * 1e10) / 1e10)
    const parts: string[] = []
    if (nums) parts.push(`Sum ${fmt(sum)}`, `Average ${fmt(sum / nums)}`, `Min ${fmt(min)}`, `Max ${fmt(max)}`)
    if (count) parts.push(`Count ${count.toLocaleString()}`)
    this.statusStats.textContent = parts.join('   ')
  }

  private setMode (mode: string): void {
    this.statusMode.textContent = mode
  }

  private showPerf (): void {
    const { evaluated, ms } = this.wb.lastRecalc
    if (evaluated) this.statusPerf.textContent = `Recalculated ${evaluated.toLocaleString()} cell${evaluated === 1 ? '' : 's'} in ${ms < 1 ? ms.toFixed(2) : ms.toFixed(0)} ms`
  }

  // ===========================================================================
  // Editing

  private startEdit (initial?: string, mode: 'enter' | 'edit' = 'enter', source: 'cell' | 'bar' = 'cell'): void {
    if (this.host.readonly) {
      this.announce('This sheet is read-only')
      return
    }
    const { row, col } = this.active
    const cell = this.activeCell()
    this.editOriginal = cell?.raw ?? ''
    this.editSource = source
    this.pointSel = null
    this.grid.ensureVisible(row, col)
    const rect = this.grid.cellRect(row, col)
    const style = cell?.style
    const font = `${style?.italic ? 'italic ' : ''}${style?.bold ? '600 ' : ''}13px ${getComputedStyle(this.host).getPropertyValue('--film-font-sans') || 'system-ui'}`
    this.grid.editing = { row, col }
    this.editor.el.setAttribute('aria-label', `Edit ${formatAddress(row, col)}`)
    this.editor.open(row, col, initial ?? this.editOriginal, mode, rect, this.grid.viewportWidth, font)
    if (source === 'bar') this.focusBar()
    this.setMode(mode === 'enter' ? 'Enter' : 'Edit')
    this.grid.invalidate()
  }

  private onEditorInput (text: string): void {
    if (this.editSource === 'cell' && this.formulaBar.value !== text) this.formulaBar.value = text
    this.grid.highlights = this.formulaHighlights(text)
    this.grid.invalidate()
  }

  /** Colour each reference in the formula being edited, as Excel does. */
  private formulaHighlights (text: string): RangeHighlight[] {
    if (!text.startsWith('=')) return []
    let toks
    try {
      toks = tokenize(text.slice(1))
    } catch {
      return []
    }
    const out: RangeHighlight[] = []
    const seen = new Map<string, string>()
    for (const t of toks) {
      if (t.type !== 'ref' || !t.ref) continue
      if (t.ref.sheet !== undefined && t.ref.sheet.toLowerCase() !== this.sheet.name.toLowerCase()) continue
      const key = t.text.replace(/\$/g, '').toUpperCase()
      let color = seen.get(key)
      if (!color) {
        color = REF_COLORS[seen.size % REF_COLORS.length]
        seen.set(key, color)
      }
      out.push({ ...normalizeRange(t.ref), color })
    }
    return out
  }

  private commitEdit (move: Dir | 'none', fillSelection = false): void {
    if (!this.editor.active) return
    let text = this.editor.value
    const { row, col } = this.editor
    this.closeEditor()
    if (text.startsWith('=')) text = normalizeFormula(text)
    const range = this.grid.range
    const area = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1)
    if (fillSelection && area > STYLE_CELL_LIMIT) {
      toast(`Ctrl+Enter fills at most ${STYLE_CELL_LIMIT.toLocaleString()} cells; only ${formatAddress(row, col)} was changed.`, { variant: 'warning' })
      if (text !== this.editOriginal) this.applyPatches([{ r: row, c: col, raw: text }], true)
    } else if (fillSelection && area > 1) {
      const patches: CellPatch[] = []
      for (let r = range.r1; r <= range.r2; r++) {
        for (let c = range.c1; c <= range.c2; c++) {
          patches.push({ r, c, raw: text.startsWith('=') ? shiftFormula(text, r - row, c - col) : text })
        }
      }
      this.applyPatches(patches, true)
    } else if (text !== this.editOriginal) {
      this.applyPatches([{ r: row, c: col, raw: text }], true)
    }
    const cell = this.sheet.get(row, col)
    if (cell?.error) toast(`${formatAddress(row, col)}: ${cell.error}`, { variant: 'danger', duration: 0 })
    if (move !== 'none') this.moveActive(move)
    else this.onSelectionChange()
  }

  private cancelEdit (): void {
    if (!this.editor.active) return
    this.closeEditor()
    this.onSelectionChange()
  }

  private closeEditor (): void {
    this.editor.close()
    this.grid.editing = null
    this.grid.highlights = []
    this.pointSel = null
    this.setMode('Ready')
    this.grid.invalidate()
    this.grid.focus()
  }

  /** Clicking a cell while typing a formula inserts its reference. */
  private onCellPointerDown (_row: number, _col: number): boolean {
    if (!this.editor.active) return false
    if (this.editSource === 'bar') this.syncEditorFromBar()
    if (this.editor.canPoint()) return true
    this.commitEdit('none')
    return false
  }

  private pointAt (range: RangeAddress): void {
    const ref = range.r1 === range.r2 && range.c1 === range.c2 ? formatAddress(range.r1, range.c1) : formatRange(range)
    if (this.editSource === 'bar') this.syncEditorFromBar()
    this.editor.point(ref)
    this.setMode('Point')
    this.announce(`Pointing ${ref.replace(':', ' to ')}`)
    if (this.editSource === 'bar') {
      this.formulaBar.value = this.editor.value
      this.focusBar(this.editor.el.selectionStart)
    } else this.editor.el.focus({ preventScroll: true })
  }

  /** Arrow keys while typing a formula move a reference "pointer". */
  private pointMove (dir: Dir, extend: boolean): void {
    if (!this.pointSel || !this.editor.pointing) {
      const { row, col } = this.editor
      this.pointSel = { ar: row, ac: col, fr: row, fc: col }
    }
    const p = this.pointSel
    const [dr, dc] = dirDelta(dir)
    if (extend) {
      p.fr = clamp(p.fr + dr, 0, MAX_ROWS - 1)
      p.fc = clamp(p.fc + dc, 0, MAX_COLS - 1)
    } else {
      p.ar = p.fr = clamp(p.fr + dr, 0, MAX_ROWS - 1)
      p.ac = p.fc = clamp(p.fc + dc, 0, MAX_COLS - 1)
    }
    this.grid.ensureVisible(p.fr, p.fc)
    this.pointAt(normalizeRange({ r1: p.ar, c1: p.ac, r2: p.fr, c2: p.fc }))
  }

  /** F4 cycles the reference under the caret through $A$1 → A$1 → $A1 → A1. */
  private toggleAbsolute (): void {
    const text = this.editor.value
    if (!text.startsWith('=')) return
    const caret = this.editor.el.selectionStart - 1
    let toks
    try {
      toks = tokenize(text.slice(1))
    } catch {
      return
    }
    const t = toks.find((x) => x.type === 'ref' && caret >= x.start && caret <= x.end)
    if (!t || !t.ref) return
    const [c1, r1] = t.ref.abs
    const next: [boolean, boolean] = c1 && r1 ? [false, true] : !c1 && r1 ? [true, false] : c1 && !r1 ? [false, false] : [true, true]
    const replaced = formatRefToken({ ...t.ref, abs: [next[0], next[1], next[0], next[1]] })
    const v = text.slice(0, t.start + 1) + replaced + text.slice(t.end + 1)
    this.editor.value = v
    const pos = t.start + 1 + replaced.length
    this.editor.el.setSelectionRange(pos, pos)
  }

  private syncEditorFromBar (): void {
    const input = this.barInput()
    if (!input) return
    this.editor.el.value = input.value
    this.editor.el.setSelectionRange(input.selectionStart ?? input.value.length, input.selectionEnd ?? input.value.length)
  }

  private barInput (): HTMLInputElement | null {
    return this.formulaBar.shadowRoot?.querySelector('input') ?? null
  }

  private focusBar (caret?: number): void {
    const input = this.barInput()
    if (!input) return
    input.focus()
    if (caret === undefined) return
    // film-input applies a new value on its next render; place the caret after it.
    void (this.formulaBar as unknown as { updateComplete: Promise<unknown> }).updateComplete.then(() => input.setSelectionRange(caret, caret))
  }

  // ===========================================================================
  // Applying changes (with undo)

  applyPatches (patches: CellPatch[], hints = false): void {
    if (this.host.readonly) return
    // Toolbar actions etc. can arrive while a cell is still being typed in:
    // keep what was typed (commitEdit closes the editor before re-entering).
    if (this.editor.active) this.commitEdit('none')
    if (!patches.length) return
    this.cancelCut()
    const selection = { ...this.grid.sel }
    const undo = this.wb.applyPatches(this.sheet.id, patches, hints)
    this.history.push({ kind: 'patch', sheetId: this.sheet.id, undo, redo: patches, selection })
    this.showPerf()
    this.onSelectionChange()
  }

  private undo (): void {
    const e = this.history.undo()
    if (!e) return this.announce('Nothing to undo')
    this.replay(e, 'undo')
    this.announce('Undone')
  }

  private redo (): void {
    const e = this.history.redo()
    if (!e) return this.announce('Nothing to redo')
    this.replay(e, 'redo')
    this.announce('Redone')
  }

  private replay (e: NonNullable<ReturnType<History['undo']>>, dir: 'undo' | 'redo'): void {
    if (this.editor.active) this.cancelEdit()
    this.cancelCut()
    if (e.kind === 'group') {
      const list = dir === 'undo' ? [...e.entries].reverse() : e.entries
      for (const sub of list) this.replay(sub, dir)
      return
    }
    if (e.kind === 'patch') {
      const sheet = this.wb.sheetSlot(e.sheetId)
      if (!sheet) return
      if (sheet !== this.sheet) this.switchSheet(sheet)
      this.wb.applyPatches(e.sheetId, dir === 'undo' ? e.undo : e.redo)
      this.grid.setSelection(e.selection)
      this.showPerf()
    } else if (e.kind === 'snapshot') {
      if (dir === 'undo') e.after = this.wb.toJSON()
      this.wb.load(dir === 'undo' ? e.before : e.after!)
      const id = dir === 'undo' ? e.sheetBefore : e.sheetAfter
      // Sheet ids survive load(), so this finds the same sheet as before.
      this.sheet = this.wb.sheetSlot(id) ?? this.wb.sheets[0]
      this.grid.setSheet(this.sheet)
      this.renderTabs()
      this.onSelectionChange()
    } else {
      const sheet = this.wb.sheetSlot(e.sheetId)
      if (!sheet) return
      if (sheet !== this.sheet) this.switchSheet(sheet)
      ;(e.axis === 'rows' ? this.grid.rows : this.grid.cols).setSize(e.index, dir === 'undo' ? e.before : e.after)
      this.grid.sizesChanged()
    }
    this.scheduleSave()
  }

  /** Clamp a (possibly whole-column) selection to where data actually is. */
  private usedSelection (): RangeAddress {
    const r = this.grid.range
    return {
      r1: r.r1,
      c1: r.c1,
      r2: Math.max(r.r1, Math.min(r.r2, this.sheet.maxRow)),
      c2: Math.max(r.c1, Math.min(r.c2, this.sheet.maxCol))
    }
  }

  private styleSelection (update: (s: Style) => Style): void {
    const r = this.usedSelection()
    const patches: CellPatch[] = []
    const area = (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1)
    if (area <= STYLE_CELL_LIMIT) {
      for (let row = r.r1; row <= r.r2; row++) {
        for (let col = r.c1; col <= r.c2; col++) {
          const cur = this.sheet.get(row, col)?.style ?? {}
          patches.push({ r: row, c: col, style: update({ ...cur }) })
        }
      }
    } else {
      // Huge, mostly empty selections: style only the cells that exist.
      for (const [c, col] of this.sheet.cols) {
        if (c < r.c1 || c > r.c2) continue
        for (const cell of col.values()) {
          if (cell.row >= r.r1 && cell.row <= r.r2) patches.push({ r: cell.row, c, style: update({ ...cell.style }) })
        }
      }
      if (patches.length > STYLE_CELL_LIMIT * 4) {
        toast('That selection is too large to format.', { variant: 'warning' })
        return
      }
    }
    this.applyPatches(patches)
  }

  private toggleStyle (key: 'bold' | 'italic' | 'underline' | 'strike'): void {
    const on = !this.activeCell()?.style?.[key]
    this.styleSelection((s) => ({ ...s, [key]: on || undefined }))
  }

  private setFormat (fmt: string | undefined): void {
    this.styleSelection((s) => ({ ...s, fmt }))
  }

  private adjustDecimals (delta: number): void {
    const cell = this.activeCell()
    let fmt = cell?.style?.fmt
    if (!fmt || fmt === 'General') {
      const shown = typeof cell?.value === 'number' ? formatGeneral(cell.value) : '0'
      const decimals = Math.max(0, (shown.split('.')[1]?.length ?? 0) + delta)
      fmt = decimals ? '0.' + '0'.repeat(decimals) : '0'
    } else if (/\.0+/.test(fmt)) {
      fmt = fmt.replace(/\.(0+)/, (_, z: string) => {
        const n = z.length + delta
        return n > 0 ? '.' + '0'.repeat(n) : ''
      })
    } else if (delta > 0) {
      fmt = fmt.replace(/([0#])(?![0#,])/, '$1.0')
    }
    this.setFormat(fmt)
  }

  private clearSelection (what: 'contents' | 'formats' | 'all'): void {
    const r = this.usedSelection()
    const patches: CellPatch[] = []
    const s = this.sheet
    // Only touch cells that exist — clearing a whole column stays cheap.
    for (const [c, col] of s.cols) {
      if (c < r.c1 || c > r.c2) continue
      for (const cell of col.values()) {
        if (cell.row < r.r1 || cell.row > r.r2) continue
        const p: CellPatch = { r: cell.row, c }
        if (what !== 'formats') p.raw = ''
        if (what !== 'contents') p.style = null
        patches.push(p)
      }
    }
    this.applyPatches(patches)
  }

  // ===========================================================================
  // Navigation

  private moveActive (dir: Dir): void {
    const r = this.grid.range
    const multi = r.r1 !== r.r2 || r.c1 !== r.c2
    const [dr, dc] = dirDelta(dir)
    let { ar, ac } = this.grid.sel
    if (multi && r.r2 - r.r1 < 100_000 && r.c2 - r.c1 < 1000) {
      // Enter/Tab cycle inside a selected block: Tab runs along rows and
      // wraps to the next row; Enter runs down columns and wraps to the next column.
      if (dc !== 0) {
        ac += dc
        if (ac > r.c2) { ac = r.c1; ar = ar + 1 > r.r2 ? r.r1 : ar + 1 }
        if (ac < r.c1) { ac = r.c2; ar = ar - 1 < r.r1 ? r.r2 : ar - 1 }
      } else {
        ar += dr
        if (ar > r.r2) { ar = r.r1; ac = ac + 1 > r.c2 ? r.c1 : ac + 1 }
        if (ar < r.r1) { ar = r.r2; ac = ac - 1 < r.c1 ? r.c2 : ac - 1 }
      }
      this.grid.setSelection({ ...this.grid.sel, ar, ac })
      this.grid.ensureVisible(ar, ac)
      return
    }
    ar = clamp(ar + dr, 0, MAX_ROWS - 1)
    ac = clamp(ac + dc, 0, MAX_COLS - 1)
    this.grid.setSelection({ ar, ac, fr: ar, fc: ac })
  }

  private filled (r: number, c: number): boolean {
    const cell = this.sheet.get(r, c)
    return !!cell && cell.raw !== ''
  }

  /** Ctrl+Arrow: jump to the edge of the current data block. */
  private jump (r: number, c: number, dir: Dir): [number, number] {
    const [dr, dc] = dirDelta(dir)
    const inBounds = (rr: number, cc: number): boolean => rr >= 0 && cc >= 0 && rr < MAX_ROWS && cc < MAX_COLS
    if (!inBounds(r + dr, c + dc)) return [r, c]
    const limitR = dr > 0 ? Math.max(this.sheet.maxRow, r) : 0
    const limitC = dc > 0 ? Math.max(this.sheet.maxCol, c) : 0
    if (this.filled(r, c) && this.filled(r + dr, c + dc)) {
      while (inBounds(r + dr, c + dc) && this.filled(r + dr, c + dc)) {
        r += dr
        c += dc
      }
      return [r, c]
    }
    r += dr
    c += dc
    while (!this.filled(r, c)) {
      // Past the used area there's nothing more to find: go to the sheet edge.
      if ((dr > 0 && r >= limitR) || (dc > 0 && c >= limitC) || (dr < 0 && r <= 0) || (dc < 0 && c <= 0)) {
        return [dr > 0 ? MAX_ROWS - 1 : dr < 0 ? 0 : r, dc > 0 ? MAX_COLS - 1 : dc < 0 ? 0 : c]
      }
      r += dr
      c += dc
    }
    return [r, c]
  }

  /** The contiguous block of data around the active cell (for sort / AutoSum). */
  private currentRegion (): RangeAddress {
    let { r1, c1, r2, c2 } = this.grid.range
    if (r1 !== r2 || c1 !== c2) return this.usedSelection()
    const s = this.sheet
    // Row checks look at a few columns; column checks walk that column's
    // sparse map (or the rows, whichever is smaller). Each edge grows in its
    // own loop, so a 100k-row block costs ~100k cheap steps, not 100k².
    const rowHas = (r: number, a: number, b: number): boolean => {
      if (r < 0 || r >= MAX_ROWS) return false
      for (let c = Math.max(0, a); c <= Math.min(b, MAX_COLS - 1); c++) if (this.filled(r, c)) return true
      return false
    }
    const colHas = (c: number, a: number, b: number): boolean => {
      const col = c >= 0 ? s.cols.get(c) : undefined
      if (!col) return false
      a = Math.max(0, a)
      if (b - a + 1 > col.size) {
        for (const cell of col.values()) if (cell.row >= a && cell.row <= b && cell.raw !== '') return true
        return false
      }
      for (let r = a; r <= b; r++) if (this.filled(r, c)) return true
      return false
    }
    for (let changed = true; changed;) {
      changed = false
      while (rowHas(r1 - 1, c1 - 1, c2 + 1)) { r1--; changed = true }
      while (rowHas(r2 + 1, c1 - 1, c2 + 1)) { r2++; changed = true }
      while (colHas(c1 - 1, r1 - 1, r2 + 1)) { c1--; changed = true }
      while (colHas(c2 + 1, r1 - 1, r2 + 1)) { c2++; changed = true }
    }
    return { r1, c1, r2, c2 }
  }

  // ===========================================================================
  // Clipboard

  private bindClipboard (): void {
    const gridHasFocus = (): boolean => !this.editor.active && this.grid.hasFocus()
    // Clipboard events are composed, so they reach the host from inside the shadow root.
    this.host.addEventListener('copy', (e) => {
      if (!gridHasFocus()) return
      e.preventDefault()
      this.copy(false, e.clipboardData!)
    })
    this.host.addEventListener('cut', (e) => {
      if (!gridHasFocus()) return
      e.preventDefault()
      this.copy(!this.host.readonly, e.clipboardData!)
    })
    this.host.addEventListener('paste', (e) => {
      if (!gridHasFocus() || this.host.readonly) return
      e.preventDefault()
      this.paste(e.clipboardData?.getData('text/plain') ?? '', false)
    })
  }

  private copy (cut: boolean, data?: DataTransfer): void {
    // Copy exactly what's selected, blanks included (pasting must clear those
    // cells too). Only huge selections like whole columns shrink to the data.
    const sel = this.grid.range
    const r = (sel.r2 - sel.r1 + 1) * (sel.c2 - sel.c1 + 1) <= 1_000_000 ? sel : this.usedSelection()
    if ((r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) > 2_000_000) {
      toast('That selection is too large to copy.', { variant: 'warning' })
      return
    }
    const cells: Clip['cells'] = []
    const rows: string[][] = []
    for (let row = r.r1; row <= r.r2; row++) {
      const line: Clip['cells'][number] = []
      const text: string[] = []
      for (let col = r.c1; col <= r.c2; col++) {
        const cell = this.sheet.get(row, col)
        line.push(cell ? { raw: cell.raw, style: cell.style, value: cell.value } : null)
        text.push(cell ? this.wb.getDisplay(cell) : '')
      }
      cells.push(line)
      rows.push(text)
    }
    const tsv = toDelimited(rows, '\t')
    this.clip = { sheetId: this.sheet.id, range: r, cells, text: tsv, cut }
    this.grid.setCopyRange(r)
    this.announce(`${cut ? 'Cut' : 'Copied'} ${formatRange(r)}`)
    const html = '<table>' + rows.map((row) => '<tr>' + row.map((t) => `<td>${escapeHtml(t)}</td>`).join('') + '</tr>').join('') + '</table>'
    if (data) {
      data.setData('text/plain', tsv)
      data.setData('text/html', html)
    } else {
      navigator.clipboard?.writeText(tsv).catch(() => {})
    }
  }

  private async pasteFromClipboardApi (valuesOnly: boolean): Promise<void> {
    try {
      const text = await navigator.clipboard.readText()
      this.paste(text, valuesOnly)
    } catch {
      // Without clipboard permission, still support pasting our own copy.
      if (this.clip) this.paste(this.clip.text, valuesOnly)
      else toast(`Use ${isMac ? '⌘' : 'Ctrl+'}V to paste from other apps.`, { variant: 'info' })
    }
  }

  private paste (text: string, valuesOnly: boolean): void {
    const target = this.grid.range
    const norm = (s: string): string => s.replace(/\r\n?/g, '\n').replace(/\n+$/, '')
    const clip = this.clip && norm(text) === norm(this.clip.text) ? this.clip : null
    let cells: Clip['cells']
    if (clip) cells = clip.cells
    else {
      const rows = parseDelimited(text.replace(/\r?\n$/, ''), '\t')
      cells = rows.map((row) => row.map((raw) => ({ raw, style: null, value: null })))
    }
    const h = cells.length
    const w = Math.max(0, ...cells.map((r) => r.length))
    if (!h || !w) return
    // Fill a larger selection with whole copies of the clipboard (Excel tiling).
    const th = target.r2 - target.r1 + 1
    const tw = target.c2 - target.c1 + 1
    const reps = th % h === 0 && tw % w === 0 && th * tw <= 1_000_000 ? [th / h, tw / w] : [1, 1]
    const patches: CellPatch[] = []
    const writtenKeys = new Set<number>()
    for (let i = 0; i < reps[0]; i++) {
      for (let j = 0; j < reps[1]; j++) {
        for (let dr = 0; dr < h; dr++) {
          for (let dc = 0; dc < w; dc++) {
            const src = cells[dr][dc] ?? null
            const r = target.r1 + i * h + dr
            const c = target.c1 + j * w + dc
            if (r >= MAX_ROWS || c >= MAX_COLS) continue
            let raw = src?.raw ?? ''
            if (clip && raw.startsWith('=')) {
              if (valuesOnly) raw = valueToRaw(src!.value)
              else if (!clip.cut) raw = shiftFormula(raw, r - (clip.range.r1 + dr), c - (clip.range.c1 + dc))
            }
            const patch: CellPatch = { r, c, raw }
            if (clip && !valuesOnly) patch.style = src?.style ?? null
            patches.push(patch)
            writtenKeys.add(r * MAX_COLS + c)
          }
        }
      }
    }
    const sameSheet = clip?.sheetId === this.sheet.id
    if (clip?.cut && sameSheet) {
      const sr = clip.range
      for (let r = sr.r1; r <= sr.r2; r++) {
        for (let c = sr.c1; c <= sr.c2; c++) if (!writtenKeys.has(r * MAX_COLS + c)) patches.unshift({ r, c, raw: '', style: null })
      }
    }
    this.applyPatches(patches, !clip)
    const source = clip?.cut && !sameSheet ? this.wb.sheetSlot(clip.sheetId) : undefined
    if (clip && source) {
      // Cut to another sheet: clear the source too, as one undo step.
      const clears: CellPatch[] = []
      for (let r = clip.range.r1; r <= clip.range.r2; r++) {
        for (let c = clip.range.c1; c <= clip.range.c2; c++) clears.push({ r, c, raw: '', style: null })
      }
      const undo = this.wb.applyPatches(source.id, clears)
      this.history.push({ kind: 'patch', sheetId: source.id, undo, redo: clears, selection: { ar: clip.range.r1, ac: clip.range.c1, fr: clip.range.r2, fc: clip.range.c2 } })
      this.history.group(2)
    }
    this.announce(`Pasted into ${formatRange({ r1: target.r1, c1: target.c1, r2: target.r1 + reps[0] * h - 1, c2: target.c1 + reps[1] * w - 1 })}`)
    if (clip?.cut) {
      this.clip = null
      this.grid.setCopyRange(null)
    }
    this.grid.setSelection({ ar: target.r1, ac: target.c1, fr: target.r1 + reps[0] * h - 1, fc: target.c1 + reps[1] * w - 1 })
  }

  // ===========================================================================
  // Fill, sort, structure

  private fill (src: RangeAddress, dest: RangeAddress, copyOnly = false): void {
    const patches = fillPatches(this.sheet, src, dest, copyOnly)
    this.applyPatches(patches)
    this.grid.setSelection({ ar: dest.r1, ac: dest.c1, fr: dest.r2, fc: dest.c2 }, false)
    this.announce(`Filled ${formatRange(dest)}`)
  }

  private fillDirection (dir: 'down' | 'right'): void {
    const before = { ...this.grid.sel }
    const r = this.usedSelection()
    if (dir === 'down') {
      if (r.r1 === r.r2) {
        if (r.r1 === 0) return
        this.fill({ ...r, r1: r.r1 - 1, r2: r.r1 - 1 }, { ...r, r1: r.r1 - 1 }, true)
      } else this.fill({ ...r, r2: r.r1 }, r, true)
    } else {
      if (r.c1 === r.c2) {
        if (r.c1 === 0) return
        this.fill({ ...r, c1: r.c1 - 1, c2: r.c1 - 1 }, { ...r, c1: r.c1 - 1 }, true)
      } else this.fill({ ...r, c2: r.c1 }, r, true)
    }
    // Ctrl+D / Ctrl+R leave the selection where it was.
    this.grid.setSelection(before, false)
  }

  private sort (ascending: boolean): void {
    let r = this.currentRegion()
    // Treat a text first row above numbers as a header.
    const first = [...span(r.c1, r.c2)].map((c) => this.sheet.get(r.r1, c)?.value)
    const second = [...span(r.c1, r.c2)].map((c) => this.sheet.get(r.r1 + 1, c)?.value)
    if (r.r2 > r.r1 && first.every((v) => typeof v === 'string' || v == null) && second.some((v) => typeof v === 'number')) r = { ...r, r1: r.r1 + 1 }
    const key = clamp(this.active.col, r.c1, r.c2)
    this.applyPatches(this.wb.sortPatches(this.sheet.id, r.r1, r.c1, r.r2, r.c2, key, ascending))
    this.grid.setSelection({ ar: r.r1, ac: key, fr: r.r2, fc: r.c2 === r.c1 ? r.c1 : r.c2 }, false)
    toast(`Sorted ${formatRange(r)} by column ${colToName(key)} ${ascending ? 'A → Z' : 'Z → A'}`, { variant: 'success', duration: 2500 })
  }

  private structural (kind: 'insertRows' | 'insertCols' | 'deleteRows' | 'deleteCols', where: 'before' | 'after' = 'before'): void {
    const r = this.grid.range
    const rows = kind.endsWith('Rows')
    const count = rows ? r.r2 - r.r1 + 1 : r.c2 - r.c1 + 1
    if ((rows && count === MAX_ROWS) || (!rows && count === MAX_COLS)) {
      toast(`Select specific ${rows ? 'rows' : 'columns'} first.`, { variant: 'warning' })
      return
    }
    const at = rows ? (where === 'after' ? r.r2 + 1 : r.r1) : where === 'after' ? r.c2 + 1 : r.c1
    this.snapshotAction(() => this.wb[kind](this.sheet.id, at, count))
    this.grid.sizesChanged()
    this.onSelectionChange()
  }

  private autoSum (): void {
    const r = this.grid.range
    if (r.r1 === r.r2 && r.c1 === r.c2) {
      const { row, col } = this.active
      const isNum = (rr: number, cc: number): boolean => typeof this.sheet.get(rr, cc)?.value === 'number'
      let top = row - 1
      while (top >= 0 && isNum(top, col)) top--
      let formula = '=SUM()'
      if (top < row - 1) formula = `=SUM(${formatAddress(top + 1, col)}:${formatAddress(row - 1, col)})`
      else {
        let left = col - 1
        while (left >= 0 && isNum(row, left)) left--
        if (left < col - 1) formula = `=SUM(${formatAddress(row, left + 1)}:${formatAddress(row, col - 1)})`
      }
      this.startEdit(formula, 'edit')
      if (formula === '=SUM()') this.editor.el.setSelectionRange(5, 5)
      return
    }
    // A block: total each column underneath it.
    const u = this.usedSelection()
    const patches: CellPatch[] = []
    for (let c = u.c1; c <= u.c2; c++) patches.push({ r: u.r2 + 1, c, raw: `=SUM(${formatAddress(u.r1, c)}:${formatAddress(u.r2, c)})` })
    this.applyPatches(patches)
  }

  // ===========================================================================
  // Find & replace

  private findMatches (): Cell[] {
    const q = this.$<ValueEl>('.find-input').value
    if (!q) return []
    const matchCase = this.$<CheckEl>('.find-case').checked
    const whole = this.$<CheckEl>('.find-whole').checked
    const inFormulas = this.$<CheckEl>('.find-formulas').checked
    const needle = matchCase ? q : q.toLowerCase()
    const out: Cell[] = []
    for (const cell of this.sheet.cells()) {
      let hay = inFormulas ? cell.raw : this.wb.getDisplay(cell)
      if (!hay) continue
      if (!matchCase) hay = hay.toLowerCase()
      if (whole ? hay === needle : hay.includes(needle)) out.push(cell)
    }
    return out.sort((a, b) => a.row - b.row || a.col - b.col)
  }

  private findNext (): void {
    const matches = this.findMatches()
    const status = this.$('.find-status')
    if (!matches.length) {
      status.textContent = 'No matches'
      return
    }
    const { row, col } = this.active
    let i = matches.findIndex((c) => c.row > row || (c.row === row && c.col > col))
    if (i < 0) i = 0
    const m = matches[i]
    this.grid.setSelection({ ar: m.row, ac: m.col, fr: m.row, fc: m.col })
    status.textContent = `Match ${i + 1} of ${matches.length.toLocaleString()} — ${formatAddress(m.row, m.col)}`
  }

  private replace (all: boolean): void {
    const q = this.$<ValueEl>('.find-input').value
    const rep = this.$<ValueEl>('.replace-input').value
    if (!q) return
    const matchCase = this.$<CheckEl>('.find-case').checked
    const whole = this.$<CheckEl>('.find-whole').checked
    const re = new RegExp(whole ? `^${escapeRe(q)}$` : escapeRe(q), matchCase ? 'g' : 'gi')
    // Replace rewrites what was typed (raw). Formulas are only touched when
    // searching in formulas: matching "1" in a displayed result must not
    // turn =A1*10 into =A2*20.
    const inFormulas = this.$<CheckEl>('.find-formulas').checked
    const matches = this.findMatches().filter((c) => inFormulas || !c.formula)
    const active = this.activeCell()
    const targets = all ? matches : matches.filter((c) => c === active)
    const patches = targets.map((c) => ({ r: c.row, c: c.col, raw: c.raw.replace(re, () => rep) })).filter((p, i) => p.raw !== targets[i].raw)
    this.applyPatches(patches, true)
    this.$('.find-status').textContent = `Replaced ${patches.length.toLocaleString()} cell${patches.length === 1 ? '' : 's'}`
    if (!all) this.findNext()
  }

  // ===========================================================================
  // Files & persistence

  private scheduleSave (): void {
    clearTimeout(this.saveTimer)
    if (!this.host.storageKey) return
    this.saveState.textContent = 'Saving…'
    this.saveTimer = window.setTimeout(() => this.save(), 700)
  }

  private save (): void {
    const cells = this.wb.sheets.reduce((n, s) => n + s.cellCount, 0)
    if (cells > AUTOSAVE_CELL_LIMIT) {
      this.saveState.textContent = 'Autosave paused (large workbook)'
      return
    }
    try {
      localStorage.setItem(this.host.storageKey, JSON.stringify({ title: this.host.filename, active: this.wb.sheets.indexOf(this.sheet), workbook: this.wb.toJSON() }))
      this.saveState.textContent = 'Saved'
    } catch {
      this.saveState.textContent = 'Autosave unavailable'
    }
  }

  /** Load the autosaved workbook; returns the saved active sheet index, or null. */
  private restore (): number | null {
    try {
      if (!this.host.storageKey) return null
      const raw = localStorage.getItem(this.host.storageKey)
      if (!raw) return null
      const data = JSON.parse(raw) as { title?: string; active?: number; workbook: WorkbookJSON }
      this.wb.load(data.workbook)
      if (data.title) this.host.filename = data.title
      return data.active ?? 0
    } catch {
      return null
    }
  }

  private loadWorkbook (json: WorkbookJSON, title?: string): void {
    if (this.editor.active) this.cancelEdit()
    this.wb.load(json)
    this.history.clear()
    this.clip = null
    this.grid.setCopyRange(null)
    this.sheet = this.wb.sheets[0]
    this.grid.setSheet(this.sheet)
    this.grid.setSelection({ ar: 0, ac: 0, fr: 0, fc: 0 })
    if (title) this.host.filename = title
    this.renderTabs()
    this.scheduleSave()
  }

  private openFile (): void {
    const input = this.$<HTMLInputElement>('.file-input')
    input.value = ''
    // A property, not addEventListener: each Open replaces the previous handler,
    // and a cancelled picker never fires change, so listeners would pile up.
    // oxlint-disable-next-line unicorn/prefer-add-event-listener
    input.onchange = async () => {
      const file = input.files?.[0]
      if (!file) return
      const text = await file.text()
      const base = file.name.replace(/\.(json|csv|tsv|txt)$/i, '').replace(/\.filmsheet$/i, '')
      try {
        if (/\.json$/i.test(file.name)) {
          this.loadWorkbook(JSON.parse(text) as WorkbookJSON, base)
        } else {
          const t0 = performance.now()
          const rows = parseDelimited(text, /\.tsv$/i.test(file.name) ? '\t' : undefined)
          const cells: [number, number, string][] = []
          rows.forEach((row, r) => row.forEach((raw, c) => raw !== '' && r < MAX_ROWS && c < MAX_COLS && cells.push([r, c, raw])))
          this.loadWorkbook({ version: 1, sheets: [{ name: base.slice(0, 31) || 'Sheet1', cells }] }, base)
          toast(`Loaded ${cells.length.toLocaleString()} cells in ${(performance.now() - t0).toFixed(0)} ms`, { variant: 'success' })
        }
      } catch (err) {
        toast(`Couldn't open ${file.name}: ${(err as Error).message}`, { variant: 'danger', duration: 0 })
      }
    }
    input.click()
  }

  private exportCsv (): void {
    const s = this.sheet
    // Walk the sparse maps: one far-away cell mustn't make this rows × 16k.
    const rows: string[][] = []
    for (const [c, col] of s.cols) {
      for (const cell of col.values()) {
        const text = this.wb.getDisplay(cell)
        if (!text) continue
        const row = (rows[cell.row] ??= [])
        while (row.length < c) row.push('')
        row[c] = csvSafe(text)
      }
    }
    for (let r = 0; r < rows.length; r++) {
      const row = (rows[r] ??= [])
      for (let c = 0; c < row.length; c++) row[c] ??= ''
    }
    download(`${this.host.filename || 'sheet'} - ${s.name}.csv`, toDelimited(rows, ','), 'text/csv')
  }

  // ===========================================================================
  // Dialogs & menus

  private prompt (title: string, label: string, value: string): Promise<string | null> {
    const dialog = this.$<DialogEl>('#prompt-dialog')
    const input = this.$<ValueEl & { label: string }>('.prompt-input')
    dialog.label = title
    input.label = label
    input.value = value
    dialog.show()
    requestAnimationFrame(() => {
      const inner = input.shadowRoot?.querySelector('input')
      inner?.focus()
      inner?.select()
    })
    return new Promise((resolve) => {
      this.promptResolve = resolve
    })
  }

  private finishPrompt (ok: boolean): void {
    const dialog = this.$<DialogEl>('#prompt-dialog')
    const resolve = this.promptResolve
    this.promptResolve = null
    const v = this.$<ValueEl>('.prompt-input').value
    dialog.close()
    resolve?.(ok ? v : null)
    this.grid.focus()
  }

  private showMenu (sel: string, x: number, y: number): void {
    const menu = this.$(sel)
    if (!menu.dataset.bound) {
      menu.dataset.bound = '1'
      menu.addEventListener('toggle', (e) => {
        if (e.newState === 'closed') {
          requestAnimationFrame(() => {
            if (this.focusLost() || menu.contains(this.root.activeElement)) this.grid.focus()
          })
        }
      })
    }
    menu.style.left = Math.min(x, window.innerWidth - 240) + 'px'
    menu.style.top = Math.min(y, window.innerHeight - 420) + 'px'
    menu.showPopover()
    requestAnimationFrame(() => menu.querySelector<HTMLElement & { focusFirst (): void }>('film-menu')?.focusFirst())
  }

  private buildShortcuts (): void {
    const mod = isMac ? '⌘' : 'Ctrl'
    const rows: [string, string][] = [
      ['Arrow keys', 'Move'], ['Shift + Arrow', 'Extend selection'], [`${mod} + Arrow`, 'Jump to edge of data'],
      [`${mod} + Shift + Arrow`, 'Select to edge of data'], ['Enter / Shift + Enter', 'Commit and move down / up'],
      ['Tab / Shift + Tab', 'Commit and move right / left'], [`${mod} + Enter`, 'Fill the selection with the entry'],
      ['F2', 'Edit cell (toggle Enter/Edit mode)'], ['F4', 'Cycle $ absolute references while editing'],
      ['Escape', 'Cancel editing'], ['Delete / Backspace', 'Clear contents'], [`${mod} + C / X / V`, 'Copy / cut / paste'],
      [`${mod} + Shift + V`, 'Paste values only'], [`${mod} + Z / ${mod} + Y`, 'Undo / redo'], [`${mod} + B / I / U`, 'Bold / italic / underline'],
      [`${mod} + D / ${mod} + R`, 'Fill down / right'], [`${mod} + A`, 'Select all'], ['Shift + Space / Ctrl + Space', 'Select row / column'],
      [`${mod} + F`, 'Find and replace'], [`${mod} + G`, 'Go to cell'], [`${mod} + Home / End`, 'First / last used cell'],
      [`${isMac ? '⌥' : 'Alt'} + =`, 'AutoSum'], [`${mod} + \``, 'Show formulas'], ['Ctrl + PageUp / PageDown', 'Previous / next sheet'],
      ['Shift + F11', 'New sheet'], ['F6 / Shift + F6', 'Move between menu bar, toolbar, formula bar, grid and sheet tabs'],
      ['F10', 'Go to the menu bar']
    ]
    this.$('.shortcuts').innerHTML = '<table>' + rows.map(([k, d]) => `<tr><td>${k.split(' + ').map((p) => `<film-kbd>${escapeHtml(p)}</film-kbd>`).join(' + ')}</td><td>${escapeHtml(d)}</td></tr>`).join('') + '</table>'
  }

  private renderFunctionList (filter = ''): void {
    const q = filter.trim().toUpperCase()
    const names = FUNCTION_NAMES.filter((n) => !q || n.includes(q) || FUNCTIONS[n].desc.toUpperCase().includes(q))
    this.$('.fn-list').innerHTML = names
      .map((n) => `<button class="fn-row" data-fn="${n}"><code>${escapeHtml(FUNCTIONS[n].sig)}</code><span>${escapeHtml(FUNCTIONS[n].desc)}</span></button>`)
      .join('') || '<p>No matching functions.</p>'
  }

  // ===========================================================================
  // Commands

  run (cmd: string): void {
    const [group] = cmd.split('.')
    if (this.host.readonly && !READONLY_COMMANDS.has(cmd)) {
      this.announce('This sheet is read-only')
      return
    }
    if (this.editor.active && group !== 'edit' && group !== 'help') this.commitEdit('none')
    switch (cmd) {
      case 'file.new':
        this.loadWorkbook({ version: 1, sheets: [{ name: 'Sheet1', cells: [] }] }, 'Untitled spreadsheet')
        break
      case 'file.open': this.openFile(); break
      case 'file.save':
        download(`${this.host.filename || 'workbook'}.filmsheet.json`, JSON.stringify(this.wb.toJSON()), 'application/json')
        break
      case 'file.exportCsv': this.exportCsv(); break
      case 'edit.undo': this.undo(); break
      case 'edit.redo': this.redo(); break
      case 'edit.copy': this.copy(false); break
      case 'edit.cut': this.copy(true); break
      case 'edit.paste': void this.pasteFromClipboardApi(false); break
      case 'edit.pasteValues': void this.pasteFromClipboardApi(true); break
      case 'edit.clear':
        this.clearSelection('contents')
        this.announce('Contents cleared')
        break
      case 'edit.clearFormats':
      case 'format.clear': this.clearSelection('formats'); break
      case 'edit.clearAll': this.clearSelection('all'); break
      case 'edit.fillDown': this.fillDirection('down'); break
      case 'edit.seriesDown':
      case 'edit.seriesRight': {
        // Keyboard equivalent of dragging the fill handle: extend the first
        // row (or column) of the selection across the rest as a series.
        const r = this.grid.range
        const down = cmd === 'edit.seriesDown'
        if (down ? r.r1 === r.r2 : r.c1 === r.c2) {
          toast(`Select the starting cells and the ${down ? 'rows below' : 'columns to the right'} to fill.`, { variant: 'info' })
          break
        }
        const u = this.usedSelection()
        const dest = down ? { ...u, r2: r.r2 } : { ...u, c2: r.c2 }
        // Use the leading filled cells as the pattern (one or more).
        let end = down ? r.r1 : r.c1
        while (end + 1 < (down ? r.r2 : r.c2) && (down ? this.filled(end + 1, r.c1) : this.filled(r.r1, end + 1))) end++
        this.fill(down ? { ...dest, r2: end } : { ...dest, c2: end }, dest)
        break
      }
      case 'edit.fillRight': this.fillDirection('right'); break
      case 'edit.selectAll': this.grid.setSelection({ ar: 0, ac: 0, fr: MAX_ROWS - 1, fc: MAX_COLS - 1 }, false); break
      case 'edit.deleteRows': this.structural('deleteRows'); break
      case 'edit.deleteCols': this.structural('deleteCols'); break
      case 'edit.find': {
        this.$<DialogEl>('#find-dialog').show()
        this.$('.find-status').textContent = ''
        requestAnimationFrame(() => this.$('.find-input').shadowRoot?.querySelector('input')?.focus())
        return
      }
      case 'view.formulas':
        this.showFormulas = !this.showFormulas
        this.grid.showFormulas = this.showFormulas
        this.grid.invalidate()
        break
      case 'view.goto': {
        const inner = this.nameBox.shadowRoot?.querySelector('input')
        inner?.focus()
        inner?.select()
        return
      }
      case 'insert.rowAbove': this.structural('insertRows', 'before'); break
      case 'insert.rowBelow': this.structural('insertRows', 'after'); break
      case 'insert.colLeft': this.structural('insertCols', 'before'); break
      case 'insert.colRight': this.structural('insertCols', 'after'); break
      case 'insert.sheet': {
        this.snapshotAction(() => {
          this.sheet = this.wb.addSheet(undefined, this.wb.sheets.indexOf(this.sheet) + 1)
        })
        this.grid.setSheet(this.sheet)
        this.renderTabs()
        this.grid.setSelection({ ar: 0, ac: 0, fr: 0, fc: 0 })
        break
      }
      case 'insert.sum': this.autoSum(); return
      case 'sheet.rename': void this.renameSheet(); return
      case 'sheet.delete': {
        if (this.wb.sheets.length === 1) {
          toast('A workbook needs at least one sheet.', { variant: 'warning' })
          break
        }
        const idx = this.wb.sheets.indexOf(this.sheet)
        const name = this.sheet.name
        this.snapshotAction(() => {
          this.wb.removeSheet(this.sheet.id)
          this.sheet = this.wb.sheets[Math.max(0, idx - 1)]
        })
        this.grid.setSheet(this.sheet)
        this.renderTabs()
        this.onSelectionChange()
        toast(`Deleted "${name}". Undo with ${isMac ? '⌘' : 'Ctrl+'}Z.`, { duration: 4000 })
        break
      }
      case 'sheet.duplicate': {
        const json = this.wb.toJSON()
        const idx = this.wb.sheets.indexOf(this.sheet)
        const copy = structuredClone(json.sheets[idx])
        delete copy.id
        let n = 2
        while (this.wb.sheetByName(`${this.sheet.name} (${n})`)) n++
        copy.name = `${this.sheet.name} (${n})`
        json.sheets.splice(idx + 1, 0, copy)
        this.snapshotAction(() => {
          this.wb.load(json)
          this.sheet = this.wb.sheets[idx + 1]
        })
        this.grid.setSheet(this.sheet)
        this.renderTabs()
        break
      }
      case 'sheet.moveLeft':
      case 'sheet.moveRight': {
        const idx = this.wb.sheets.indexOf(this.sheet)
        this.wb.moveSheet(this.sheet.id, idx + (cmd === 'sheet.moveLeft' ? -1 : 1))
        this.renderTabs()
        break
      }
      case 'format.bold': this.toggleStyle('bold'); break
      case 'format.italic': this.toggleStyle('italic'); break
      case 'format.underline': this.toggleStyle('underline'); break
      case 'format.strike': this.toggleStyle('strike'); break
      case 'format.alignLeft': this.styleSelection((s) => ({ ...s, align: s.align === 'left' ? undefined : 'left' })); break
      case 'format.alignCenter': this.styleSelection((s) => ({ ...s, align: s.align === 'center' ? undefined : 'center' })); break
      case 'format.alignRight': this.styleSelection((s) => ({ ...s, align: s.align === 'right' ? undefined : 'right' })); break
      case 'format.colWidth':
      case 'format.rowHeight':
        void this.promptSize(cmd === 'format.colWidth' ? 'cols' : 'rows')
        return
      case 'format.autofit': {
        const r = this.grid.range
        for (let c = r.c1; c <= Math.min(r.c2, r.c1 + 200); c++) this.setSize('cols', c, this.grid.measureColumn(c))
        this.announce('Column width fitted to contents')
        break
      }
      case 'format.decInc': this.adjustDecimals(1); break
      case 'format.decDec': this.adjustDecimals(-1); break
      case 'data.sortAsc': this.sort(true); break
      case 'data.sortDesc': this.sort(false); break
      case 'help.shortcuts':
        this.buildShortcuts()
        this.$<DialogEl>('#shortcuts-dialog').show()
        return
      case 'help.functions':
        this.renderFunctionList()
        this.$<DialogEl>('#functions-dialog').show()
        return
      default:
        if (cmd.startsWith('format.num.')) {
          const key = cmd.slice('format.num.'.length) as keyof typeof FORMATS
          this.setFormat(key === 'general' ? undefined : FORMATS[key])
        }
    }
    if (!this.editor.active) this.grid.focus()
  }

  // ===========================================================================
  // Wiring

  private bindChrome (): void {
    this.root.addEventListener('film-select', (e) => {
      const value = (e as CustomEvent<{ value?: string }>).detail?.value
      if (!value) return
      for (const m of this.root.querySelectorAll<HTMLElement>('.context-menu:popover-open')) m.hidePopover()
      this.run(value)
    })
    this.root.addEventListener('click', (e) => {
      const target = (e.target as HTMLElement).closest<HTMLElement>('[data-cmd]')
      if (target && !target.closest('film-menu')) this.run(target.dataset.cmd!)
    })

    this.bindToolbar()

    // Formula bar: typing here edits the active cell.
    this.formulaBar.addEventListener('input', () => {
      const text = this.formulaBar.value
      if (!this.editor.active) this.startEdit(text, 'edit', 'bar')
      else {
        this.editSource = 'bar'
        this.editor.stopPointing()
        this.editor.value = text
        this.syncEditorFromBar()
      }
    })
    this.formulaBar.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault()
        if (this.editor.active) this.commitEdit(e.shiftKey ? 'up' : 'down', e.ctrlKey || e.metaKey)
        else this.grid.focus()
      } else if (e.key === 'Escape') {
        e.preventDefault()
        if (this.editor.active) this.cancelEdit()
        else this.grid.focus()
      } else if (e.key === 'Tab' && this.editor.active) {
        // While editing, Tab commits like in the cell; otherwise it moves focus normally.
        e.preventDefault()
        this.commitEdit(e.shiftKey ? 'left' : 'right')
      }
      if (e.key !== 'F6' && e.key !== 'F10') e.stopPropagation()
    })

    // Name box: jump to a cell or range.
    this.nameBox.addEventListener('keydown', (e) => {
      e.stopPropagation()
      if (e.key === 'Escape') {
        this.onSelectionChange()
        this.grid.focus()
      }
      if (e.key !== 'Enter') return
      const r = parseRangeAddress(this.nameBox.value)
      if (!r) {
        toast(`"${this.nameBox.value}" isn't a cell reference.`, { variant: 'warning' })
        return
      }
      this.grid.setSelection({ ar: r.r1, ac: r.c1, fr: r.r2, fc: r.c2 })
      this.grid.ensureVisible(r.r1, r.c1)
      this.grid.focus()
    })

    // Sheet tabs
    this.tabs.addEventListener('film-tab-change', (e) => {
      const id = Number((e as CustomEvent<{ name: string }>).detail.name)
      const s = this.wb.sheetSlot(id)
      if (s && s !== this.sheet) this.switchSheet(s)
    })
    this.tabs.addEventListener('click', (e) => {
      if ((e.target as HTMLElement).closest('film-tab') && e.detail === 1) requestAnimationFrame(() => this.grid.focus())
    })
    this.tabs.addEventListener('keydown', (e) => {
      if (e.key === 'F2' && (e.target as HTMLElement).closest('film-tab')) {
        e.preventDefault()
        void this.renameSheet()
      }
    })
    this.tabs.addEventListener('dblclick', (e) => {
      if ((e.target as HTMLElement).closest('film-tab')) void this.renameSheet()
    })
    this.tabs.addEventListener('contextmenu', (e) => {
      const tab = (e.target as HTMLElement).closest<HTMLElement>('film-tab')
      if (!tab) return
      e.preventDefault()
      const s = this.wb.sheetSlot(Number(tab.getAttribute('panel')))
      if (s && s !== this.sheet) this.switchSheet(s)
      this.showMenu('#tab-menu', e.clientX, e.clientY - 180)
    })

    // Number format select
    this.formatSelect.addEventListener('change', () => {
      const v = this.formatSelect.value as keyof typeof FORMATS | 'custom'
      if (v === 'custom') return
      this.setFormat(v === 'general' ? undefined : FORMATS[v])
      this.grid.focus()
    })

    // Colour swatches
    for (const dd of this.root.querySelectorAll<HTMLElement & { open: boolean }>('.color-dropdown')) {
      const target = dd.dataset.target as 'color' | 'fill'
      const box = this.$('.swatches', dd)
      const reset = document.createElement('button')
      reset.className = 'swatch-reset'
      reset.textContent = target === 'color' ? 'Automatic' : 'No fill'
      reset.dataset.color = ''
      reset.addEventListener('click', () => {
        this.styleSelection((s) => ({ ...s, [target]: undefined }))
        dd.open = false
        this.grid.focus()
      })
      box.append(reset)
      for (const [color, name] of SWATCHES) {
        const b = document.createElement('button')
        b.className = 'swatch'
        b.style.background = color
        b.title = name
        b.dataset.color = color
        b.setAttribute('aria-label', name)
        b.addEventListener('click', () => {
          this.styleSelection((s) => ({ ...s, [target]: color }))
          // CSS sets the bar's fill, so set it as a style, not an attribute.
          dd.querySelector<SVGRectElement>('.swatch-bar')?.style.setProperty('fill', color)
          dd.open = false
          this.grid.focus()
        })
        box.append(b)
      }
      // Mark the active cell's current colour when the palette opens.
      dd.addEventListener('film-open', () => {
        const current = this.activeCell()?.style?.[target] ?? ''
        for (const b of box.querySelectorAll<HTMLElement>('button')) b.setAttribute('aria-pressed', String(b.dataset.color === current))
        box.querySelector<HTMLElement>('[aria-pressed="true"]')?.focus()
      })
      // Escape / click-away: put focus back in the grid rather than on <body>.
      dd.addEventListener('film-close', () => requestAnimationFrame(() => {
        if (this.focusLost()) this.grid.focus()
      }))
    }

    for (const q of ['(prefers-color-scheme: dark)', '(forced-colors: active)', '(prefers-contrast: more)']) {
      const mq = matchMedia(q)
      const refresh = (): number => requestAnimationFrame(() => this.grid.refreshTheme())
      mq.addEventListener('change', refresh)
      this.cleanups.push(() => mq.removeEventListener('change', refresh))
    }

    // Dialogs
    this.$('.prompt-ok').addEventListener('click', () => this.finishPrompt(true))
    this.$('.prompt-cancel').addEventListener('click', () => this.finishPrompt(false))
    this.$('#prompt-dialog').addEventListener('film-close', () => {
      if (this.promptResolve) this.finishPrompt(false)
    })
    // Enter in a text field submits; on a button it must activate that button.
    const fromTextField = (e: KeyboardEvent): boolean => e.composedPath()[0] instanceof HTMLInputElement && (e.composedPath()[0] as HTMLInputElement).type !== 'checkbox'
    this.$('#prompt-dialog').addEventListener('keydown', (e) => {
      e.stopPropagation()
      if (e.key === 'Enter' && fromTextField(e)) {
        e.preventDefault()
        this.finishPrompt(true)
      }
    })
    this.$('.find-next').addEventListener('click', () => this.findNext())
    this.$('.find-replace').addEventListener('click', () => this.replace(false))
    this.$('.find-replace-all').addEventListener('click', () => this.replace(true))
    this.$('#find-dialog').addEventListener('keydown', (e) => {
      e.stopPropagation()
      if (e.key === 'Enter' && fromTextField(e)) {
        e.preventDefault()
        this.findNext()
      }
    })
    this.$('.fn-search').addEventListener('input', () => this.renderFunctionList(this.$<ValueEl>('.fn-search').value))
    this.$('.fn-list').addEventListener('click', (e) => {
      const name = (e.target as HTMLElement).closest<HTMLElement>('[data-fn]')?.dataset.fn
      if (!name) return
      this.$<DialogEl>('#functions-dialog').close()
      if (this.editor.active) {
        this.editor.insert(name + '(')
        this.editor.el.focus()
      } else this.startEdit(`=${name}(`, 'enter')
    })
    for (const d of this.root.querySelectorAll('film-dialog')) {
      d.addEventListener('film-close', () => requestAnimationFrame(() => !this.editor.active && this.grid.focus()))
      d.addEventListener('keydown', (e) => e.stopPropagation())
    }
  }

  /** Menu hints are written with ⌃; show ⌘ on Macs and spelled-out keys elsewhere. */
  private localizeShortcuts (): void {
    for (const kbd of this.root.querySelectorAll('film-menu-item film-kbd')) {
      const t = kbd.textContent ?? ''
      kbd.textContent = isMac ? t.replace('⌃', '⌘') : t.replace('⌃', 'Ctrl+').replace('⌥', 'Alt+').replace('⇧', 'Shift+')
    }
  }

  /**
   * role="toolbar" pattern: one tab stop, arrow keys move between controls.
   * The real buttons live in Film's shadow roots, so tabindex is managed there.
   */
  private bindToolbar (): void {
    const toolbar = this.$('.toolbar')
    const hosts = [...toolbar.querySelectorAll<HTMLElement & { updateComplete: Promise<unknown>; label?: string }>('film-icon-button, film-select, film-button')]
    const buttons = (): HTMLButtonElement[] => hosts.map((h) => h.shadowRoot?.querySelector('button')).filter((b): b is HTMLButtonElement => !!b)
    void Promise.all(hosts.map((h) => h.updateComplete)).then(() => {
      buttons().forEach((b, i) => { b.tabIndex = i === 0 ? 0 : -1 })
      // Visible names for icon-only buttons on hover.
      for (const h of hosts) if (h.label) h.title = h.label
    })
    toolbar.addEventListener('focusin', (e) => {
      const target = e.composedPath()[0]
      const list = buttons()
      if (!list.includes(target as HTMLButtonElement)) return
      for (const b of list) b.tabIndex = b === target ? 0 : -1
    })
    toolbar.addEventListener('keydown', (e) => {
      if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(e.key)) return
      const list = buttons()
      const i = list.indexOf(e.composedPath()[0] as HTMLButtonElement)
      if (i < 0) return
      e.preventDefault()
      const next = e.key === 'Home' ? 0 : e.key === 'End' ? list.length - 1 : (i + (e.key === 'ArrowRight' ? 1 : -1) + list.length) % list.length
      list[next].focus()
    })
  }

  /** The page regions F6 cycles through, in order. */
  private regions (): { name: string; selector: string; has: (el: Element) => boolean; focus: () => void }[] {
    const focusIn = (sel: string): void => {
      const host = this.$(sel)
      const target = host.matches('film-input') ? host.shadowRoot?.querySelector('input') : host
      ;(target as HTMLElement | null)?.focus()
    }
    return [
      {
        name: 'Menu bar',
        selector: 'film-menu-bar',
        has: (el) => !!el.closest('film-menu-bar'),
        focus: () => this.root.querySelector<HTMLElement & { focus (): void }>('film-menu-bar-item:not([disabled])')?.focus()
      },
      {
        name: 'Toolbar',
        selector: '.toolbar',
        has: (el) => !!el.closest('.toolbar'),
        focus: () => {
          const first = this.root.querySelector<HTMLElement>('.toolbar film-icon-button')
          ;(first?.shadowRoot?.querySelector('button') ?? first)?.focus()
        }
      },
      { name: 'Formula bar', selector: '.formula-bar', has: (el) => !!el.closest('.formula-bar'), focus: () => focusIn('.formula-input') },
      { name: 'Grid', selector: '.grid-host', has: () => this.grid.hasFocus() || this.editor.active, focus: () => (this.editor.active ? this.editor.el.focus() : this.grid.focus()) },
      {
        name: 'Sheet tabs',
        selector: '.sheet-bar',
        has: (el) => !!el.closest('.sheet-bar'),
        focus: () => this.tabs.querySelector<HTMLElement>('film-tab[active]')?.focus()
      }
    ]
  }

  private cycleRegion (dir: 1 | -1): void {
    // Chrome hidden by the `chrome` attribute isn't a region.
    const regions = this.regions().filter((r) => r.name === 'Grid' || this.root.querySelector<HTMLElement>(r.selector)?.checkVisibility())
    const active = this.root.activeElement ?? document.body
    // Shadow roots hide the real target; activeElement is the light-DOM host.
    let i = regions.findIndex((r) => r.has(active))
    if (i < 0) i = regions.findIndex((r) => r.name === 'Grid')
    const next = regions[(i + dir + regions.length) % regions.length]
    next.focus()
    this.announce(next.name)
  }

  private bindKeyboard (): void {
    // Listen on our own shadow root: keys typed elsewhere on the page are never ours.
    this.root.addEventListener('keydown', (ev) => {
      const e = ev as KeyboardEvent
      if ((e.target as HTMLElement).closest?.('film-dialog')) return
      if (e.key === 'F6') {
        e.preventDefault()
        e.stopPropagation()
        this.cycleRegion(e.shiftKey ? -1 : 1)
      } else if (e.key === 'F10' && !e.shiftKey) {
        e.preventDefault()
        e.stopPropagation()
        const menu = this.regions()[0]
        if (this.root.querySelector<HTMLElement>(menu.selector)?.checkVisibility()) menu.focus()
      }
    }, { capture: true })
    this.editor.el.addEventListener('keydown', (e) => this.onEditorKey(e))
    this.root.addEventListener('keydown', (ev) => {
      const e = ev as KeyboardEvent
      if (e.defaultPrevented) return
      const t = e.target as HTMLElement
      if (t === this.editor.el) return
      // Leave typing in other inputs and dialogs alone.
      if (t.closest('film-dialog, input, textarea, film-input, film-select, film-search, .context-menu, film-menu-bar, .toolbar, .sheet-bar')) return
      this.onGridKey(e)
    })
  }

  private onEditorKey (e: KeyboardEvent): void {
    const ed = this.editor
    const mod = e.ctrlKey || e.metaKey
    if (ed.suggesting) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault()
        ed.moveSuggestion(e.key === 'ArrowDown' ? 1 : -1)
        return
      }
      // Tab always takes the suggestion; Enter only once the user has picked
      // one with the arrows — otherwise "=SUM(A:A" + Enter must commit.
      if (e.key === 'Tab' || (e.key === 'Enter' && !mod && ed.navigated)) {
        e.preventDefault()
        ed.acceptSuggestion()
        return
      }
      if (e.key === 'Enter') ed.hidePopup()
      if (e.key === 'Escape') {
        e.preventDefault()
        ed.hidePopup()
        return
      }
    }
    switch (e.key) {
      case 'Enter':
        e.preventDefault()
        this.commitEdit(e.shiftKey ? 'up' : 'down', mod)
        return
      case 'Tab':
        e.preventDefault()
        this.commitEdit(e.shiftKey ? 'left' : 'right')
        return
      case 'Escape':
        e.preventDefault()
        this.cancelEdit()
        return
      case 'F2':
        e.preventDefault()
        ed.mode = ed.mode === 'enter' ? 'edit' : 'enter'
        this.setMode(ed.mode === 'enter' ? 'Enter' : 'Edit')
        this.announce(ed.mode === 'enter' ? 'Enter mode: arrow keys move between cells' : 'Edit mode: arrow keys move the cursor')
        return
      case 'F4':
        e.preventDefault()
        this.toggleAbsolute()
        return
      case 'ArrowUp':
      case 'ArrowDown':
      case 'ArrowLeft':
      case 'ArrowRight': {
        if (ed.mode === 'edit') return
        const dir = e.key.slice(5).toLowerCase() as Dir
        if (ed.canPoint()) {
          e.preventDefault()
          this.pointMove(dir, e.shiftKey)
          return
        }
        e.preventDefault()
        this.commitEdit(dir)
      }
    }
  }

  private onGridKey (e: KeyboardEvent): void {
    // The editor can be open without focus (after clicking a header, a swatch…).
    // Keep its text rather than letting this key start a new edit over it.
    if (this.editor.active && !['Shift', 'Control', 'Meta', 'Alt'].includes(e.key)) {
      if (e.key === 'Escape') {
        e.preventDefault()
        this.cancelEdit()
        return
      }
      this.commitEdit('none')
    }
    const mod = e.ctrlKey || e.metaKey
    const key = e.key
    const sel = this.grid.sel
    const cmd = (c: string): void => {
      e.preventDefault()
      this.run(c)
    }
    if (mod && !e.altKey) {
      switch (key.toLowerCase()) {
        case 'z': return cmd(e.shiftKey ? 'edit.redo' : 'edit.undo')
        case 'y': return cmd('edit.redo')
        case 'b': return cmd('format.bold')
        case 'i': return cmd('format.italic')
        case 'u': return cmd('format.underline')
        case '5': return cmd('format.strike')
        case 'a': return cmd('edit.selectAll')
        case 'd': return cmd('edit.fillDown')
        case 'r': return cmd('edit.fillRight')
        case 'f': case 'h': return cmd('edit.find')
        case 'g': return cmd('view.goto')
        case 's': return cmd('file.save')
        case 'o': return cmd('file.open')
        case '`': return cmd('view.formulas')
        case '\\': return cmd('format.clear')
        case '/': return cmd('help.shortcuts')
        case 'v': if (e.shiftKey) return cmd('edit.pasteValues'); return // native paste event
        case 'c': case 'x': return // native copy/cut events
        case ' ':
          e.preventDefault()
          this.grid.setSelection({ ar: 0, ac: sel.ac, fr: MAX_ROWS - 1, fc: sel.fc }, false)
          return
        case 'home':
          e.preventDefault()
          this.grid.setSelection({ ar: 0, ac: 0, fr: 0, fc: 0 })
          return
        case 'end': {
          e.preventDefault()
          const r = Math.max(0, this.sheet.maxRow)
          const c = Math.max(0, this.sheet.maxCol)
          this.grid.setSelection({ ar: r, ac: c, fr: r, fc: c })
          return
        }
        case 'pageup':
        case 'pagedown': {
          e.preventDefault()
          const i = this.wb.sheets.indexOf(this.sheet) + (key === 'PageUp' ? -1 : 1)
          const s = this.wb.sheets[i]
          if (s) this.switchSheet(s)
          return
        }
      }
    }
    if (e.altKey && (key === '=' || e.code === 'Equal')) return cmd('insert.sum')
    if (key === 'ContextMenu' || (key === 'F10' && e.shiftKey)) {
      e.preventDefault()
      const rect = this.grid.cellRect(sel.ar, sel.ac)
      const host = this.grid.el.getBoundingClientRect()
      this.showMenu('#context-menu', host.left + rect.x + rect.w / 2, host.top + rect.y + rect.h)
      return
    }
    if (key === 'F11' && e.shiftKey) return cmd('insert.sheet')

    switch (key) {
      case 'ArrowUp':
      case 'ArrowDown':
      case 'ArrowLeft':
      case 'ArrowRight': {
        e.preventDefault()
        const dir = key.slice(5).toLowerCase() as Dir
        const [dr, dc] = dirDelta(dir)
        if (e.shiftKey) {
          const [fr, fc] = mod ? this.jump(sel.fr, sel.fc, dir) : [sel.fr + dr, sel.fc + dc]
          this.grid.setSelection({ ...sel, fr, fc })
        } else {
          const [r, c] = mod ? this.jump(sel.ar, sel.ac, dir) : [sel.ar + dr, sel.ac + dc]
          this.grid.setSelection({ ar: r, ac: c, fr: r, fc: c })
        }
        return
      }
      case 'Enter':
        e.preventDefault()
        this.moveActive(e.shiftKey ? 'up' : 'down')
        return
      case 'Tab':
        e.preventDefault()
        this.moveActive(e.shiftKey ? 'left' : 'right')
        return
      case 'Home':
        e.preventDefault()
        this.grid.setSelection({ ar: sel.ar, ac: 0, fr: sel.ar, fc: 0 })
        return
      case 'End': {
        e.preventDefault()
        const c = Math.max(0, this.sheet.maxCol)
        this.grid.setSelection({ ar: sel.ar, ac: c, fr: sel.ar, fc: c })
        return
      }
      case 'PageDown':
      case 'PageUp': {
        e.preventDefault()
        const n = this.grid.visibleRowCount() * (key === 'PageDown' ? 1 : -1)
        const r = clamp(sel.ar + n, 0, MAX_ROWS - 1)
        this.grid.setSelection({ ar: r, ac: sel.ac, fr: r, fc: sel.ac })
        return
      }
      case 'F2':
        e.preventDefault()
        this.startEdit(undefined, 'edit')
        return
      case 'Delete':
      case 'Backspace':
        e.preventDefault()
        this.clearSelection('contents')
        return
      case 'Escape':
        if (this.clip) {
          this.clip = null
          this.grid.setCopyRange(null)
        }
        return
      case ' ':
        if (e.shiftKey) {
          e.preventDefault()
          this.grid.setSelection({ ar: sel.ar, ac: 0, fr: sel.fr, fc: MAX_COLS - 1 }, false)
          return
        }
    }
    if (key.length === 1 && !mod && !e.altKey) {
      e.preventDefault()
      this.startEdit(key, 'enter')
    }
  }
}

// =============================================================================
// Helpers

function dirDelta (dir: Dir): [number, number] {
  return dir === 'up' ? [-1, 0] : dir === 'down' ? [1, 0] : dir === 'left' ? [0, -1] : [0, 1]
}

function clamp (v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v
}

function * span (a: number, b: number): Generator<number> {
  for (let i = a; i <= b; i++) yield i
}

function escapeRe (s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

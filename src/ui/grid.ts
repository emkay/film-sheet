import { colToName, formatAddress, MAX_COLS, MAX_ROWS, normalizeRange, type RangeAddress } from '../engine/address.js'
import { isError } from '../engine/values.js'
import type { Sheet, Workbook } from '../engine/workbook.js'
import { Axis } from './axis.js'

export const DEFAULT_COL_WIDTH = 100
export const DEFAULT_ROW_HEIGHT = 24
const HEADER_HEIGHT = 24
/** Browsers cap element sizes (~17M px in Firefox); scroll space is scaled above this. */
const MAX_SCROLL_PX = 8_000_000
const FONT_SIZE = 13
const PAD = 4

/** anchor (ar, ac) is the active cell; focus (fr, fc) is the opposite corner. */
export interface Selection {
  ar: number
  ac: number
  fr: number
  fc: number
}

export interface RangeHighlight extends RangeAddress {
  color: string
}

type Hit =
  | { area: 'corner' }
  | { area: 'col'; col: number; resize?: number }
  | { area: 'row'; row: number; resize?: number }
  | { area: 'cell'; row: number; col: number; fill?: boolean }

type DragMode =
  | { kind: 'cells' }
  | { kind: 'cols' }
  | { kind: 'rows' }
  | { kind: 'resizeCol'; col: number; start: number; startSize: number }
  | { kind: 'resizeRow'; row: number; start: number; startSize: number }
  | { kind: 'fill'; src: RangeAddress }
  | { kind: 'pick'; r: number; c: number }

interface Colors {
  bg: string
  text: string
  muted: string
  grid: string
  headerBg: string
  headerText: string
  headerActive: string
  accent: string
  error: string
}

export interface GridCallbacks {
  /** Selection moved (by mouse); also fired for keyboard via setSelection. */
  selectionChange?: () => void
  /** Return true to take over a cell press (e.g. formula reference picking). */
  cellPointerDown?: (row: number, col: number, e: PointerEvent) => boolean
  /** Formula reference pick in progress. */
  pick?: (range: RangeAddress, done: boolean) => void
  edit?: (row: number, col: number) => void
  fill?: (src: RangeAddress, dest: RangeAddress) => void
  resized?: (axis: 'rows' | 'cols', index: number, before: number, after: number) => void
  contextMenu?: (x: number, y: number, area: 'cell' | 'row' | 'col') => void
  /** A press anywhere other than a cell (headers, corner, fill handle). */
  otherPointerDown?: () => void
  /** The visible area moved (scroll or resize). */
  viewportChange?: () => void
}

export class GridView {
  readonly el: HTMLElement
  readonly scroller: HTMLElement
  private readonly canvas: HTMLCanvasElement
  private readonly ctx: CanvasRenderingContext2D
  private readonly spacer: HTMLElement
  readonly rows: Axis
  readonly cols: Axis
  sel: Selection = { ar: 0, ac: 0, fr: 0, fc: 0 }
  highlights: RangeHighlight[] = []
  copyRange: RangeAddress | null = null
  /** Hide the active-cell text while the in-cell editor covers it. */
  editing: { row: number; col: number } | null = null
  /** Draw formulas instead of their results (Ctrl+`). */
  showFormulas = false
  /** No fill handle when the sheet can't be edited. */
  readonly = false
  callbacks: GridCallbacks = {}

  private scrollX = 0
  private scrollY = 0
  private scaleX = 1
  private scaleY = 1
  private width = 0
  private height = 0
  private dpr = 1
  private headerW = 46
  private frame = 0
  private drag: DragMode | null = null
  private fillTarget: RangeAddress | null = null
  private lastPointer = { x: 0, y: 0 }
  private autoScrollRaf = 0
  private colors!: Colors
  private fontFamily = 'system-ui, sans-serif'
  private readonly widthCache = new Map<string, number>()
  private requested: { x: number; y: number } | null = null
  private readonly resizeObserver = new ResizeObserver(() => this.resize())
  /** Screen x of the first visible column, cached per frame for overflow checks. */
  private firstColX = 0
  private dashOffset = 0
  private antsTimer = 0
  private readonly slots: A11ySlot[]
  private activeSlot = 0
  private focused = false

  constructor (
    private readonly wb: Workbook,
    public sheet: Sheet
  ) {
    this.rows = new Axis(MAX_ROWS, DEFAULT_ROW_HEIGHT, sheet.rowHeights)
    this.cols = new Axis(MAX_COLS, DEFAULT_COL_WIDTH, sheet.colWidths)

    this.el = document.createElement('div')
    this.el.className = 'grid'
    this.canvas = document.createElement('canvas')
    this.canvas.className = 'grid-canvas'
    this.canvas.setAttribute('aria-hidden', 'true')
    this.scroller = document.createElement('div')
    this.scroller.className = 'grid-scroller'
    // The scroller only provides scrollbars and pointer input. Keyboard focus
    // and the accessibility tree live on the active-cell proxy (see syncA11y).
    this.scroller.tabIndex = -1
    this.scroller.setAttribute('aria-hidden', 'true')
    this.scroller.addEventListener('focus', () => this.focus())
    this.spacer = document.createElement('div')
    this.spacer.className = 'grid-spacer'
    this.scroller.append(this.spacer)
    this.el.setAttribute('role', 'grid')
    this.el.setAttribute('aria-rowcount', String(MAX_ROWS + 1))
    this.el.setAttribute('aria-colcount', String(MAX_COLS + 1))
    this.el.setAttribute('aria-multiselectable', 'true')
    this.el.setAttribute('aria-describedby', 'grid-help')
    this.el.setAttribute('aria-label', sheet.name)
    this.slots = [this.buildSlot(0), this.buildSlot(1)]
    this.el.append(this.canvas, this.scroller, ...this.slots.map((x) => x.root))
    this.el.addEventListener('focusin', () => this.setFocused(true))
    this.el.addEventListener('focusout', () => this.setFocused(false))
    this.ctx = this.canvas.getContext('2d', { alpha: false })!

    this.scroller.addEventListener('scroll', this.onScroll, { passive: true })
    this.scroller.addEventListener('pointerdown', this.onPointerDown)
    this.scroller.addEventListener('pointermove', this.onPointerMove)
    this.scroller.addEventListener('pointerup', this.onPointerUp)
    this.scroller.addEventListener('pointercancel', this.onPointerUp)
    this.scroller.addEventListener('dblclick', this.onDblClick)
    this.scroller.addEventListener('contextmenu', this.onContextMenu)
    this.resizeObserver.observe(this.el)
    this.refreshTheme()
  }

  // ---------------------------------------------------------------------------
  // Public API

  /** Stop observers and timers and remove the grid from the page. */
  dispose (): void {
    this.resizeObserver.disconnect()
    clearInterval(this.antsTimer)
    cancelAnimationFrame(this.frame)
    cancelAnimationFrame(this.autoScrollRaf)
    this.el.remove()
  }

  setSheet (sheet: Sheet): void {
    this.sheet = sheet
    this.el.setAttribute('aria-label', sheet.name)
    this.rows.use(sheet.rowHeights)
    this.cols.use(sheet.colWidths)
    this.setCopyRange(null)
    this.updateSpacer()
    this.invalidate()
  }

  /** Re-read sizes after structural edits. */
  sizesChanged (): void {
    this.rows.invalidate()
    this.cols.invalidate()
    this.updateSpacer()
    this.invalidate()
  }

  get range (): RangeAddress {
    return normalizeRange({ r1: this.sel.ar, c1: this.sel.ac, r2: this.sel.fr, c2: this.sel.fc })
  }

  setSelection (sel: Selection, scroll = true): void {
    this.sel = {
      ar: clamp(sel.ar, 0, MAX_ROWS - 1),
      ac: clamp(sel.ac, 0, MAX_COLS - 1),
      fr: clamp(sel.fr, 0, MAX_ROWS - 1),
      fc: clamp(sel.fc, 0, MAX_COLS - 1)
    }
    if (scroll) this.ensureVisible(this.sel.fr, this.sel.fc)
    this.invalidate()
    this.syncA11y(true)
    this.callbacks.selectionChange?.()
  }

  // ---------------------------------------------------------------------------
  // Accessibility: the canvas is invisible to assistive tech, so the active
  // cell is mirrored into a tiny DOM grid fragment (header row + one row with
  // its row header and the cell). Two slots alternate so every move is a real
  // focus change that screen readers announce, and screen magnifiers follow
  // the focused element, which is laid over the cell on screen.

  private buildSlot (i: number): A11ySlot {
    const root = document.createElement('div')
    root.className = 'a11y-slot'
    root.hidden = true
    const headRow = document.createElement('div')
    headRow.setAttribute('role', 'row')
    headRow.setAttribute('aria-rowindex', '1')
    const colHeader = document.createElement('div')
    colHeader.setAttribute('role', 'columnheader')
    colHeader.className = 'sr-only'
    headRow.append(colHeader)
    const headGroup = document.createElement('div')
    headGroup.setAttribute('role', 'rowgroup')
    headGroup.append(headRow)
    const row = document.createElement('div')
    row.setAttribute('role', 'row')
    const rowHeader = document.createElement('div')
    rowHeader.setAttribute('role', 'rowheader')
    rowHeader.setAttribute('aria-colindex', '1')
    rowHeader.className = 'sr-only'
    const cell = document.createElement('div')
    cell.setAttribute('role', 'gridcell')
    cell.className = 'cell-proxy'
    cell.tabIndex = -1
    const desc = document.createElement('div')
    desc.id = `cell-desc-${i}`
    desc.hidden = true
    cell.setAttribute('aria-describedby', desc.id)
    row.append(rowHeader, cell)
    const bodyGroup = document.createElement('div')
    bodyGroup.setAttribute('role', 'rowgroup')
    bodyGroup.append(row)
    root.append(headGroup, bodyGroup, desc)
    return { root, colHeader, row, rowHeader, cell, desc }
  }

  /** Does keyboard focus belong to the grid (rather than the editor or chrome)? */
  hasFocus (): boolean {
    // Inside a shadow root, document.activeElement is just the host.
    const a = (this.el.getRootNode() as Document | ShadowRoot).activeElement
    return a === this.scroller || (a instanceof HTMLElement && a.classList.contains('cell-proxy') && this.el.contains(a))
  }

  focus (): void {
    this.slots[this.activeSlot].root.hidden = false
    this.slots[this.activeSlot].cell.focus({ preventScroll: true })
  }

  /**
   * Refresh the active-cell proxy. With `move`, write it into the other slot
   * and move focus there (if the grid has focus) so the change is announced;
   * otherwise update the current slot in place (e.g. its value changed).
   */
  syncA11y (move = false): void {
    const hadFocus = this.hasFocus()
    const next = move ? 1 - this.activeSlot : this.activeSlot
    const slot = this.slots[next]
    const { ar, ac } = this.sel
    const r = this.range
    const cell = this.sheet.get(ar, ac)
    const address = formatAddress(ar, ac)
    const display = cell && cell.raw !== '' ? this.wb.getDisplay(cell) : ''
    slot.colHeader.setAttribute('aria-colindex', String(ac + 2))
    slot.colHeader.textContent = colToName(ac)
    slot.row.setAttribute('aria-rowindex', String(ar + 2))
    slot.rowHeader.textContent = String(ar + 1)
    slot.cell.setAttribute('aria-colindex', String(ac + 2))
    slot.cell.setAttribute('aria-selected', 'true')
    slot.cell.textContent = `${address}, ${display || 'blank'}`
    const notes: string[] = []
    if (cell?.raw.startsWith('=')) notes.push(`Formula ${cell.raw}`)
    if (cell?.error) notes.push(cell.error)
    const st = cell?.style
    if (st) {
      const styles = [st.bold && 'bold', st.italic && 'italic', st.underline && 'underlined', st.strike && 'strikethrough'].filter(Boolean)
      if (styles.length) notes.push(styles.join(', '))
    }
    if (r.r1 !== r.r2 || r.c1 !== r.c2) notes.push('In selection ' + (r.r1 === r.r2 && r.c1 === r.c2 ? address : `${formatAddress(r.r1, r.c1)} to ${formatAddress(r.r2, r.c2)}`))
    slot.desc.textContent = notes.join('. ')
    // Lay the proxy over the cell so magnifiers and focus rings land there.
    const rect = this.cellRect(ar, ac)
    Object.assign(slot.cell.style, { left: rect.x + 'px', top: rect.y + 'px', width: rect.w + 'px', height: rect.h + 'px' })
    if (!move) return
    slot.root.hidden = false
    const prev = this.slots[this.activeSlot]
    this.activeSlot = next
    if (hadFocus) slot.cell.focus({ preventScroll: true })
    prev.root.hidden = true
  }

  private setFocused (on: boolean): void {
    // focusout fires before focusin when focus moves between the two slots.
    queueMicrotask(() => {
      const now = this.hasFocus()
      if (now !== this.focused || on !== now) {
        this.focused = now
        this.invalidate()
      }
    })
  }

  /** Viewport in cells, for PageUp/PageDown. */
  visibleRowCount (): number {
    return Math.max(1, Math.floor((this.height - HEADER_HEIGHT) / DEFAULT_ROW_HEIGHT) - 1)
  }

  ensureVisible (row: number, col: number): void {
    const viewW = this.width - this.headerW
    const viewH = this.height - HEADER_HEIGHT
    const x = this.cols.offset(col)
    const w = this.cols.size(col)
    const y = this.rows.offset(row)
    const h = this.rows.size(row)
    let sx = this.scrollX
    let sy = this.scrollY
    // Whole-row / whole-column selections shouldn't yank the other axis.
    if (!(this.sel.ac === 0 && this.sel.fc === MAX_COLS - 1)) {
      if (x < sx) sx = x
      else if (x + w > sx + viewW) sx = Math.min(x, x + w - viewW)
    }
    if (!(this.sel.ar === 0 && this.sel.fr === MAX_ROWS - 1)) {
      if (y < sy) sy = y
      else if (y + h > sy + viewH) sy = Math.min(y, y + h - viewH)
    }
    if (sx !== this.scrollX || sy !== this.scrollY) this.scrollTo(sx, sy)
  }

  scrollTo (x: number, y: number): void {
    this.scrollX = Math.max(0, x)
    this.scrollY = Math.max(0, y)
    this.requested = { x: this.scrollX, y: this.scrollY }
    this.scroller.scrollLeft = this.scrollX / this.scaleX
    this.scroller.scrollTop = this.scrollY / this.scaleY
    this.updateHeaderWidth()
    this.invalidate()
  }

  /** Row-header width fits the largest visible row number. */
  private updateHeaderWidth (): void {
    const last = this.rows.indexAt(this.scrollY + this.height)
    const digits = String(last + 1).length
    // ~7px per digit at the header font; measured precisely in draw().
    const w = Math.max(46, digits * 7 + 22)
    if (w > this.headerW) this.headerW = w
  }

  /** Cell rectangle in coordinates relative to the grid element. */
  cellRect (row: number, col: number): { x: number; y: number; w: number; h: number } {
    return {
      x: this.headerW + this.cols.offset(col) - this.scrollX,
      y: HEADER_HEIGHT + this.rows.offset(row) - this.scrollY,
      w: this.cols.size(col),
      h: this.rows.size(row)
    }
  }

  get viewportWidth (): number {
    return this.width
  }

  invalidate (): void {
    if (this.frame) return
    this.frame = requestAnimationFrame(() => {
      this.frame = 0
      this.draw()
    })
  }

  /** Resolve Film's theme tokens to concrete colours the canvas can use. */
  refreshTheme (): void {
    // Resolve tokens where the grid lives, so themes set on any ancestor apply.
    const probe = document.createElement('span')
    probe.style.display = 'none'
    this.el.append(probe)
    const resolve = (expr: string, fallback: string): string => {
      probe.style.color = ''
      probe.style.color = expr
      return getComputedStyle(probe).color || fallback
    }
    const forced = matchMedia('(forced-colors: active)').matches
    const moreContrast = matchMedia('(prefers-contrast: more)').matches
    this.colors = forced
      ? {
          // Windows High Contrast: use the user's system colours throughout.
          bg: resolve('Canvas', '#fff'),
          text: resolve('CanvasText', '#000'),
          muted: resolve('GrayText', '#666'),
          grid: resolve('GrayText', '#666'),
          headerBg: resolve('Canvas', '#fff'),
          headerText: resolve('CanvasText', '#000'),
          headerActive: resolve('Canvas', '#fff'),
          accent: resolve('Highlight', '#00f'),
          error: resolve('CanvasText', '#000')
        }
      : {
          bg: resolve('var(--film-color-surface)', '#fff'),
          text: resolve('var(--film-color-text)', '#111'),
          muted: resolve('var(--film-color-text-muted)', '#666'),
          grid: moreContrast
            ? resolve('var(--film-color-control-border)', '#888')
            : resolve('color-mix(in oklab, var(--film-color-border) 55%, var(--film-color-surface))', '#ddd'),
          headerBg: resolve('var(--film-color-background)', '#f4f4f4'),
          headerText: resolve('var(--film-color-text-muted)', '#555'),
          headerActive: resolve('color-mix(in oklab, var(--film-color-focus) 14%, var(--film-color-background))', '#dde'),
          accent: resolve('var(--film-color-focus)', '#1a73e8'),
          // ≥4.5:1 on surfaces and the selection tint in both schemes.
          error: resolve('light-dark(oklch(0.48 0.19 25), oklch(0.8 0.13 25))', '#b00')
        }
    const family = getComputedStyle(this.el).getPropertyValue('--film-font-sans').trim()
    if (family) this.fontFamily = family
    probe.remove()
    this.widthCache.clear()
    this.invalidate()
  }

  /** Animate the dashed "copied" border, like Excel's marching ants. */
  setCopyRange (range: RangeAddress | null): void {
    this.copyRange = range
    clearInterval(this.antsTimer)
    if (range && !matchMedia('(prefers-reduced-motion: reduce)').matches) {
      this.antsTimer = window.setInterval(() => {
        this.dashOffset = (this.dashOffset + 1) % 8
        this.invalidate()
      }, 80)
    }
    this.invalidate()
  }

  /** Width needed to show every value in a column (double-click autofit). */
  measureColumn (col: number): number {
    const map = this.sheet.cols.get(col)
    let max = 0
    if (map) {
      let n = 0
      for (const cell of map.values()) {
        if (++n > 20000) break
        const text = this.wb.getDisplay(cell)
        if (!text) continue
        this.ctx.font = this.fontFor(cell.style?.bold, cell.style?.italic)
        max = Math.max(max, this.measure(text, this.ctx.font))
      }
    }
    return max ? Math.ceil(max + PAD * 2 + 4) : DEFAULT_COL_WIDTH
  }

  // ---------------------------------------------------------------------------
  // Layout & scrolling

  private resize (): void {
    const rect = this.el.getBoundingClientRect()
    this.width = this.scroller.clientWidth || rect.width
    this.height = this.scroller.clientHeight || rect.height
    this.dpr = window.devicePixelRatio || 1
    this.canvas.width = Math.round(this.width * this.dpr)
    this.canvas.height = Math.round(this.height * this.dpr)
    this.canvas.style.width = this.width + 'px'
    this.canvas.style.height = this.height + 'px'
    this.updateSpacer()
    // The scale factors may have changed: re-derive the logical scroll position.
    this.scrollX = this.scroller.scrollLeft * this.scaleX
    this.scrollY = this.scroller.scrollTop * this.scaleY
    this.draw()
    this.callbacks.viewportChange?.()
  }

  private updateSpacer (): void {
    const totalW = this.headerW + this.cols.total() + 40
    const totalH = HEADER_HEIGHT + this.rows.total() + 40
    const spW = Math.min(totalW, MAX_SCROLL_PX)
    const spH = Math.min(totalH, MAX_SCROLL_PX)
    this.spacer.style.width = spW + 'px'
    this.spacer.style.height = spH + 'px'
    this.scaleX = spW < totalW ? (totalW - this.width) / Math.max(1, spW - this.width) : 1
    this.scaleY = spH < totalH ? (totalH - this.height) / Math.max(1, spH - this.height) : 1
  }

  private readonly onScroll = (): void => {
    const x = this.scroller.scrollLeft * this.scaleX
    const y = this.scroller.scrollTop * this.scaleY
    // With scaled scrolling the browser rounds scrollTop, so a programmatic
    // scroll lands up to one scale step off. Keep the exact position we asked for.
    const r = this.requested
    this.scrollX = r && Math.abs(x - r.x) <= this.scaleX ? r.x : x
    this.scrollY = r && Math.abs(y - r.y) <= this.scaleY ? r.y : y
    this.requested = null
    // Draw now (keeps scrolling smooth) and drop any frame already queued.
    if (this.frame) {
      cancelAnimationFrame(this.frame)
      this.frame = 0
    }
    this.draw()
    this.syncA11y()
    this.callbacks.viewportChange?.()
  }

  get headerWidth (): number {
    return this.headerW
  }

  // ---------------------------------------------------------------------------
  // Hit testing & pointer interaction

  private hit (x: number, y: number): Hit {
    if (x < this.headerW && y < HEADER_HEIGHT) return { area: 'corner' }
    if (y < HEADER_HEIGHT) {
      const cx = x - this.headerW + this.scrollX
      const col = this.cols.indexAt(cx)
      const left = this.cols.offset(col)
      const right = left + this.cols.size(col)
      if (right - cx <= 4) return { area: 'col', col, resize: col }
      if (cx - left <= 3 && col > 0) return { area: 'col', col, resize: col - 1 }
      return { area: 'col', col }
    }
    if (x < this.headerW) {
      const cy = y - HEADER_HEIGHT + this.scrollY
      const row = this.rows.indexAt(cy)
      const top = this.rows.offset(row)
      const bottom = top + this.rows.size(row)
      if (bottom - cy <= 3) return { area: 'row', row, resize: row }
      if (cy - top <= 2 && row > 0) return { area: 'row', row, resize: row - 1 }
      return { area: 'row', row }
    }
    const row = this.rows.indexAt(y - HEADER_HEIGHT + this.scrollY)
    const col = this.cols.indexAt(x - this.headerW + this.scrollX)
    const r = this.range
    const corner = this.cellRect(r.r2, r.c2)
    const fx = corner.x + corner.w
    const fy = corner.y + corner.h
    const fill = !this.editing && !this.readonly && Math.abs(x - fx) <= 5 && Math.abs(y - fy) <= 5
    return { area: 'cell', row, col, fill }
  }

  private local (e: PointerEvent | MouseEvent): { x: number; y: number } {
    const rect = this.scroller.getBoundingClientRect()
    return { x: e.clientX - rect.left, y: e.clientY - rect.top }
  }

  private readonly onPointerDown = (e: PointerEvent): void => {
    const { x, y } = this.local(e)
    if (x > this.width || y > this.height) return // native scrollbar
    if (e.button === 2) {
      // Right-click keeps the selection if it was inside it.
      const h = this.hit(x, y)
      if (h.area === 'cell') {
        const r = this.range
        if (h.row < r.r1 || h.row > r.r2 || h.col < r.c1 || h.col > r.c2) this.setSelection({ ar: h.row, ac: h.col, fr: h.row, fc: h.col }, false)
      }
      return
    }
    if (e.button !== 0) return
    const h = this.hit(x, y)
    if (e.pointerType === 'touch') {
      // Touch pans natively (touch-action), so a tap only selects.
      if (h.area !== 'cell') return
      if (this.callbacks.cellPointerDown?.(h.row, h.col, e)) {
        e.preventDefault()
        this.callbacks.pick?.({ r1: h.row, c1: h.col, r2: h.row, c2: h.col }, true)
      } else {
        this.focus()
        this.setSelection({ ar: h.row, ac: h.col, fr: h.row, fc: h.col }, false)
      }
      return
    }
    this.lastPointer = { x, y }
    if (h.area === 'cell' && !h.fill && this.callbacks.cellPointerDown?.(h.row, h.col, e)) {
      // Formula reference picking: keep focus in the editor.
      e.preventDefault()
      this.drag = { kind: 'pick', r: h.row, c: h.col }
      this.scroller.setPointerCapture(e.pointerId)
      this.callbacks.pick?.({ r1: h.row, c1: h.col, r2: h.row, c2: h.col }, false)
      return
    }
    if (h.area !== 'cell' || h.fill) this.callbacks.otherPointerDown?.()
    this.focus()
    this.scroller.setPointerCapture(e.pointerId)
    switch (h.area) {
      case 'corner':
        this.setSelection({ ar: 0, ac: 0, fr: MAX_ROWS - 1, fc: MAX_COLS - 1 }, false)
        return
      case 'col':
        if (h.resize !== undefined) {
          this.drag = { kind: 'resizeCol', col: h.resize, start: x, startSize: this.cols.size(h.resize) }
          return
        }
        this.drag = { kind: 'cols' }
        if (e.shiftKey) this.setSelection({ ...this.sel, ar: 0, fr: MAX_ROWS - 1, fc: h.col }, false)
        else this.setSelection({ ar: 0, ac: h.col, fr: MAX_ROWS - 1, fc: h.col }, false)
        return
      case 'row':
        if (h.resize !== undefined) {
          this.drag = { kind: 'resizeRow', row: h.resize, start: y, startSize: this.rows.size(h.resize) }
          return
        }
        this.drag = { kind: 'rows' }
        if (e.shiftKey) this.setSelection({ ...this.sel, ac: 0, fc: MAX_COLS - 1, fr: h.row }, false)
        else this.setSelection({ ar: h.row, ac: 0, fr: h.row, fc: MAX_COLS - 1 }, false)
        return
      case 'cell':
        if (h.fill) {
          this.drag = { kind: 'fill', src: this.range }
          this.fillTarget = null
          return
        }
        this.drag = { kind: 'cells' }
        if (e.shiftKey) this.setSelection({ ...this.sel, fr: h.row, fc: h.col }, false)
        else this.setSelection({ ar: h.row, ac: h.col, fr: h.row, fc: h.col }, false)
    }
  }

  private readonly onPointerMove = (e: PointerEvent): void => {
    const { x, y } = this.local(e)
    this.lastPointer = { x, y }
    if (!this.drag) {
      this.updateCursor(x, y)
      return
    }
    this.dragTo(x, y)
    this.startAutoScroll()
  }

  private dragTo (x: number, y: number): void {
    const d = this.drag
    if (!d) return
    const row = this.rows.indexAt(Math.max(0, y - HEADER_HEIGHT) + this.scrollY)
    const col = this.cols.indexAt(Math.max(0, x - this.headerW) + this.scrollX)
    switch (d.kind) {
      case 'cells':
        if (row !== this.sel.fr || col !== this.sel.fc) this.setSelection({ ...this.sel, fr: row, fc: col }, false)
        break
      case 'cols':
        if (col !== this.sel.fc) this.setSelection({ ...this.sel, fc: col }, false)
        break
      case 'rows':
        if (row !== this.sel.fr) this.setSelection({ ...this.sel, fr: row }, false)
        break
      case 'resizeCol': {
        const size = Math.max(16, Math.round(d.startSize + x - d.start))
        this.cols.setSize(d.col, size)
        this.updateSpacer()
        this.invalidate()
        break
      }
      case 'resizeRow': {
        const size = Math.max(12, Math.round(d.startSize + y - d.start))
        this.rows.setSize(d.row, size)
        this.updateSpacer()
        this.invalidate()
        break
      }
      case 'fill': {
        const s = d.src
        // Extend along whichever axis the pointer has moved further outside the source.
        const dy = row > s.r2 ? row - s.r2 : row < s.r1 ? row - s.r1 : 0
        const dx = col > s.c2 ? col - s.c2 : col < s.c1 ? col - s.c1 : 0
        let t: RangeAddress | null = null
        if (Math.abs(dy) >= Math.abs(dx) && dy !== 0) t = dy > 0 ? { ...s, r2: row } : { ...s, r1: row }
        else if (dx !== 0) t = dx > 0 ? { ...s, c2: col } : { ...s, c1: col }
        this.fillTarget = t
        this.invalidate()
        break
      }
      case 'pick':
        this.callbacks.pick?.(normalizeRange({ r1: d.r, c1: d.c, r2: row, c2: col }), false)
        break
    }
  }

  private startAutoScroll (): void {
    if (this.autoScrollRaf) return
    const step = (): void => {
      this.autoScrollRaf = 0
      if (!this.drag || this.drag.kind === 'resizeCol' || this.drag.kind === 'resizeRow') return
      const { x, y } = this.lastPointer
      let vx = 0
      let vy = 0
      if (x > this.width - 10) vx = x - this.width + 10
      else if (x < this.headerW && this.drag.kind !== 'rows') vx = x - this.headerW
      if (y > this.height - 10) vy = y - this.height + 10
      else if (y < HEADER_HEIGHT && this.drag.kind !== 'cols') vy = y - HEADER_HEIGHT
      if (this.drag.kind === 'cols') vy = 0
      if (this.drag.kind === 'rows') vx = 0
      if (vx || vy) {
        this.scrollTo(this.scrollX + vx * 0.6, this.scrollY + vy * 0.6)
        this.dragTo(x, y)
        this.autoScrollRaf = requestAnimationFrame(step)
      }
    }
    this.autoScrollRaf = requestAnimationFrame(step)
  }

  private readonly onPointerUp = (e: PointerEvent): void => {
    const d = this.drag
    this.drag = null
    if (this.scroller.hasPointerCapture(e.pointerId)) this.scroller.releasePointerCapture(e.pointerId)
    if (!d) return
    if (d.kind === 'fill' && this.fillTarget) {
      const t = this.fillTarget
      this.fillTarget = null
      this.callbacks.fill?.(d.src, t)
      this.invalidate()
    } else if (d.kind === 'fill') {
      this.invalidate()
    } else if (d.kind === 'resizeCol') {
      this.callbacks.resized?.('cols', d.col, d.startSize, this.cols.size(d.col))
    } else if (d.kind === 'resizeRow') {
      this.callbacks.resized?.('rows', d.row, d.startSize, this.rows.size(d.row))
    } else if (d.kind === 'pick') {
      const { x, y } = this.local(e)
      const row = this.rows.indexAt(Math.max(0, y - HEADER_HEIGHT) + this.scrollY)
      const col = this.cols.indexAt(Math.max(0, x - this.headerW) + this.scrollX)
      this.callbacks.pick?.(normalizeRange({ r1: d.r, c1: d.c, r2: row, c2: col }), true)
    }
  }

  private readonly onDblClick = (e: MouseEvent): void => {
    const { x, y } = this.local(e)
    const h = this.hit(x, y)
    if (h.area === 'cell' && !h.fill) this.callbacks.edit?.(h.row, h.col)
    else if (h.area === 'col' && h.resize !== undefined) {
      const before = this.cols.size(h.resize)
      this.cols.setSize(h.resize, this.measureColumn(h.resize))
      this.updateSpacer()
      this.invalidate()
      this.callbacks.resized?.('cols', h.resize, before, this.cols.size(h.resize))
    } else if (h.area === 'row' && h.resize !== undefined) {
      const before = this.rows.size(h.resize)
      this.rows.setSize(h.resize, DEFAULT_ROW_HEIGHT)
      this.updateSpacer()
      this.invalidate()
      this.callbacks.resized?.('rows', h.resize, before, DEFAULT_ROW_HEIGHT)
    }
  }

  private readonly onContextMenu = (e: MouseEvent): void => {
    e.preventDefault()
    const { x, y } = this.local(e)
    const h = this.hit(x, y)
    const area = h.area === 'col' ? 'col' : h.area === 'row' ? 'row' : 'cell'
    this.callbacks.contextMenu?.(e.clientX, e.clientY, area)
  }

  private updateCursor (x: number, y: number): void {
    const h = this.hit(x, y)
    let cursor = 'cell'
    if (h.area === 'col') cursor = h.resize !== undefined ? 'col-resize' : 's-resize'
    else if (h.area === 'row') cursor = h.resize !== undefined ? 'row-resize' : 'e-resize'
    else if (h.area === 'corner') cursor = 'se-resize'
    else if (h.fill) cursor = 'crosshair'
    if (this.scroller.style.cursor !== cursor) this.scroller.style.cursor = cursor
  }

  // ---------------------------------------------------------------------------
  // Rendering

  private fontFor (bold?: boolean, italic?: boolean): string {
    return `${italic ? 'italic ' : ''}${bold ? '600 ' : ''}${FONT_SIZE}px ${this.fontFamily}`
  }

  private measure (text: string, font: string): number {
    const key = font + '\u0000' + text
    let w = this.widthCache.get(key)
    if (w === undefined) {
      if (this.widthCache.size > 20000) this.widthCache.clear()
      w = this.ctx.measureText(text).width
      this.widthCache.set(key, w)
    }
    return w
  }

  private draw (): void {
    if (!this.width || !this.height) return
    const ctx = this.ctx
    const c = this.colors
    const { width, height } = this
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0)
    ctx.fillStyle = c.bg
    ctx.fillRect(0, 0, width, height)

    // Visible window
    const r0 = this.rows.indexAt(this.scrollY)
    const c0 = this.cols.indexAt(this.scrollX)
    this.firstColX = this.headerW + this.cols.offset(c0) - this.scrollX
    ctx.font = `${FONT_SIZE - 1}px ${this.fontFamily}`
    const lastRowGuess = this.rows.indexAt(this.scrollY + height)
    this.headerW = Math.max(46, Math.ceil(ctx.measureText(String(lastRowGuess + 1)).width) + 22)
    const hw = this.headerW
    const rowYs: number[] = []
    const rowIdx: number[] = []
    for (let r = r0, y = HEADER_HEIGHT + this.rows.offset(r0) - this.scrollY; r < MAX_ROWS && y < height; r++) {
      rowIdx.push(r)
      rowYs.push(y)
      y += this.rows.size(r)
    }
    const colXs: number[] = []
    const colIdx: number[] = []
    for (let col = c0, x = hw + this.cols.offset(c0) - this.scrollX; col < MAX_COLS && x < width; col++) {
      colIdx.push(col)
      colXs.push(x)
      x += this.cols.size(col)
    }
    const nR = rowIdx.length
    const nC = colIdx.length
    const lastCol = colIdx[nC - 1]
    const lastColEnd = colXs[nC - 1] + this.cols.size(lastCol)
    const lastRowEnd = rowYs[nR - 1] + this.rows.size(rowIdx[nR - 1])

    ctx.save()
    ctx.beginPath()
    ctx.rect(hw, HEADER_HEIGHT, width - hw, height - HEADER_HEIGHT)
    ctx.clip()

    // 1. Cell fills
    const sheet = this.sheet
    for (let j = 0; j < nC; j++) {
      const col = sheet.cols.get(colIdx[j])
      if (!col) continue
      for (let i = 0; i < nR; i++) {
        const fill = col.get(rowIdx[i])?.style?.fill
        if (fill) {
          ctx.fillStyle = fill
          ctx.fillRect(colXs[j], rowYs[i], this.cols.size(colIdx[j]), this.rows.size(rowIdx[i]))
        }
      }
    }

    // 2. Grid lines (one path)
    ctx.beginPath()
    for (let i = 0; i < nR; i++) {
      const y = Math.round(rowYs[i] + this.rows.size(rowIdx[i])) - 0.5
      ctx.moveTo(hw, y)
      ctx.lineTo(lastColEnd, y)
    }
    for (let j = 0; j < nC; j++) {
      const x = Math.round(colXs[j] + this.cols.size(colIdx[j])) - 0.5
      ctx.moveTo(x, HEADER_HEIGHT)
      ctx.lineTo(x, lastRowEnd)
    }
    ctx.strokeStyle = c.grid
    ctx.lineWidth = 1
    ctx.stroke()

    // 3. Text
    ctx.textBaseline = 'middle'
    // Left-aligned text in cells to the left of the viewport may overflow into it.
    const lookBack = 12
    const firstCol = Math.max(0, c0 - lookBack)
    for (let i = 0; i < nR; i++) {
      const r = rowIdx[i]
      const y = rowYs[i]
      const h = this.rows.size(r)
      let x = colXs[0]
      for (let k = c0 - 1; k >= firstCol; k--) x -= this.cols.size(k)
      for (let col = firstCol; col <= lastCol; col++) {
        const w = this.cols.size(col)
        const cell = sheet.get(r, col)
        if (cell && cell.raw !== '' && !(this.editing && this.editing.row === r && this.editing.col === col)) {
          this.drawCellText(cell, r, col, x, y, w, h, lastColEnd, col < c0)
        }
        x += w
      }
    }
    ctx.restore()

    // 4. Highlights: formula references being edited
    ctx.save()
    ctx.beginPath()
    ctx.rect(hw, HEADER_HEIGHT, width - hw, height - HEADER_HEIGHT)
    ctx.clip()
    for (const hl of this.highlights) {
      const rect = this.rangeRect(hl)
      if (!rect) continue
      ctx.globalAlpha = 0.1
      ctx.fillStyle = hl.color
      ctx.fillRect(rect.x, rect.y, rect.w, rect.h)
      ctx.globalAlpha = 1
      ctx.strokeStyle = hl.color
      ctx.lineWidth = 2
      ctx.strokeRect(rect.x + 1, rect.y + 1, rect.w - 2, rect.h - 2)
    }

    // 5. Selection
    const sr = this.range
    const srect = this.rangeRect(sr)
    if (srect) {
      const multi = sr.r1 !== sr.r2 || sr.c1 !== sr.c2
      if (multi) {
        // Shade the range but leave the active cell clear (even-odd hole).
        const a = this.cellRect(this.sel.ar, this.sel.ac)
        ctx.beginPath()
        ctx.rect(srect.x, srect.y, srect.w, srect.h)
        ctx.rect(a.x, a.y, a.w, a.h)
        ctx.globalAlpha = 0.12
        ctx.fillStyle = c.accent
        ctx.fill('evenodd')
        ctx.globalAlpha = 1
      }
      // A grey selection says "the grid doesn't have keyboard focus".
      ctx.strokeStyle = this.focused || this.editing ? c.accent : c.muted
      ctx.lineWidth = 2
      ctx.strokeRect(srect.x + 0.5, srect.y + 0.5, srect.w - 1, srect.h - 1)
      if (multi) {
        const a = this.cellRect(this.sel.ar, this.sel.ac)
        ctx.lineWidth = 1
        ctx.strokeRect(a.x + 1.5, a.y + 1.5, a.w - 3, a.h - 3)
      }
      // fill handle
      if (!this.editing && !this.readonly) {
        const fx = srect.x + srect.w
        const fy = srect.y + srect.h
        ctx.fillStyle = c.bg
        ctx.fillRect(fx - 4.5, fy - 4.5, 9, 9)
        ctx.fillStyle = c.accent
        ctx.fillRect(fx - 3.5, fy - 3.5, 7, 7)
      }
    }
    if (this.fillTarget) {
      const rect = this.rangeRect(this.fillTarget)
      if (rect) {
        ctx.setLineDash([4, 3])
        ctx.strokeStyle = c.muted
        ctx.lineWidth = 1
        ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.w - 1, rect.h - 1)
        ctx.setLineDash([])
      }
    }
    if (this.copyRange) {
      const rect = this.rangeRect(this.copyRange)
      if (rect) {
        ctx.setLineDash([5, 3])
        ctx.lineDashOffset = -this.dashOffset
        ctx.strokeStyle = c.accent
        ctx.lineWidth = 2
        ctx.strokeRect(rect.x + 1, rect.y + 1, rect.w - 2, rect.h - 2)
        ctx.setLineDash([])
      }
    }
    ctx.restore()

    // 6. Headers
    this.drawHeaders(rowIdx, rowYs, colIdx, colXs, sr)
  }

  private drawCellText (
    cell: NonNullable<ReturnType<Sheet['get']>>,
    _r: number,
    col: number,
    x: number,
    y: number,
    w: number,
    h: number,
    viewRight: number,
    offscreen: boolean
  ): void {
    const asFormula = this.showFormulas && cell.raw[0] === '='
    const text = asFormula ? cell.raw : this.wb.getDisplay(cell)
    if (!text) return
    const ctx = this.ctx
    const style = cell.style
    const v = asFormula ? cell.raw : cell.value
    const isNum = typeof v === 'number'
    const align = style?.align ?? (isNum ? 'right' : typeof v === 'boolean' || isError(v) ? 'center' : 'left')
    if (offscreen && align !== 'left') return
    const font = this.fontFor(style?.bold, style?.italic)
    ctx.font = font
    let tw = this.measure(text, font)
    let shown = text
    let clipRight = x + w
    const avail = w - PAD * 2
    if (tw > avail) {
      if (isNum && !cell.style?.fmt?.includes('@')) {
        // Numbers never overflow: show ### like Excel
        const n = Math.max(1, Math.floor(avail / this.measure('#', font)))
        shown = '#'.repeat(n)
        tw = this.measure(shown, font)
      } else if (align === 'left') {
        // Overflow into empty neighbours
        let k = col + 1
        let edge = x + w
        while (edge - x - PAD < tw && edge < viewRight + 200 && k < MAX_COLS) {
          const next = this.sheet.get(_r, k)
          if (next && next.raw !== '') break
          edge += this.cols.size(k)
          k++
        }
        clipRight = edge
      }
    }
    if (offscreen && clipRight <= this.firstColX) return
    let tx: number
    if (align === 'right') tx = x + w - PAD - tw
    else if (align === 'center') tx = x + (w - tw) / 2
    else tx = x + PAD
    ctx.fillStyle = style?.color ?? (style?.fill ? readableOn(style.fill) : isError(v) ? this.colors.error : this.colors.text)
    const needsClip = tx < x || tx + tw > clipRight
    if (needsClip) {
      ctx.save()
      ctx.beginPath()
      ctx.rect(x, y, clipRight - x - 1, h)
      ctx.clip()
    }
    const ty = y + h / 2 + 0.5
    ctx.fillText(shown, tx, ty)
    if (style?.underline || style?.strike) {
      ctx.fillRect(tx, style.underline ? ty + FONT_SIZE / 2 - 1 : ty, tw, 1)
    }
    if (needsClip) ctx.restore()
  }

  private drawHeaders (rowIdx: number[], rowYs: number[], colIdx: number[], colXs: number[], sr: RangeAddress): void {
    const ctx = this.ctx
    const c = this.colors
    const hw = this.headerW
    const { width, height } = this
    ctx.fillStyle = c.headerBg
    ctx.fillRect(0, 0, width, HEADER_HEIGHT)
    ctx.fillRect(0, 0, hw, height)
    ctx.font = `${FONT_SIZE - 1}px ${this.fontFamily}`
    ctx.textBaseline = 'middle'
    ctx.textAlign = 'center'

    // Column headers
    ctx.save()
    ctx.beginPath()
    ctx.rect(hw, 0, width - hw, HEADER_HEIGHT)
    ctx.clip()
    for (let j = 0; j < colIdx.length; j++) {
      const col = colIdx[j]
      const x = colXs[j]
      const w = this.cols.size(col)
      const inSel = col >= sr.c1 && col <= sr.c2
      if (inSel) {
        ctx.fillStyle = c.headerActive
        ctx.fillRect(x, 0, w, HEADER_HEIGHT)
        ctx.fillStyle = c.accent
        ctx.fillRect(x, HEADER_HEIGHT - 2, w, 2)
      }
      ctx.fillStyle = inSel ? c.text : c.headerText
      ctx.fillText(colToName(col), x + w / 2, HEADER_HEIGHT / 2 + 0.5)
    }
    ctx.beginPath()
    for (let j = 0; j < colIdx.length; j++) {
      const x = Math.round(colXs[j] + this.cols.size(colIdx[j])) - 0.5
      ctx.moveTo(x, 0)
      ctx.lineTo(x, HEADER_HEIGHT)
    }
    ctx.strokeStyle = c.grid
    ctx.stroke()
    ctx.restore()

    // Row headers
    ctx.save()
    ctx.beginPath()
    ctx.rect(0, HEADER_HEIGHT, hw, height - HEADER_HEIGHT)
    ctx.clip()
    for (let i = 0; i < rowIdx.length; i++) {
      const r = rowIdx[i]
      const y = rowYs[i]
      const h = this.rows.size(r)
      const inSel = r >= sr.r1 && r <= sr.r2
      if (inSel) {
        ctx.fillStyle = c.headerActive
        ctx.fillRect(0, y, hw, h)
        ctx.fillStyle = c.accent
        ctx.fillRect(hw - 2, y, 2, h)
      }
      ctx.fillStyle = inSel ? c.text : c.headerText
      ctx.fillText(String(r + 1), hw / 2, y + h / 2 + 0.5)
    }
    ctx.beginPath()
    for (let i = 0; i < rowIdx.length; i++) {
      const y = Math.round(rowYs[i] + this.rows.size(rowIdx[i])) - 0.5
      ctx.moveTo(0, y)
      ctx.lineTo(hw, y)
    }
    ctx.strokeStyle = c.grid
    ctx.stroke()
    ctx.restore()
    ctx.textAlign = 'start'

    // Header borders + corner
    ctx.fillStyle = c.headerBg
    ctx.fillRect(0, 0, hw, HEADER_HEIGHT)
    ctx.fillStyle = c.muted
    ctx.beginPath()
    ctx.moveTo(hw - 4, HEADER_HEIGHT - 4)
    ctx.lineTo(hw - 4, 8)
    ctx.lineTo(hw - 16, HEADER_HEIGHT - 4)
    ctx.closePath()
    ctx.globalAlpha = 0.35
    ctx.fill()
    ctx.globalAlpha = 1
    ctx.strokeStyle = c.grid
    ctx.beginPath()
    ctx.moveTo(0, HEADER_HEIGHT - 0.5)
    ctx.lineTo(width, HEADER_HEIGHT - 0.5)
    ctx.moveTo(hw - 0.5, 0)
    ctx.lineTo(hw - 0.5, height)
    ctx.stroke()
  }

  /** Screen rectangle for a range, clipped loosely to the viewport. */
  private rangeRect (r: RangeAddress): { x: number; y: number; w: number; h: number } | null {
    const x1 = this.headerW + this.cols.offset(r.c1) - this.scrollX
    const y1 = HEADER_HEIGHT + this.rows.offset(r.r1) - this.scrollY
    const x2 = this.headerW + this.cols.offset(r.c2 + 1) - this.scrollX
    const y2 = HEADER_HEIGHT + this.rows.offset(r.r2 + 1) - this.scrollY
    if (x2 < this.headerW || y2 < HEADER_HEIGHT || x1 > this.width || y1 > this.height) return null
    // Clamp huge ranges so the canvas doesn't deal with million-pixel rects.
    const cx1 = Math.max(x1, this.headerW - 4)
    const cy1 = Math.max(y1, HEADER_HEIGHT - 4)
    const cx2 = Math.min(x2, this.width + 4)
    const cy2 = Math.min(y2, this.height + 4)
    return { x: cx1, y: cy1, w: cx2 - cx1, h: cy2 - cy1 }
  }
}

interface A11ySlot {
  root: HTMLElement
  colHeader: HTMLElement
  row: HTMLElement
  rowHeader: HTMLElement
  cell: HTMLElement
  desc: HTMLElement
}

/** Black or white, whichever reads better on a hex fill colour. */
function readableOn (fill: string): string {
  const m = /^#([0-9a-f]{6})$/i.exec(fill)
  if (!m) return '#000'
  const n = parseInt(m[1], 16)
  const lin = (v: number): number => {
    const c = v / 255
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4
  }
  const L = 0.2126 * lin(n >> 16) + 0.7152 * lin((n >> 8) & 255) + 0.0722 * lin(n & 255)
  // Contrast with black is (L + 0.05) / 0.05, with white 1.05 / (L + 0.05).
  return (L + 0.05) / 0.05 >= 1.05 / (L + 0.05) ? '#000' : '#fff'
}

function clamp (v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v
}

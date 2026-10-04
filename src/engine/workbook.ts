import { MAX_COLS, MAX_ROWS } from './address.js'
import { CompileError, compileFormula, type CompileEnv, type CompiledFormula } from './compiler.js'
import { formatValue, parseInput } from './format.js'
import type { FnContext } from './functions.js'
import { renameSheetInFormula, shiftFormula, transformFormulaForStructure, transformSpan } from './refs.js'
import { tokenize } from './tokenizer.js'
import { compareScalars, Errors, FormulaError, toScalar, type RangeValue, type Scalar } from './values.js'

export interface Style {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  align?: 'left' | 'center' | 'right'
  color?: string
  fill?: string
  /** Number format code, e.g. "0.00%" */
  fmt?: string
}

export const CLEAN = 0
export const DIRTY = 1
export const EVALUATING = 2

interface RangeListener {
  sheetId: number
  r1: number
  c1: number
  r2: number
  c2: number
  cell: Cell
}

/** Ranges wider than this many columns go in a per-sheet list instead of per-column buckets. */
const WIDE_RANGE_COLS = 64

export class Cell {
  /** Exactly what the user typed (formulas include the "="). */
  raw = ''
  value: Scalar = null
  /** @internal */
  formula: CompiledFormula | null = null
  /** Parse/compile error message, for the UI. */
  error: string | undefined
  style: Style | null = null
  /** Cached display text; undefined means it must be re-rendered. */
  display: string | undefined
  /** @internal */
  state = CLEAN
  /** @internal */
  depKeys: number[] | null = null
  /** @internal */
  rangeRegs: RangeListener[] | null = null

  constructor (readonly sheet: Sheet, public row: number, public col: number) {}
}

export class Sheet {
  readonly cols = new Map<number, Map<number, Cell>>()
  readonly colWidths = new Map<number, number>()
  readonly rowHeights = new Map<number, number>()
  /** Upper bounds of the used area (never shrinks until a rebuild). */
  maxRow = -1
  maxCol = -1
  cellCount = 0
  /** @internal */
  env!: CompileEnv

  constructor (readonly id: number, public name: string) {}

  get (r: number, c: number): Cell | undefined {
    return this.cols.get(c)?.get(r)
  }

  ensure (r: number, c: number): Cell {
    let col = this.cols.get(c)
    if (!col) {
      col = new Map()
      this.cols.set(c, col)
    }
    let cell = col.get(r)
    if (!cell) {
      cell = new Cell(this, r, c)
      col.set(r, cell)
      this.cellCount++
      if (r > this.maxRow) this.maxRow = r
      if (c > this.maxCol) this.maxCol = c
    }
    return cell
  }

  remove (r: number, c: number): void {
    const col = this.cols.get(c)
    if (col?.delete(r)) {
      this.cellCount--
      if (col.size === 0) this.cols.delete(c)
    }
  }

  * cells (): IterableIterator<Cell> {
    for (const col of this.cols.values()) yield * col.values()
  }

  recomputeBounds (): void {
    let mr = -1
    let mc = -1
    let n = 0
    for (const [c, col] of this.cols) {
      if (col.size && c > mc) mc = c
      for (const r of col.keys()) {
        n++
        if (r > mr) mr = r
      }
    }
    this.maxRow = mr
    this.maxCol = mc
    this.cellCount = n
  }
}

class SheetRange implements RangeValue {
  readonly kind = 'range' as const
  constructor (
    private readonly wb: Workbook,
    readonly sheet: number,
    readonly r1: number,
    readonly c1: number,
    readonly r2: number,
    readonly c2: number
  ) {}

  get (dr: number, dc: number): Scalar {
    return this.wb.read(this.sheet, this.r1 + dr, this.c1 + dc)
  }

  usedRows (): number {
    const s = this.wb.sheetSlot(this.sheet)
    return s ? Math.min(this.r2, s.maxRow) - this.r1 + 1 : 0
  }

  usedCols (): number {
    const s = this.wb.sheetSlot(this.sheet)
    return s ? Math.min(this.c2, s.maxCol) - this.c1 + 1 : 0
  }

  /**
   * Visit every non-empty cell, in no particular order. If any cell in the
   * range is still pending, all of them are recorded and evaluation aborts.
   */
  forEach (fn: (v: Scalar, r: number, c: number) => void): void {
    // The hottest loop in the engine (SUM(A:A) and friends), so it's written
    // out by hand rather than going through scan()'s per-cell callback.
    const s = this.wb.sheetSlot(this.sheet)
    if (!s) return
    const { r1, c1 } = this
    const r2 = Math.min(this.r2, s.maxRow)
    const c2 = Math.min(this.c2, s.maxCol)
    if (r2 < r1 || c2 < c1) return
    const missing = this.wb.missing
    const before = missing.length
    const wide = c2 - c1 + 1 > s.cols.size
    for (const [c, col] of wide ? s.cols : this.colsIn(s, c1, c2)) {
      if (c < c1 || c > c2) continue
      if (r2 - r1 + 1 > col.size) {
        for (const cell of col.values()) {
          const r = cell.row
          if (r < r1 || r > r2) continue
          if (cell.state !== CLEAN) missing.push(cell)
          else if (cell.value !== null) fn(cell.value, r - r1, c - c1)
        }
      } else {
        for (let r = r1; r <= r2; r++) {
          const cell = col.get(r)
          if (cell === undefined) continue
          if (cell.state !== CLEAN) missing.push(cell)
          else if (cell.value !== null) fn(cell.value, r - r1, c - c1)
        }
      }
    }
    if (missing.length > before) throw PENDING
  }

  /** [col, map] pairs for a narrow column span. */
  private colsIn (s: Sheet, c1: number, c2: number): [number, Map<number, Cell>][] {
    const out: [number, Map<number, Cell>][] = []
    for (let c = c1; c <= c2; c++) {
      const col = s.cols.get(c)
      if (col) out.push([c, col])
    }
    return out
  }

  /** Like forEach, but row by row, left to right (for CONCAT, NPV…). */
  forEachOrdered (fn: (v: Scalar, r: number, c: number) => void): void {
    const missing = this.wb.missing
    const before = missing.length
    const cells: Cell[] = []
    this.scan((cell) => {
      if (cell.state !== CLEAN) missing.push(cell)
      else if (cell.value !== null) cells.push(cell)
    })
    if (missing.length > before) throw PENDING
    cells.sort((a, b) => a.row - b.row || a.col - b.col)
    for (const cell of cells) fn(cell.value, cell.row - this.r1, cell.col - this.c1)
  }

  /** Every stored cell inside the range, walking the sparse maps. */
  private scan (visit: (cell: Cell) => void): void {
    const s = this.wb.sheetSlot(this.sheet)
    if (!s) return
    const { r1, c1 } = this
    const r2 = Math.min(this.r2, s.maxRow)
    const c2 = Math.min(this.c2, s.maxCol)
    if (r2 < r1 || c2 < c1) return
    const visitCol = (col: Map<number, Cell>): void => {
      if (r2 - r1 + 1 > col.size) {
        for (const cell of col.values()) if (cell.row >= r1 && cell.row <= r2) visit(cell)
      } else {
        for (let r = r1; r <= r2; r++) {
          const cell = col.get(r)
          if (cell !== undefined) visit(cell)
        }
      }
    }
    if (c2 - c1 + 1 > s.cols.size) {
      for (const [c, col] of s.cols) if (c >= c1 && c <= c2) visitCol(col)
    } else {
      for (let c = c1; c <= c2; c++) {
        const col = s.cols.get(c)
        if (col) visitCol(col)
      }
    }
  }
}

export interface CellPatch {
  r: number
  c: number
  /** New raw input; omitted to leave it unchanged. */
  raw?: string
  /** New style; null clears it, omitted leaves it unchanged. */
  style?: Style | null
}

export interface SheetJSON {
  /** Stable id, so undo history keeps pointing at the right sheet across reloads. */
  id?: number
  name: string
  cells: [number, number, string, Style?][]
  colWidths?: [number, number][]
  rowHeights?: [number, number][]
}

export interface WorkbookJSON {
  version: 1
  sheets: SheetJSON[]
}

export interface RecalcStats {
  evaluated: number
  ms: number
}

const CELLS_PER_SHEET = MAX_ROWS * MAX_COLS

/**
 * Thrown (not returned) when a formula reads a cell that hasn't been
 * calculated yet. Aborting at once means the rest of the formula — including
 * an IF branch chosen from a placeholder value — never runs on bad inputs.
 */
const PENDING = { pending: true } as const

function cellKey (sheetId: number, r: number, c: number): number {
  return sheetId * CELLS_PER_SHEET + r * MAX_COLS + c
}

function styleEmpty (s: Style | null | undefined): boolean {
  if (!s) return true
  for (const k in s) if ((s as Record<string, unknown>)[k] !== undefined) return false
  return true
}

/**
 * The workbook owns cells, the dependency graph and recalculation.
 *
 * Recalc is incremental: an edit marks its transitive dependents dirty, and
 * only those are re-evaluated. Evaluation uses an explicit stack (no
 * recursion), so chains of a million dependent cells are fine. When a formula
 * reads a cell that is still dirty, that read is recorded; the formula is
 * re-queued above its missing inputs and retried once they're clean. Reading
 * a cell that is currently being evaluated means a cycle.
 */
export class Workbook {
  sheets: Sheet[] = []
  private slots: (Sheet | undefined)[] = []
  private nextSheetId = 1
  private readonly cellDeps = new Map<number, Set<Cell>>()
  private readonly colListeners = new Map<number, Map<number, Set<RangeListener>>>()
  private readonly wideListeners = new Map<number, Set<RangeListener>>()
  private readonly volatileCells = new Set<Cell>()
  private dirty: Cell[] = []
  /** Set when some formula names a sheet that doesn't exist (yet). */
  private unresolvedSheetRefs = false
  /** Cells whose value changed since the last recalc, as [sheetId, row, col] triples. */
  private changed: number[] = []
  /** Dirty cells read during the current evaluation (see class docs). */
  /** @internal */
  readonly missing: Cell[] = []
  private batchDepth = 0
  private readonly listeners = new Set<() => void>()
  lastRecalc: RecalcStats = { evaluated: 0, ms: 0 }

  private readonly ctx: FnContext & { sheetId: number } = {
    row: 0,
    col: 0,
    sheetId: 0,
    makeRange: (sheetId, r1, c1, r2, c2) => {
      if (r1 < 0 || c1 < 0 || r2 >= MAX_ROWS || c2 >= MAX_COLS || r2 < r1 || c2 < c1) return Errors.REF
      return new SheetRange(this, sheetId, r1, c1, r2, c2)
    },
    rangeFromText: (text, sheetId) => this.rangeFromText(text, sheetId)
  }

  constructor () {
    this.addSheet('Sheet1')
  }

  // -------------------------------------------------------------------------
  // Sheets

  /** The sheet with this id, if it still exists. */
  sheetSlot (id: number): Sheet | undefined {
    return this.slots[id]
  }

  getSheet (id: number): Sheet {
    const s = this.slots[id]
    if (!s) throw new Error(`No sheet ${id}`)
    return s
  }

  sheetByName (name: string): Sheet | undefined {
    const lower = name.toLowerCase()
    return this.sheets.find((s) => s.name.toLowerCase() === lower)
  }

  addSheet (name?: string, index = this.sheets.length): Sheet {
    let n = this.sheets.length + 1
    let finalName = name ?? `Sheet${n}`
    while (this.sheetByName(finalName)) finalName = name ? `${name} (${++n})` : `Sheet${++n}`
    const sheet = new Sheet(this.nextSheetId++, finalName)
    sheet.env = this.makeEnv(sheet)
    this.slots[sheet.id] = sheet
    this.sheets.splice(index, 0, sheet)
    // Formulas that referred to a missing sheet of this name can now resolve.
    if (this.unresolvedSheetRefs) this.rebuild()
    return sheet
  }

  removeSheet (id: number): void {
    if (this.sheets.length <= 1) return
    this.sheets = this.sheets.filter((s) => s.id !== id)
    this.slots[id] = undefined
    this.rebuild()
  }

  renameSheet (id: number, name: string): boolean {
    const sheet = this.getSheet(id)
    const trimmed = name.trim()
    if (!trimmed || /[[\]:*?/\\]/.test(trimmed)) return false
    const clash = this.sheetByName(trimmed)
    if (clash && clash !== sheet) return false
    const old = sheet.name
    sheet.name = trimmed
    for (const s of this.sheets) {
      for (const cell of s.cells()) {
        if (cell.raw[0] === '=') cell.raw = renameSheetInFormula(cell.raw, old, trimmed)
      }
    }
    this.rebuild()
    return true
  }

  moveSheet (id: number, toIndex: number): void {
    const i = this.sheets.findIndex((s) => s.id === id)
    if (i < 0) return
    const [s] = this.sheets.splice(i, 1)
    this.sheets.splice(Math.max(0, Math.min(toIndex, this.sheets.length)), 0, s)
    this.emit()
  }

  private makeEnv (sheet: Sheet): CompileEnv {
    return {
      resolveSheet: (name) => {
        if (name === undefined) return sheet.id
        const id = this.sheetByName(name)?.id
        if (id === undefined) this.unresolvedSheetRefs = true
        return id
      },
      read: this.read,
      range: (sheetId, r1, c1, r2, c2) => new SheetRange(this, sheetId, r1, c1, r2, c2)
    }
  }

  // -------------------------------------------------------------------------
  // Reading

  /** Hot path: every cell reference in every formula goes through here. */
  /** @internal */
  readonly read = (sheetId: number, r: number, c: number): Scalar => {
    const sheet = this.slots[sheetId]
    if (sheet === undefined) return Errors.REF
    const col = sheet.cols.get(c)
    if (col === undefined) return null
    const cell = col.get(r)
    if (cell === undefined) return null
    if (cell.state !== CLEAN) {
      this.missing.push(cell)
      throw PENDING
    }
    return cell.value
  }

  getCell (sheetId: number, r: number, c: number): Cell | undefined {
    return this.slots[sheetId]?.get(r, c)
  }

  getValue (sheetId: number, r: number, c: number): Scalar {
    return this.getCell(sheetId, r, c)?.value ?? null
  }

  /** Text to draw for a cell, cached until its value or format changes. */
  getDisplay (cell: Cell): string {
    let d = cell.display
    if (d === undefined) {
      d = formatValue(cell.value, cell.style?.fmt)
      cell.display = d
    }
    return d
  }

  private rangeFromText (text: string, sheetId: number): RangeValue | FormulaError {
    try {
      const toks = tokenize(text.trim())
      if (toks.length !== 1 || toks[0].type !== 'ref') return Errors.REF
      const ref = toks[0].ref!
      const sid = ref.sheet === undefined ? sheetId : this.sheetByName(ref.sheet)?.id
      if (sid === undefined) return Errors.REF
      return new SheetRange(this, sid, Math.min(ref.r1, ref.r2), Math.min(ref.c1, ref.c2), Math.max(ref.r1, ref.r2), Math.max(ref.c1, ref.c2))
    } catch {
      return Errors.REF
    }
  }

  // -------------------------------------------------------------------------
  // Writing

  onChange (fn: () => void): () => void {
    this.listeners.add(fn)
    return () => this.listeners.delete(fn)
  }

  private emit (): void {
    for (const fn of this.listeners) fn()
  }

  /** Run several edits and recalculate once at the end. */
  batch<T> (fn: () => T): T {
    this.batchDepth++
    try {
      return fn()
    } finally {
      if (--this.batchDepth === 0) this.recalc()
    }
  }

  /** Set one cell's input. */
  setCell (sheetId: number, r: number, c: number, raw: string): void {
    this.applyPatches(sheetId, [{ r, c, raw }])
  }

  /**
   * Apply cell edits as one batch and return the patches that undo them.
   * `hints` lets typed values like "12%" or "$5" pick a number format.
   */
  applyPatches (sheetId: number, patches: CellPatch[], hints = false): CellPatch[] {
    const sheet = this.getSheet(sheetId)
    const inverse: CellPatch[] = []
    this.batch(() => {
      for (const p of patches) {
        const before = sheet.get(p.r, p.c)
        inverse.push({ r: p.r, c: p.c, raw: before?.raw ?? '', style: before?.style ?? null })
        if (p.style !== undefined) this.setStyle(sheet, p.r, p.c, p.style)
        if (p.raw !== undefined) this.setRaw(sheet, p.r, p.c, p.raw, hints)
      }
    })
    inverse.reverse()
    return inverse
  }

  private setStyle (sheet: Sheet, r: number, c: number, style: Style | null): void {
    const next = styleEmpty(style) ? null : style
    let cell = sheet.get(r, c)
    if (!cell) {
      if (!next) return
      cell = sheet.ensure(r, c)
    }
    const fmtChanged = (cell.style?.fmt ?? '') !== (next?.fmt ?? '')
    cell.style = next
    cell.display = undefined
    // Switching to/from text format changes how the raw input is interpreted.
    if (fmtChanged && cell.raw !== '' && (cell.style?.fmt === '@' || !cell.formula)) this.setRaw(sheet, r, c, cell.raw, false, true)
    this.maybeRemove(cell)
  }

  private setRaw (sheet: Sheet, r: number, c: number, raw: string, hints: boolean, force = false): void {
    let cell = sheet.get(r, c)
    if (!cell) {
      if (raw === '') return
      cell = sheet.ensure(r, c)
    }
    if (cell.raw === raw && !force) return
    this.detach(cell)
    cell.raw = raw
    cell.display = undefined
    cell.error = undefined
    const textFormat = cell.style?.fmt === '@'
    const parsed = textFormat ? { kind: raw === '' ? 'empty' : 'value', value: raw === '' ? null : raw } as const : parseInput(raw)
    if (parsed.kind === 'formula') {
      this.attach(cell)
    } else {
      cell.value = parsed.value
      cell.state = CLEAN
      if (hints && 'format' in parsed && parsed.format && !cell.style?.fmt) cell.style = { ...cell.style, fmt: parsed.format }
    }
    this.changed.push(sheet.id, r, c)
    this.maybeRemove(cell)
  }

  private maybeRemove (cell: Cell): void {
    if (cell.raw === '' && !cell.style && !cell.formula) cell.sheet.remove(cell.row, cell.col)
  }

  private attach (cell: Cell): void {
    let compiled: CompiledFormula
    try {
      compiled = compileFormula(cell.raw.slice(1), cell.sheet.env)
    } catch (e) {
      cell.formula = null
      cell.value = Errors.PARSE
      cell.error = e instanceof CompileError ? e.message : String(e)
      cell.state = CLEAN
      return
    }
    cell.formula = compiled
    for (const p of compiled.precedents) {
      if (p.r1 === p.r2 && p.c1 === p.c2) {
        const key = cellKey(p.sheetId, p.r1, p.c1)
        let set = this.cellDeps.get(key)
        if (!set) {
          set = new Set()
          this.cellDeps.set(key, set)
        }
        set.add(cell)
        ;(cell.depKeys ??= []).push(key)
      } else {
        const l: RangeListener = { ...p, cell }
        ;(cell.rangeRegs ??= []).push(l)
        if (p.c2 - p.c1 + 1 > WIDE_RANGE_COLS) {
          let set = this.wideListeners.get(p.sheetId)
          if (!set) this.wideListeners.set(p.sheetId, (set = new Set()))
          set.add(l)
        } else {
          let bySheet = this.colListeners.get(p.sheetId)
          if (!bySheet) this.colListeners.set(p.sheetId, (bySheet = new Map()))
          for (let c = p.c1; c <= p.c2; c++) {
            let set = bySheet.get(c)
            if (!set) bySheet.set(c, (set = new Set()))
            set.add(l)
          }
        }
      }
    }
    if (compiled.volatile) this.volatileCells.add(cell)
    cell.state = DIRTY
    this.dirty.push(cell)
  }

  private detach (cell: Cell): void {
    if (!cell.formula) return
    if (cell.depKeys) {
      for (const key of cell.depKeys) {
        const set = this.cellDeps.get(key)
        if (set) {
          set.delete(cell)
          if (set.size === 0) this.cellDeps.delete(key)
        }
      }
      cell.depKeys = null
    }
    if (cell.rangeRegs) {
      for (const l of cell.rangeRegs) {
        if (l.c2 - l.c1 + 1 > WIDE_RANGE_COLS) this.wideListeners.get(l.sheetId)?.delete(l)
        else {
          const bySheet = this.colListeners.get(l.sheetId)
          if (bySheet) {
            for (let c = l.c1; c <= l.c2; c++) {
              const set = bySheet.get(c)
              if (set) {
                set.delete(l)
                if (set.size === 0) bySheet.delete(c)
              }
            }
          }
        }
      }
      cell.rangeRegs = null
    }
    this.volatileCells.delete(cell)
    cell.formula = null
    cell.state = CLEAN
  }

  /**
   * Mark every formula that (transitively) reads any of the seed positions
   * as dirty. Range listeners are copied into a per-call worklist the first
   * time their bucket is hit, and removed from it once marked, so a bucket of
   * n listeners costs O(n) per call rather than O(n) per visited cell.
   */
  private markDependents (seeds: number[]): void {
    const queue: number[] = seeds.slice()
    const dirty = this.dirty
    const pendingBuckets = new Map<Set<RangeListener>, RangeListener[]>()
    const mark = (d: Cell): void => {
      d.state = DIRTY
      dirty.push(d)
      queue.push(d.sheet.id, d.row, d.col)
    }
    const scanBucket = (bucket: Set<RangeListener>, row: number, col: number, checkCol: boolean): void => {
      let list = pendingBuckets.get(bucket)
      if (!list) {
        list = []
        for (const l of bucket) if (l.cell.state === CLEAN) list.push(l)
        pendingBuckets.set(bucket, list)
      }
      for (let i = 0; i < list.length;) {
        const l = list[i]
        if (l.cell.state !== CLEAN) {
          list[i] = list[list.length - 1]
          list.pop()
          continue
        }
        if (row >= l.r1 && row <= l.r2 && (!checkCol || (col >= l.c1 && col <= l.c2))) {
          mark(l.cell)
          list[i] = list[list.length - 1]
          list.pop()
          continue
        }
        i++
      }
    }
    while (queue.length) {
      const col = queue.pop()!
      const row = queue.pop()!
      const sid = queue.pop()!
      const set = this.cellDeps.get(cellKey(sid, row, col))
      if (set) for (const d of set) if (d.state === CLEAN) mark(d)
      const bucket = this.colListeners.get(sid)?.get(col)
      if (bucket) scanBucket(bucket, row, col, false)
      const wide = this.wideListeners.get(sid)
      if (wide) scanBucket(wide, row, col, true)
    }
  }

  // -------------------------------------------------------------------------
  // Recalculation

  recalc (): void {
    const t0 = performance.now()
    const seeds = this.changed
    this.changed = []
    for (const v of this.volatileCells) {
      if (v.state === CLEAN) {
        v.state = DIRTY
        this.dirty.push(v)
        seeds.push(v.sheet.id, v.row, v.col)
      }
    }
    this.markDependents(seeds)
    const list = this.dirty
    this.dirty = []
    let evaluated = 0
    for (const cell of list) {
      if (cell.state !== CLEAN) evaluated += this.compute(cell)
    }
    this.lastRecalc = { evaluated, ms: performance.now() - t0 }
    this.emit()
  }

  private compute (root: Cell): number {
    const stack: Cell[] = [root]
    const missing = this.missing
    const ctx = this.ctx
    let evaluated = 0
    while (stack.length) {
      const cell = stack[stack.length - 1]
      if (cell.state === CLEAN) {
        stack.pop()
        continue
      }
      const f = cell.formula
      if (!f) {
        cell.state = CLEAN
        stack.pop()
        continue
      }
      cell.state = EVALUATING
      missing.length = 0
      ctx.row = cell.row
      ctx.col = cell.col
      ctx.sheetId = cell.sheet.id
      let v: Scalar = null
      try {
        v = toScalar(f.evaluate(ctx))
      } catch (e) {
        if (e !== PENDING) {
          console.error('Formula evaluation failed', cell.raw, e)
          v = Errors.VALUE
          missing.length = 0
        }
      }
      evaluated++
      if (missing.length === 0) {
        this.finish(cell, v)
        stack.pop()
        continue
      }
      let cyclic = false
      for (let i = 0; i < missing.length; i++) {
        if (missing[i].state === EVALUATING) {
          cyclic = true
          break
        }
      }
      if (cyclic) {
        this.finish(cell, Errors.CIRC)
        stack.pop()
        continue
      }
      for (let i = 0; i < missing.length; i++) {
        if (missing[i].state === DIRTY) stack.push(missing[i])
      }
    }
    missing.length = 0
    return evaluated
  }

  private finish (cell: Cell, v: Scalar): void {
    const value = v === null ? 0 : v
    if (value !== cell.value) cell.display = undefined
    cell.value = value
    cell.state = CLEAN
  }

  /** Recompile everything (after structural changes or sheet renames). */
  rebuild (): void {
    this.cellDeps.clear()
    this.colListeners.clear()
    this.wideListeners.clear()
    this.volatileCells.clear()
    this.dirty = []
    this.changed = []
    this.unresolvedSheetRefs = false
    this.batch(() => {
      for (const s of this.sheets) {
        s.recomputeBounds()
        for (const cell of s.cells()) {
          cell.formula = null
          cell.depKeys = null
          cell.rangeRegs = null
          cell.state = CLEAN
          cell.display = undefined
          if (cell.raw.length > 1 && cell.raw[0] === '=' && cell.style?.fmt !== '@') this.attach(cell)
        }
      }
    })
  }

  // -------------------------------------------------------------------------
  // Structure

  insertRows (sheetId: number, at: number, count: number): void {
    this.structural(sheetId, 'rows', at, count)
  }

  deleteRows (sheetId: number, at: number, count: number): void {
    this.structural(sheetId, 'rows', at, -count)
  }

  insertCols (sheetId: number, at: number, count: number): void {
    this.structural(sheetId, 'cols', at, count)
  }

  deleteCols (sheetId: number, at: number, count: number): void {
    this.structural(sheetId, 'cols', at, -count)
  }

  private structural (sheetId: number, axis: 'rows' | 'cols', at: number, count: number): void {
    const target = this.getSheet(sheetId)
    for (const s of this.sheets) {
      const onSheet = (name: string | undefined): boolean => (name === undefined ? s === target : this.sheetByName(name) === target)
      for (const cell of s.cells()) {
        if (cell.raw[0] === '=') cell.raw = transformFormulaForStructure(cell.raw, axis, at, count, onSheet)
      }
    }
    const moved: Cell[] = []
    for (const cell of target.cells()) {
      const pos = axis === 'rows' ? cell.row : cell.col
      const span = transformSpan(pos, pos, at, count)
      if (!span) continue
      if (axis === 'rows') cell.row = span[0]
      else cell.col = span[0]
      if (cell.row < MAX_ROWS && cell.col < MAX_COLS) moved.push(cell)
    }
    target.cols.clear()
    for (const cell of moved) {
      let col = target.cols.get(cell.col)
      if (!col) target.cols.set(cell.col, (col = new Map()))
      col.set(cell.row, cell)
    }
    const sizes = axis === 'rows' ? target.rowHeights : target.colWidths
    const entries = [...sizes]
    sizes.clear()
    for (const [i, size] of entries) {
      const span = transformSpan(i, i, at, count)
      if (span) sizes.set(span[0], size)
    }
    this.rebuild()
  }

  /**
   * Patches that sort a block of rows by one column. Formulas move with their
   * row (relative references are shifted like a cut/paste would).
   */
  sortPatches (sheetId: number, r1: number, c1: number, r2: number, c2: number, keyCol: number, ascending: boolean): CellPatch[] {
    const sheet = this.getSheet(sheetId)
    const rows: number[] = []
    for (let r = r1; r <= Math.min(r2, sheet.maxRow); r++) rows.push(r)
    const keyOf = (r: number): Scalar => sheet.get(r, keyCol)?.value ?? null
    rows.sort((a, b) => {
      const x = keyOf(a)
      const y = keyOf(b)
      // Blanks always sort last, errors just before them.
      if (x === null || y === null) return x === null ? (y === null ? 0 : 1) : -1
      const c = compareScalars(x, y)
      return (ascending ? c : -c) || a - b
    })
    const patches: CellPatch[] = []
    rows.forEach((from, i) => {
      const to = r1 + i
      if (from === to) return
      for (let c = c1; c <= Math.min(c2, sheet.maxCol); c++) {
        const src = sheet.get(from, c)
        const raw = src?.raw ?? ''
        patches.push({ r: to, c, raw: raw[0] === '=' ? shiftFormula(raw, to - from, 0) : raw, style: src?.style ?? null })
      }
    })
    return patches
  }

  // -------------------------------------------------------------------------
  // Serialisation

  toJSON (): WorkbookJSON {
    return {
      version: 1,
      sheets: this.sheets.map((s) => {
        const cells: SheetJSON['cells'] = []
        for (const cell of s.cells()) {
          if (cell.raw === '' && !cell.style) continue
          cells.push(cell.style ? [cell.row, cell.col, cell.raw, cell.style] : [cell.row, cell.col, cell.raw])
        }
        const out: SheetJSON = { id: s.id, name: s.name, cells }
        if (s.colWidths.size) out.colWidths = [...s.colWidths]
        if (s.rowHeights.size) out.rowHeights = [...s.rowHeights]
        return out
      })
    }
  }

  static fromJSON (json: WorkbookJSON): Workbook {
    const wb = new Workbook()
    wb.load(json)
    return wb
  }

  /**
   * Replace the whole workbook contents. The input is validated first, so a
   * bad file throws and leaves the current workbook untouched.
   */
  load (json: WorkbookJSON): void {
    validateWorkbookJSON(json)
    this.sheets = []
    this.slots = []
    const sheetsIn: SheetJSON[] = json.sheets.length ? json.sheets : [{ name: 'Sheet1', cells: [] }]
    // Keep ids that are present and unique; give the rest fresh ones.
    const taken = new Set<number>()
    for (const sj of sheetsIn) if (sj.id !== undefined && sj.id > 0 && !taken.has(sj.id)) taken.add(sj.id)
    let next = Math.max(0, ...taken) + 1
    const usedIds = new Set<number>()
    const usedNames = new Set<string>()
    for (const sj of sheetsIn) {
      const id = sj.id !== undefined && taken.has(sj.id) && !usedIds.has(sj.id) ? sj.id : next++
      usedIds.add(id)
      let name = sj.name
      for (let n = 2; usedNames.has(name.toLowerCase()); n++) name = `${sj.name} (${n})`
      usedNames.add(name.toLowerCase())
      const sheet = new Sheet(id, name)
      sheet.env = this.makeEnv(sheet)
      this.slots[sheet.id] = sheet
      this.sheets.push(sheet)
      for (const [r, c, raw, style] of sj.cells) {
        const cell = sheet.ensure(r, c)
        cell.raw = raw
        cell.style = style ?? null
        if (raw[0] !== '=' || raw.length < 2 || style?.fmt === '@') {
          cell.value = style?.fmt === '@' ? (raw === '' ? null : raw) : parseInput(raw).value
        }
      }
      for (const [i, w] of sj.colWidths ?? []) sheet.colWidths.set(i, w)
      for (const [i, h] of sj.rowHeights ?? []) sheet.rowHeights.set(i, h)
    }
    this.nextSheetId = next
    this.rebuild()
  }

  /** Total number of formula cells (for the status bar / benchmarks). */
  formulaCount (): number {
    let n = 0
    for (const s of this.sheets) for (const cell of s.cells()) if (cell.formula) n++
    return n
  }
}

/** Throw a readable error unless `json` looks like a saved workbook. */
export function validateWorkbookJSON (json: unknown): asserts json is WorkbookJSON {
  const fail = (why: string): never => {
    throw new Error(`Not a FilmSheet workbook (${why})`)
  }
  if (typeof json !== 'object' || json === null) fail('not an object')
  const sheets = (json as WorkbookJSON).sheets
  if (!Array.isArray(sheets)) fail('missing "sheets"')
  for (const sj of sheets) {
    if (typeof sj !== 'object' || sj === null || typeof sj.name !== 'string' || !sj.name.trim()) fail('a sheet has no name')
    if (!Array.isArray(sj.cells)) fail(`sheet "${sj.name}" has no cells array`)
    for (const cell of sj.cells) {
      if (!Array.isArray(cell)) fail('bad cell entry')
      const [r, c, raw] = cell
      if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS || typeof raw !== 'string') {
        fail(`bad cell ${JSON.stringify(cell).slice(0, 40)}`)
      }
    }
  }
}

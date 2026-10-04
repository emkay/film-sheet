import type { CellPatch, WorkbookJSON } from '../engine/workbook.js'
import type { Selection } from './grid.js'

export type HistoryEntry =
  | { kind: 'patch'; sheetId: number; undo: CellPatch[]; redo: CellPatch[]; selection: Selection }
  /** Whole-workbook change. `after` is filled in when the entry is first undone. Sheets are identified by id. */
  | { kind: 'snapshot'; before: WorkbookJSON; after?: WorkbookJSON; sheetBefore: number; sheetAfter: number }
  | { kind: 'size'; sheetId: number; axis: 'rows' | 'cols'; index: number; before: number; after: number }
  /** Several entries undone/redone as one step (e.g. a cut pasted onto another sheet). */
  | { kind: 'group'; entries: HistoryEntry[] }

const LIMIT = 200

export class History {
  private undoStack: HistoryEntry[] = []
  private redoStack: HistoryEntry[] = []

  push (entry: HistoryEntry): void {
    this.undoStack.push(entry)
    if (this.undoStack.length > LIMIT) this.undoStack.shift()
    this.redoStack = []
  }

  undo (): HistoryEntry | undefined {
    const e = this.undoStack.pop()
    if (e) this.redoStack.push(e)
    return e
  }

  redo (): HistoryEntry | undefined {
    const e = this.redoStack.pop()
    if (e) this.undoStack.push(e)
    return e
  }

  /** Combine the last `n` entries into one undo step. */
  group (n: number): void {
    if (n < 2 || this.undoStack.length < n) return
    const entries = this.undoStack.splice(this.undoStack.length - n, n)
    this.undoStack.push({ kind: 'group', entries })
  }

  clear (): void {
    this.undoStack = []
    this.redoStack = []
  }

  get canUndo (): boolean {
    return this.undoStack.length > 0
  }

  get canRedo (): boolean {
    return this.redoStack.length > 0
  }
}

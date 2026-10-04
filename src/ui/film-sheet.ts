import { LitElement, unsafeCSS, type PropertyValues } from 'lit'
import type { Workbook, WorkbookJSON } from '../engine/workbook.js'
import { SheetController } from './controller.js'
import styles from './styles.css?inline'
import { sheetTemplate } from './template.js'

export type SheetChrome = 'full' | 'minimal' | 'none'

/** `detail` of the `film-sheet-select` event. */
export interface SheetSelectDetail {
  /** Name of the active sheet. */
  sheet: string
  /** The active cell, e.g. "B3". */
  address: string
  /** The selected range, e.g. "B3:D9" (or the address for a single cell). */
  range: string
}

/**
 * <film-sheet> — an Excel-like spreadsheet with formulas, built on Film.
 *
 * The page must load Film's theme CSS (`@mk/film/css/themes/default/index.css`);
 * the component follows whatever Film theme and colour scheme the page uses.
 *
 * @fires film-sheet-change - After any change to the workbook's cells or sheets.
 * @fires film-sheet-select - When the selection or active sheet changes. `detail` is a {@link SheetSelectDetail}.
 */
export class FilmSheet extends LitElement {
  static override properties = {
    readonly: { type: Boolean, reflect: true },
    storageKey: { type: String, attribute: 'storage-key' },
    chrome: { type: String, reflect: true },
    filename: { type: String }
  }

  static override styles = unsafeCSS(styles)

  /** Prevent all edits; viewing, selecting, copying and searching still work. */
  declare readonly: boolean
  /**
   * localStorage key to autosave to and restore from. Empty (the default)
   * means no persistence; use `toJSON()` / `load()` and the change event.
   */
  declare storageKey: string
  /** How much UI to show around the grid. */
  declare chrome: SheetChrome
  /** Base name for downloads (Save as .json, Export CSV). */
  declare filename: string

  private controller?: SheetController
  /** Data waiting for the controller (set before first render or while disconnected). */
  private pending?: WorkbookJSON

  constructor () {
    super()
    this.readonly = false
    this.storageKey = ''
    this.chrome = 'full'
    this.filename = 'workbook'
  }

  /** The live workbook engine, for reading values or scripting edits. */
  get workbook (): Workbook | undefined {
    return this.controller?.wb
  }

  /** The workbook as JSON; assigning replaces it. */
  get data (): WorkbookJSON | undefined {
    return this.toJSON()
  }

  set data (json: WorkbookJSON | undefined) {
    if (json) this.load(json)
  }

  /** Replace the workbook. Throws (leaving the current one intact) if `json` isn't a workbook. */
  load (json: WorkbookJSON): void {
    if (this.controller) this.controller.load(json)
    else this.pending = json
  }

  /** The workbook as JSON, suitable for `load()` or saving. */
  toJSON (): WorkbookJSON | undefined {
    return this.controller?.wb.toJSON() ?? this.pending
  }

  /** Move keyboard focus to the active cell. */
  override focus (): void {
    this.controller?.grid.focus()
  }

  /** Re-read theme colours (call after changing Film tokens on an ancestor yourself). */
  refreshTheme (): void {
    this.controller?.grid.refreshTheme()
  }

  /** @internal For tests and debugging: the controller behind the element. */
  get _controller (): SheetController | undefined {
    return this.controller
  }

  protected override render (): unknown {
    return sheetTemplate
  }

  protected override firstUpdated (): void {
    this.start()
  }

  protected override updated (changed: PropertyValues<this>): void {
    if (changed.has('readonly') && this.controller) this.controller.applyReadonly()
  }

  override connectedCallback (): void {
    super.connectedCallback()
    // Re-attached after being removed: rebuild from what we had.
    if (this.hasUpdated && !this.controller) this.start()
  }

  override disconnectedCallback (): void {
    super.disconnectedCallback()
    if (!this.controller) return
    this.pending = this.controller.wb.toJSON()
    this.controller.dispose()
    this.controller = undefined
  }

  private start (): void {
    const initial = this.pending
    this.pending = undefined
    this.controller = new SheetController(this, this.renderRoot as ShadowRoot, initial)
    if (this.readonly) this.controller.applyReadonly()
  }
}

if (!customElements.get('film-sheet')) customElements.define('film-sheet', FilmSheet)

declare global {
  interface HTMLElementTagNameMap {
    'film-sheet': FilmSheet
  }
  interface HTMLElementEventMap {
    'film-sheet-change': CustomEvent<null>
    'film-sheet-select': CustomEvent<SheetSelectDetail>
  }
}

import { FUNCTION_NAMES, FUNCTIONS } from '../engine/functions.js'

/** Characters after which a cell reference can be inserted by pointing. */
const REF_INSERTABLE = /[=(,+\-*/^&<>:;]\s*$/

/**
 * The in-cell editor: a textarea laid over the active cell. It knows about
 * formula editing conveniences — function autocomplete, the signature hint
 * for the argument under the caret, and "pointing" (inserting references by
 * clicking or using arrow keys while typing a formula).
 */
export class CellEditor {
  readonly el: HTMLTextAreaElement
  private readonly popup: HTMLElement
  private readonly hint: HTMLElement
  private readonly measureCtx = document.createElement('canvas').getContext('2d')!
  active = false
  /** "enter": started by typing, arrows commit; "edit": F2/double-click, arrows move the caret. */
  mode: 'enter' | 'edit' = 'enter'
  row = 0
  col = 0
  /** The span of a reference inserted by pointing, which further pointing replaces. */
  private pointSpan: { start: number; end: number } | null = null
  private suggestions: string[] = []
  private suggestIndex = 0
  /** True once the user has arrowed through the list; only then does Enter accept. */
  navigated = false
  private minWidth = 0
  private maxWidth = 0
  onInput?: (text: string) => void
  /** Short messages for the app's live region (suggestion counts, argument hints). */
  onAnnounce?: (text: string) => void
  private lastHint = ''
  private lastSuggestCount = -1

  constructor (host: HTMLElement) {
    this.el = document.createElement('textarea')
    this.el.className = 'cell-editor'
    this.el.spellcheck = false
    this.el.setAttribute('aria-label', 'Cell editor')
    this.el.wrap = 'off'
    this.popup = document.createElement('div')
    this.popup.className = 'autocomplete'
    this.popup.id = 'fn-suggestions'
    this.popup.setAttribute('role', 'listbox')
    this.popup.setAttribute('aria-label', 'Function suggestions')
    this.hint = document.createElement('div')
    this.hint.className = 'sig-hint'
    this.hint.id = 'fn-signature'
    // A textarea can't be a combobox, but it can own a list and point at the
    // highlighted option, which screen readers follow.
    this.el.setAttribute('aria-autocomplete', 'list')
    this.el.setAttribute('aria-controls', this.popup.id)
    this.el.setAttribute('aria-describedby', this.hint.id)
    host.append(this.el, this.popup, this.hint)
    this.el.addEventListener('input', () => {
      this.pointSpan = null
      this.afterChange()
    })
    this.el.addEventListener('keyup', (e) => {
      if (e.key.startsWith('Arrow') || e.key === 'Home' || e.key === 'End') this.updateHint()
    })
    this.el.addEventListener('click', () => this.updateHint())
    this.popup.addEventListener('pointerdown', (e) => {
      e.preventDefault()
      const item = (e.target as HTMLElement).closest<HTMLElement>('[data-fn]')
      if (item) this.accept(item.dataset.fn!)
    })
  }

  get value (): string {
    return this.el.value
  }

  set value (v: string) {
    this.el.value = v
    this.pointSpan = null
    this.afterChange()
  }

  get isFormula (): boolean {
    return this.el.value.startsWith('=')
  }

  open (row: number, col: number, text: string, mode: 'enter' | 'edit', rect: { x: number; y: number; w: number; h: number }, maxWidth: number, font: string): void {
    this.active = true
    this.mode = mode
    this.row = row
    this.col = col
    this.pointSpan = null
    this.minWidth = rect.w + 1
    this.maxWidth = Math.max(rect.w, maxWidth - rect.x - 8)
    this.measureCtx.font = font
    this.el.style.font = font
    Object.assign(this.el.style, {
      display: 'block',
      left: rect.x - 1 + 'px',
      top: rect.y - 1 + 'px',
      height: rect.h + 1 + 'px',
      width: this.minWidth + 'px'
    })
    this.el.value = text
    this.el.focus({ preventScroll: true })
    const end = text.length
    this.el.setSelectionRange(end, end)
    this.afterChange()
  }

  /** Follow the cell when the grid scrolls; hide while it's off-screen. */
  reposition (rect: { x: number; y: number; w: number; h: number }, visible: boolean): void {
    if (!this.active) return
    this.el.style.left = rect.x - 1 + 'px'
    this.el.style.top = rect.y - 1 + 'px'
    this.el.style.visibility = visible ? '' : 'hidden'
    const below = rect.y + rect.h + 1 + 'px'
    this.popup.style.left = this.hint.style.left = this.el.style.left
    this.popup.style.top = this.hint.style.top = below
    if (!visible) {
      this.popup.style.display = 'none'
      this.hint.style.display = 'none'
    }
  }

  /** Remove the editor's elements from the page. */
  dispose (): void {
    this.el.remove()
    this.popup.remove()
    this.hint.remove()
  }

  close (): void {
    this.el.style.visibility = ''
    this.lastHint = ''
    this.lastSuggestCount = -1
    this.active = false
    this.el.style.display = 'none'
    this.hidePopup()
    this.hint.style.display = 'none'
    this.pointSpan = null
  }

  /** Is the caret somewhere a reference can be pointed in? */
  canPoint (): boolean {
    if (!this.active || !this.isFormula) return false
    if (this.pointSpan) return true
    const pos = this.el.selectionStart
    if (pos !== this.el.selectionEnd) return false
    return REF_INSERTABLE.test(this.el.value.slice(0, pos))
  }

  /** Insert or replace the pointed reference. */
  point (ref: string): void {
    const v = this.el.value
    const start = this.pointSpan ? this.pointSpan.start : this.el.selectionStart
    const end = this.pointSpan ? this.pointSpan.end : this.el.selectionEnd
    this.el.value = v.slice(0, start) + ref + v.slice(end)
    this.pointSpan = { start, end: start + ref.length }
    this.el.setSelectionRange(start + ref.length, start + ref.length)
    this.afterChange()
  }

  get pointing (): boolean {
    return this.pointSpan !== null
  }

  stopPointing (): void {
    this.pointSpan = null
  }

  /** Insert text at the caret (e.g. from the function picker). */
  insert (text: string): void {
    this.el.setRangeText(text, this.el.selectionStart, this.el.selectionEnd, 'end')
    this.pointSpan = null
    this.afterChange()
  }

  // Autocomplete -------------------------------------------------------------

  get suggesting (): boolean {
    return this.suggestions.length > 0
  }

  moveSuggestion (delta: number): void {
    if (!this.suggestions.length) return
    this.suggestIndex = (this.suggestIndex + delta + this.suggestions.length) % this.suggestions.length
    this.navigated = true
    this.renderPopup()
  }

  acceptSuggestion (): boolean {
    const name = this.suggestions[this.suggestIndex]
    if (!name) return false
    this.accept(name)
    return true
  }

  hidePopup (): void {
    this.suggestions = []
    this.navigated = false
    this.popup.style.display = 'none'
    this.el.removeAttribute('aria-activedescendant')
  }

  private accept (name: string): void {
    const pos = this.el.selectionStart
    const before = this.el.value.slice(0, pos)
    const m = /[A-Za-z][A-Za-z0-9.]*$/.exec(before)
    const start = m ? pos - m[0].length : pos
    const after = this.el.value.slice(pos)
    const insertParen = !after.startsWith('(')
    this.el.setRangeText(name + (insertParen ? '(' : ''), start, pos, 'end')
    if (!insertParen) this.el.setSelectionRange(start + name.length + 1, start + name.length + 1)
    this.pointSpan = null
    this.hidePopup()
    this.afterChange()
  }

  private afterChange (): void {
    this.autosize()
    this.updateSuggestions()
    this.updateHint()
    this.onInput?.(this.el.value)
  }

  private autosize (): void {
    if (!this.active) return
    const lines = this.el.value.split('\n')
    let w = 0
    for (const line of lines) w = Math.max(w, this.measureCtx.measureText(line).width)
    const width = Math.min(this.maxWidth, Math.max(this.minWidth, Math.ceil(w) + 16))
    this.el.style.width = width + 'px'
  }

  private updateSuggestions (): void {
    const pos = this.el.selectionStart
    if (!this.active || !this.isFormula || pos !== this.el.selectionEnd) return this.hidePopup()
    const before = this.el.value.slice(0, pos)
    if (inString(before)) return this.hidePopup()
    const m = /(?:^=|[=(,+\-*/^&<>;\s])([A-Za-z][A-Za-z0-9.]*)$/.exec(before)
    if (!m) return this.hidePopup()
    const prefix = m[1].toUpperCase()
    const matches = FUNCTION_NAMES.filter((n) => n.startsWith(prefix))
    if (!matches.length || (matches.length === 1 && matches[0] === prefix)) return this.hidePopup()
    const changed = matches.join() !== this.suggestions.join()
    this.suggestions = matches.slice(0, 50)
    if (changed) {
      this.suggestIndex = 0
      this.navigated = false
      if (this.suggestions.length !== this.lastSuggestCount) {
        this.lastSuggestCount = this.suggestions.length
        this.onAnnounce?.(`${this.suggestions.length} function${this.suggestions.length === 1 ? '' : 's'} suggested. Arrow keys to choose, Tab to insert.`)
      }
    }
    this.renderPopup()
  }

  private renderPopup (): void {
    const items = this.suggestions
    this.popup.replaceChildren(
      ...items.map((name, i) => {
        const div = document.createElement('div')
        div.className = 'ac-item' + (i === this.suggestIndex ? ' selected' : '')
        div.id = `fn-opt-${name.replace(/\./g, '-')}`
        div.dataset.fn = name
        div.setAttribute('role', 'option')
        div.setAttribute('aria-selected', String(i === this.suggestIndex))
        const strong = document.createElement('strong')
        strong.textContent = name
        const desc = document.createElement('span')
        desc.textContent = FUNCTIONS[name].desc
        div.append(strong, desc)
        return div
      })
    )
    this.popup.style.display = 'block'
    this.popup.style.left = this.el.style.left
    this.popup.style.top = parseFloat(this.el.style.top) + parseFloat(this.el.style.height) + 2 + 'px'
    const selected = this.popup.querySelector<HTMLElement>('.selected')
    selected?.scrollIntoView({ block: 'nearest' })
    if (selected) this.el.setAttribute('aria-activedescendant', selected.id)
    this.hint.style.display = 'none'
  }

  /** Show the signature of the function call the caret is inside. */
  private updateHint (): void {
    if (!this.active || !this.isFormula || this.suggestions.length) {
      this.hint.style.display = 'none'
      return
    }
    const before = this.el.value.slice(0, this.el.selectionStart)
    const call = innermostCall(before)
    const def = call && FUNCTIONS[call.name]
    if (!call || !def) {
      this.hint.style.display = 'none'
      return
    }
    const m = /^([^(]+)\((.*)\)$/.exec(def.sig)
    const params = m ? m[2].split(/,\s*/) : []
    this.hint.replaceChildren()
    this.hint.append(document.createTextNode((m ? m[1] : call.name) + '('))
    params.forEach((p, i) => {
      if (i) this.hint.append(document.createTextNode(', '))
      // Repeating parameters ("…") highlight the last named one.
      const active = i === Math.min(call.arg, params.length - 1) || (i === params.length - 2 && params[params.length - 1] === '…' && call.arg >= i)
      if (active && p !== '…') {
        const b = document.createElement('b')
        b.textContent = p
        this.hint.append(b)
        // Announce when the caret moves into a different argument.
        const key = `${call.name}:${p}`
        if (key !== this.lastHint) {
          this.lastHint = key
          this.onAnnounce?.(`${call.name}, ${p.replace(/^\[(.*)\]$/, '$1, optional')}`)
        }
      } else this.hint.append(document.createTextNode(p))
    })
    this.hint.append(document.createTextNode(')'))
    const desc = document.createElement('div')
    desc.className = 'sig-desc'
    desc.textContent = def.desc
    this.hint.append(desc)
    this.hint.style.display = 'block'
    this.hint.style.left = this.el.style.left
    this.hint.style.top = parseFloat(this.el.style.top) + parseFloat(this.el.style.height) + 2 + 'px'
  }
}

function inString (text: string): boolean {
  let n = 0
  for (const ch of text) if (ch === '"') n++
  return n % 2 === 1
}

/** Find the innermost unclosed function call and which argument the caret is in. */
export function innermostCall (text: string): { name: string; arg: number } | null {
  const stack: { name: string; arg: number }[] = []
  let inStr = false
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (ch === '"') inStr = !inStr
    if (inStr) continue
    if (ch === '(') {
      const m = /([A-Za-z][A-Za-z0-9.]*)\s*$/.exec(text.slice(0, i))
      stack.push({ name: m ? m[1].toUpperCase() : '', arg: 0 })
    } else if (ch === ')') stack.pop()
    else if (ch === ',' && stack.length) stack[stack.length - 1].arg++
  }
  for (let i = stack.length - 1; i >= 0; i--) if (stack[i].name) return stack[i]
  return null
}

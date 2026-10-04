import { html } from 'lit'

/**
 * The static markup of <film-sheet>. It has no bindings: the controller
 * drives everything imperatively, so Lit renders this once and never
 * touches it again.
 */
export const sheetTemplate = html`
    <div class="sheet">
      <p id="grid-help" class="sr-only">Arrow keys move between cells; Shift plus arrows selects. Type to replace a cell, F2 to edit it. F6 moves between the menu bar, toolbar, formula bar, grid and sheet tabs; F10 opens the menus.</p>
      <div class="announcer sr-only" role="status" aria-live="polite" aria-atomic="true"></div>
      <film-menu-bar class="menubar">
        <film-menu-bar-item label="File">
          <film-menu>
            <film-menu-item value="file.new">New workbook<film-kbd>⌃N</film-kbd></film-menu-item>
            <film-menu-item value="file.open">Open…<film-kbd>⌃O</film-kbd></film-menu-item>
            <film-menu-item value="file.save">Save as .json<film-kbd>⌃S</film-kbd></film-menu-item>
            <film-menu-item value="file.exportCsv">Export sheet as CSV</film-menu-item>
          </film-menu>
        </film-menu-bar-item>
        <film-menu-bar-item label="Edit">
          <film-menu>
            <film-menu-item value="edit.undo">Undo<film-kbd>⌃Z</film-kbd></film-menu-item>
            <film-menu-item value="edit.redo">Redo<film-kbd>⌃Y</film-kbd></film-menu-item>
            <film-menu-item value="edit.cut">Cut<film-kbd>⌃X</film-kbd></film-menu-item>
            <film-menu-item value="edit.copy">Copy<film-kbd>⌃C</film-kbd></film-menu-item>
            <film-menu-item value="edit.paste">Paste<film-kbd>⌃V</film-kbd></film-menu-item>
            <film-menu-item value="edit.pasteValues">Paste values only<film-kbd>⌃⇧V</film-kbd></film-menu-item>
            <film-menu-item value="edit.fillDown">Fill down<film-kbd>⌃D</film-kbd></film-menu-item>
            <film-menu-item value="edit.fillRight">Fill right<film-kbd>⌃R</film-kbd></film-menu-item>
            <film-menu-item value="edit.seriesDown">Fill series down</film-menu-item>
            <film-menu-item value="edit.seriesRight">Fill series right</film-menu-item>
            <film-menu-item value="edit.find">Find and replace…<film-kbd>⌃F</film-kbd></film-menu-item>
            <film-menu-item>Clear
              <film-menu slot="submenu">
                <film-menu-item value="edit.clear">Contents<film-kbd>Del</film-kbd></film-menu-item>
                <film-menu-item value="edit.clearFormats">Formats</film-menu-item>
                <film-menu-item value="edit.clearAll">All</film-menu-item>
              </film-menu>
            </film-menu-item>
            <film-menu-item>Delete
              <film-menu slot="submenu">
                <film-menu-item value="edit.deleteRows">Selected rows</film-menu-item>
                <film-menu-item value="edit.deleteCols">Selected columns</film-menu-item>
              </film-menu>
            </film-menu-item>
          </film-menu>
        </film-menu-bar-item>
        <film-menu-bar-item label="View">
          <film-menu>
            <film-menu-item value="view.formulas">Show formulas<film-kbd>⌃\`</film-kbd></film-menu-item>
            <film-menu-item value="view.goto">Go to cell…<film-kbd>⌃G</film-kbd></film-menu-item>
          </film-menu>
        </film-menu-bar-item>
        <film-menu-bar-item label="Insert">
          <film-menu>
            <film-menu-item value="insert.rowAbove">Row above</film-menu-item>
            <film-menu-item value="insert.rowBelow">Row below</film-menu-item>
            <film-menu-item value="insert.colLeft">Column left</film-menu-item>
            <film-menu-item value="insert.colRight">Column right</film-menu-item>
            <film-menu-item value="insert.sheet">Sheet<film-kbd>⇧F11</film-kbd></film-menu-item>
            <film-menu-item value="insert.sum">AutoSum<film-kbd>⌥=</film-kbd></film-menu-item>
            <film-menu-item value="help.functions">Function…</film-menu-item>
          </film-menu>
        </film-menu-bar-item>
        <film-menu-bar-item label="Format">
          <film-menu>
            <film-menu-item value="format.bold">Bold<film-kbd>⌃B</film-kbd></film-menu-item>
            <film-menu-item value="format.italic">Italic<film-kbd>⌃I</film-kbd></film-menu-item>
            <film-menu-item value="format.underline">Underline<film-kbd>⌃U</film-kbd></film-menu-item>
            <film-menu-item value="format.strike">Strikethrough<film-kbd>⌃5</film-kbd></film-menu-item>
            <film-menu-item>Number
              <film-menu slot="submenu">
                <film-menu-item value="format.num.general">General</film-menu-item>
                <film-menu-item value="format.num.number">Number <span class="eg">1,234.56</span></film-menu-item>
                <film-menu-item value="format.num.currency">Currency <span class="eg">$1,234.56</span></film-menu-item>
                <film-menu-item value="format.num.percent">Percent <span class="eg">12.34%</span></film-menu-item>
                <film-menu-item value="format.num.scientific">Scientific <span class="eg">1.23E+03</span></film-menu-item>
                <film-menu-item value="format.num.date">Date <span class="eg">2026-10-03</span></film-menu-item>
                <film-menu-item value="format.num.datetime">Date time <span class="eg">2026-10-03 14:30</span></film-menu-item>
                <film-menu-item value="format.num.time">Time <span class="eg">14:30:00</span></film-menu-item>
                <film-menu-item value="format.num.text">Plain text</film-menu-item>
              </film-menu>
            </film-menu-item>
            <film-menu-item>Align
              <film-menu slot="submenu">
                <film-menu-item value="format.alignLeft">Left</film-menu-item>
                <film-menu-item value="format.alignCenter">Center</film-menu-item>
                <film-menu-item value="format.alignRight">Right</film-menu-item>
              </film-menu>
            </film-menu-item>
            <film-menu-item value="format.colWidth">Column width…</film-menu-item>
            <film-menu-item value="format.rowHeight">Row height…</film-menu-item>
            <film-menu-item value="format.autofit">Autofit column width</film-menu-item>
            <film-menu-item value="format.clear">Clear formatting<film-kbd>⌃\\</film-kbd></film-menu-item>
          </film-menu>
        </film-menu-bar-item>
        <film-menu-bar-item label="Data">
          <film-menu>
            <film-menu-item value="data.sortAsc">Sort range A → Z</film-menu-item>
            <film-menu-item value="data.sortDesc">Sort range Z → A</film-menu-item>
          </film-menu>
        </film-menu-bar-item>
        <film-menu-bar-item label="Sheet">
          <film-menu>
            <film-menu-item value="insert.sheet">New sheet</film-menu-item>
            <film-menu-item value="sheet.rename">Rename…</film-menu-item>
            <film-menu-item value="sheet.duplicate">Duplicate</film-menu-item>
            <film-menu-item value="sheet.moveLeft">Move left</film-menu-item>
            <film-menu-item value="sheet.moveRight">Move right</film-menu-item>
            <film-menu-item value="sheet.delete">Delete</film-menu-item>
          </film-menu>
        </film-menu-bar-item>
        <film-menu-bar-item label="Help">
          <film-menu>
            <film-menu-item value="help.shortcuts">Keyboard shortcuts<film-kbd>⌃/</film-kbd></film-menu-item>
            <film-menu-item value="help.functions">Function reference</film-menu-item>
          </film-menu>
        </film-menu-bar-item>
      </film-menu-bar>

      <div class="toolbar" role="toolbar" aria-label="Formatting">
        <film-button-group label="History">
          <film-icon-button label="Undo" data-cmd="edit.undo"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M9 14 4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-4"/></svg></film-icon-button>
          <film-icon-button label="Redo" data-cmd="edit.redo"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="m15 14 5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h4"/></svg></film-icon-button>
        </film-button-group>
        <film-divider vertical space="s-3"></film-divider>
        <film-select class="format-select" aria-label="Number format" value="general">
          <film-select-option value="general">General</film-select-option>
          <film-select-option value="number">Number</film-select-option>
          <film-select-option value="currency">Currency</film-select-option>
          <film-select-option value="percent">Percent</film-select-option>
          <film-select-option value="scientific">Scientific</film-select-option>
          <film-select-option value="date">Date</film-select-option>
          <film-select-option value="datetime">Date time</film-select-option>
          <film-select-option value="time">Time</film-select-option>
          <film-select-option value="text">Plain text</film-select-option>
          <film-select-option value="custom" disabled>Custom</film-select-option>
        </film-select>
        <film-button-group label="Number">
          <film-icon-button label="Format as currency" data-cmd="format.num.currency"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M12 3v18"/><path d="M17 7H9.5a3 3 0 0 0 0 6h5a3 3 0 0 1 0 6H6"/></svg></film-icon-button>
          <film-icon-button label="Format as percent" data-cmd="format.num.percent"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M19 5 5 19"/><circle cx="6.5" cy="6.5" r="2.5"/><circle cx="17.5" cy="17.5" r="2.5"/></svg></film-icon-button>
          <film-icon-button label="Decrease decimal places" data-cmd="format.decDec"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><text x="2" y="16" font-size="10" stroke="none" fill="currentColor">.0</text><path d="m20 14-4 0m0 0 2-2m-2 2 2 2"/></svg></film-icon-button>
          <film-icon-button label="Increase decimal places" data-cmd="format.decInc"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><text x="1" y="16" font-size="9" stroke="none" fill="currentColor">.00</text><path d="m17 14 5 0m0 0-2-2m2 2-2 2"/></svg></film-icon-button>
        </film-button-group>
        <film-divider vertical space="s-3"></film-divider>
        <film-button-group label="Text style">
          <film-icon-button label="Bold" data-cmd="format.bold" data-toggle="bold"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M7 5h6a3.5 3.5 0 0 1 0 7H7zM7 12h7a3.5 3.5 0 0 1 0 7H7z"/></svg></film-icon-button>
          <film-icon-button label="Italic" data-cmd="format.italic" data-toggle="italic"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M14 5h5M5 19h5M15 5 9 19"/></svg></film-icon-button>
          <film-icon-button label="Underline" data-cmd="format.underline" data-toggle="underline"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M7 4v7a5 5 0 0 0 10 0V4M5 20h14"/></svg></film-icon-button>
          <film-icon-button label="Strikethrough" data-cmd="format.strike" data-toggle="strike"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M16 6a4 4 0 0 0-4-2c-3 0-4.5 2-4 4M4 12h16M8 18a4 4 0 0 0 4 2c3 0 4.5-2 4-4.5"/></svg></film-icon-button>
        </film-button-group>
        <film-dropdown class="color-dropdown" data-target="color">
          <film-icon-button slot="trigger" label="Text colour"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="m6 16 6-12 6 12M8 12h8"/><rect class="swatch-bar" x="3" y="19" width="18" height="3" stroke="none"/></svg></film-icon-button>
          <div class="swatches" role="group" aria-label="Text colour"></div>
        </film-dropdown>
        <film-dropdown class="color-dropdown" data-target="fill">
          <film-icon-button slot="trigger" label="Fill colour"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="m5 11 7-7 7 7-7 7z"/><path d="M19 13s2 2.5 2 3.5a2 2 0 0 1-4 0c0-1 2-3.5 2-3.5"/><rect class="swatch-bar" x="3" y="19" width="18" height="3" stroke="none"/></svg></film-icon-button>
          <div class="swatches" role="group" aria-label="Fill colour"></div>
        </film-dropdown>
        <film-divider vertical space="s-3"></film-divider>
        <film-button-group label="Alignment">
          <film-icon-button label="Align left" data-cmd="format.alignLeft" data-align="left"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M4 6h16M4 12h10M4 18h13"/></svg></film-icon-button>
          <film-icon-button label="Align center" data-cmd="format.alignCenter" data-align="center"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M4 6h16M7 12h10M5 18h14"/></svg></film-icon-button>
          <film-icon-button label="Align right" data-cmd="format.alignRight" data-align="right"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M4 6h16M10 12h10M7 18h13"/></svg></film-icon-button>
        </film-button-group>
        <film-divider vertical space="s-3"></film-divider>
        <film-button-group label="Rows and columns">
          <film-icon-button label="Insert row above" data-cmd="insert.rowAbove"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><rect x="3" y="12" width="18" height="8" rx="1"/><path d="M12 3v6M9 6h6"/></svg></film-icon-button>
          <film-icon-button label="Insert column left" data-cmd="insert.colLeft"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><rect x="12" y="3" width="8" height="18" rx="1"/><path d="M3 12h6M6 9v6"/></svg></film-icon-button>
          <film-icon-button label="Sort A to Z" data-cmd="data.sortAsc"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M7 4v16m0 0-3-3m3 3 3-3"/><text x="13" y="11" font-size="8" stroke="none" fill="currentColor">A</text><text x="13" y="20" font-size="8" stroke="none" fill="currentColor">Z</text></svg></film-icon-button>
        </film-button-group>
        <film-tooltip content="AutoSum (⌥=)">
          <film-button class="sum-button" variant="neutral" size="small" data-cmd="insert.sum"><span aria-hidden="true">Σ</span> Sum</film-button>
        </film-tooltip>
      </div>

      <div class="formula-bar">
        <film-input class="name-box" aria-label="Name box" value="A1"></film-input>
        <span class="fx" aria-hidden="true">fx</span>
        <film-input class="formula-input" aria-label="Formula bar" placeholder=""></film-input>
      </div>

      <main class="grid-host"></main>

      <footer class="sheet-bar">
        <film-icon-button label="Add sheet" data-cmd="insert.sheet"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M12 5v14M5 12h14"/></svg></film-icon-button>
        <film-tabs class="sheet-tabs" aria-label="Sheets"></film-tabs>
        <div class="status">
          <span class="status-mode">Ready</span>
          <span class="status-stats"></span>
          <span class="status-perf"></span>
          <span class="save-state"></span>
        </div>
      </footer>
    </div>

    <div id="context-menu" class="context-menu" popover="auto">
      <film-menu>
        <film-menu-item value="edit.cut">Cut<film-kbd>⌃X</film-kbd></film-menu-item>
        <film-menu-item value="edit.copy">Copy<film-kbd>⌃C</film-kbd></film-menu-item>
        <film-menu-item value="edit.paste">Paste<film-kbd>⌃V</film-kbd></film-menu-item>
        <film-menu-item value="edit.pasteValues">Paste values only</film-menu-item>
        <film-menu-item value="insert.rowAbove">Insert row above</film-menu-item>
        <film-menu-item value="insert.rowBelow">Insert row below</film-menu-item>
        <film-menu-item value="insert.colLeft">Insert column left</film-menu-item>
        <film-menu-item value="insert.colRight">Insert column right</film-menu-item>
        <film-menu-item value="edit.deleteRows">Delete rows</film-menu-item>
        <film-menu-item value="edit.deleteCols">Delete columns</film-menu-item>
        <film-menu-item value="edit.clear">Clear contents</film-menu-item>
        <film-menu-item value="data.sortAsc">Sort A → Z</film-menu-item>
        <film-menu-item value="data.sortDesc">Sort Z → A</film-menu-item>
      </film-menu>
    </div>

    <div id="tab-menu" class="context-menu" popover="auto">
      <film-menu>
        <film-menu-item value="sheet.rename">Rename…</film-menu-item>
        <film-menu-item value="sheet.duplicate">Duplicate</film-menu-item>
        <film-menu-item value="sheet.moveLeft">Move left</film-menu-item>
        <film-menu-item value="sheet.moveRight">Move right</film-menu-item>
        <film-menu-item value="sheet.delete">Delete</film-menu-item>
      </film-menu>
    </div>

    <film-dialog id="prompt-dialog" label="Rename sheet">
      <form class="dialog-form" method="dialog">
        <film-input class="prompt-input" label="Name"></film-input>
      </form>
      <film-button slot="footer" variant="neutral" class="prompt-cancel">Cancel</film-button>
      <film-button slot="footer" class="prompt-ok">OK</film-button>
    </film-dialog>

    <film-dialog id="find-dialog" label="Find and replace">
      <film-stack space="s-1">
        <film-input class="find-input" label="Find"></film-input>
        <film-input class="replace-input" label="Replace with"></film-input>
        <film-cluster space="s-1">
          <film-checkbox class="find-case">Match case</film-checkbox>
          <film-checkbox class="find-whole">Match entire cell</film-checkbox>
          <film-checkbox class="find-formulas">Search in formulas</film-checkbox>
        </film-cluster>
        <film-text class="find-status" size="s-1" role="status"></film-text>
      </film-stack>
      <film-button slot="footer" variant="neutral" class="find-replace-all">Replace all</film-button>
      <film-button slot="footer" variant="neutral" class="find-replace">Replace</film-button>
      <film-button slot="footer" class="find-next">Find next</film-button>
    </film-dialog>


    <film-dialog id="shortcuts-dialog" label="Keyboard shortcuts">
      <div class="shortcuts"></div>
    </film-dialog>

    <film-dialog id="functions-dialog" label="Function reference">
      <film-stack space="s-1">
        <film-search class="fn-search" label="Search functions" placeholder="Search…"></film-search>
        <div class="fn-list"></div>
      </film-stack>
    </film-dialog>

    <input type="file" class="file-input" accept=".csv,.tsv,.txt,.json" hidden />
`

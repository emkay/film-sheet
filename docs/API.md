# API

## `<film-sheet>`

```js
import '@mk/film/css/themes/default/index.css' // Film's theme tokens (required)
import '@mk/film-sheet'                         // registers <film-sheet> and the Film elements it uses
```

```html
<film-sheet style="height: 480px"></film-sheet>
```

The component needs a height, either from your CSS or from its 32em default.

### Attributes and properties

| Attribute | Property | Default | |
|---|---|---|---|
| `readonly` | `readonly` | `false` | Blocks every edit. Viewing, selecting, copying, find and export still work. |
| `storage-key` | `storageKey` | `""` | localStorage key to autosave to and restore from. Empty means no persistence. |
| `chrome` | `chrome` | `"full"` | `full` shows everything, `minimal` shows formula bar + grid + tabs, `none` shows only the grid. |
| — | `filename` | `"workbook"` | Base name for *Save as .json* and *Export CSV*. |
| — | `data` | | The workbook as JSON; assigning replaces it. |
| — | `workbook` | | The live engine `Workbook`. |

### Methods

- `load(json)` replaces the workbook. Invalid data throws, and the current
  workbook is left unchanged.
- `toJSON()` returns the workbook as plain JSON, for saving or for passing to `load()`.
- `focus()` puts keyboard focus on the active cell.
- `refreshTheme()` re-reads theme colours. Call it after changing Film tokens on an
  ancestor yourself; changes to `data-theme` / `data-film-theme` on `<html>`
  and colour-scheme changes are picked up automatically.

### Events

Both events bubble and are composed.

| Event | `detail` | When |
|---|---|---|
| `film-sheet-change` | `null` | After any change to cells, formats or sheets |
| `film-sheet-select` | `{ sheet, address, range }` | Selection or active sheet changed |

```js
const sheet = document.querySelector('film-sheet')
sheet.load({ version: 1, sheets: [{ name: 'Budget', cells: [[0, 0, 'Rent'], [0, 1, '1200'], [1, 1, '=B1*12']] }] })
sheet.addEventListener('film-sheet-change', () => save(sheet.toJSON()))
```

### Styling

- **Colours:** the component follows the page's Film theme, so set
  `data-film-theme` / `data-theme` on `<html>`.
- **Size:** everything is sized from `--film-sheet-size` (default `14px`), not
  from the page's root font size:

  ```css
  film-sheet { --film-sheet-size: 13px; }
  ```

### Behaviour on the page

Several sheets can share a page; each keeps its own data, focus, keyboard and
clipboard handling. The component never changes the document title, `<html>`
attributes or globals. It only uses localStorage when `storage-key` is set,
and only grabs focus when you call `focus()`. Removing it from the DOM tears
everything down, and re-adding it restores its data.

## `@mk/film-sheet/engine`

The formula engine on its own, with no DOM. It runs in browsers, Node, Deno,
Bun and workers.

```js
import { Workbook, registerFunction, toNumber, isError } from '@mk/film-sheet/engine'

const wb = new Workbook()
const id = wb.sheets[0].id
wb.setCell(id, 0, 0, '2')       // A1
wb.setCell(id, 1, 0, '=A1^10')  // A2
wb.getValue(id, 1, 0)           // 1024

// Batch many edits and recalculate once
wb.batch(() => { for (let r = 0; r < 1000; r++) wb.setCell(id, r, 1, `=A1*${r}`) })

// Custom functions (register before loading workbooks that use them)
registerFunction('DOUBLE', {
  min: 1, max: 1,
  fn: ([x]) => { const n = toNumber(x); return isError(n) ? n : n * 2 },
  sig: 'DOUBLE(number)', desc: 'Twice the number'
})
```

The main parts of the engine API:

- **`Workbook`:**
  - Sheets: `sheets`, `addSheet`, `removeSheet`, `renameSheet`, `moveSheet`.
  - Cells: `setCell`, `applyPatches` (returns undo patches), `getCell`, `getValue`, `getDisplay`.
  - Structure: `insertRows` / `deleteRows` / `insertCols` / `deleteCols`, `sortPatches`.
  - Other: `batch`, `onChange`, `toJSON` / `load`, `lastRecalc`.
- **Helpers:** address parsing and formatting, `formatValue` / `parseInput` (number
  formats), `Errors` and value coercions, `registerFunction`, `shiftFormula`.

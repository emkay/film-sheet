# @mk/film-sheet

An Excel-like spreadsheet web component with a fast, incremental formula
engine, built on the [Film](https://www.npmjs.com/package/@mk/film) design
language.

- `<film-sheet>`: a complete spreadsheet UI with menus, toolbar, formula bar,
  sheet tabs, undo/redo, clipboard, fill handle, find and replace, and full
  keyboard and screen-reader support. The grid is canvas-rendered and
  virtualised up to Excel's 1,048,576 × 16,384 cells.
- `@mk/film-sheet/engine`: the formula engine on its own, with no DOM. It runs
  in browsers, Node, Deno, Bun and workers. It has about 140 functions, incremental
  recalculation, cross-sheet references and structural edits.

## Install

```sh
npm i @mk/film-sheet @mk/film lit
```

`@mk/film` and `lit` are peer dependencies, so your app shares one copy of each.

## Using the component

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
  `data-film-theme` / `data-theme` on `<html>` (see Film's docs).
- **Size:** everything is sized from one value, `--film-sheet-size` (default
  `14px`), not from the page's root font size:

  ```css
  film-sheet { --film-sheet-size: 13px; }
  ```

### Behaviour on the page

Several sheets can share a page; each keeps its own data, focus, keyboard and
clipboard handling. The component never changes the document title, `<html>`
attributes or globals. It only uses localStorage when `storage-key` is set,
and only grabs focus when you call `focus()`. Removing it from the DOM tears
everything down, and re-adding it restores its data.

## Using the engine

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

## Development

```sh
npm install
npx playwright install chromium  # once, for the browser tests
npm run dev            # demo app at http://localhost:5173
npm test               # unit tests (engine + logic)
npm run test:e2e       # browser tests (Playwright + Chromium)
npm run lint           # oxlint, including type-aware rules
npm run check          # typecheck + lint + coverage + e2e
npm run build:lib      # the package, into dist/
npm run build          # the demo site, into dist-demo/
```

### Layout

- `src/engine/`: the formula engine (tokenizer, parser, closure compiler,
  functions, workbook with an incremental dependency graph).
- `src/ui/`: the component. `film-sheet.ts` is the element, `controller.ts`
  holds commands, keyboard, clipboard and undo, `grid.ts` is the canvas grid
  and its accessible active-cell proxy, and `editor.ts` is the in-cell editor.
- `demo/`: the demo page. It hosts one `<film-sheet>` and adds page-level things
  the component deliberately doesn't own: the theme picker, document title,
  sample data and a performance test.
- `test/`: unit tests. `e2e/`: browser tests, including `e2e/embed.html`, a
  page with two embedded sheets used to test the component API.

### Tooling notes

- **Linting:** this project uses TypeScript 7 (the native compiler), which has no
  JavaScript API, so linting uses [oxlint](https://oxc.rs) instead of
  typescript-eslint. Its type-aware mode runs on the same native compiler.
- **Coverage:** thresholds cover everything that runs under Node. The canvas and DOM
  code is covered by the Playwright tests instead.
- **Type declarations:** `tsc -p tsconfig.build.json` emits them with
  `stripInternal`, so `@internal` members aren't part of the public types.
  Relative imports use `.js` extensions so the declarations also work for
  consumers on `moduleResolution: node16`.

### Performance

- An edit only re-evaluates the cells that depend on it.
- Evaluation uses an explicit stack, so dependency chains of any length work.
- Whole-column ranges only visit stored cells.

The demo's *Performance test…* builds up to 2M cells. In Node, a 1M-cell
sheet (750k formulas) loads in about 1.7 s, and editing an input that feeds a
250k-deep chain recalculates in about 130 ms.

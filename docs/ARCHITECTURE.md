# Architecture

FilmSheet has two layers: a synchronous formula engine with no DOM
(`src/engine/`) and a web component that renders and edits it
(`src/ui/`). Performance is a primary goal throughout, so most of the
decisions below come down to doing less work per edit and per frame.

```
src/
  index.ts            package entry: registers <film-sheet>
  engine/             formula engine (no DOM) — package entry engine/index.ts
    tokenizer.ts      formula text → tokens (refs are single tokens, so rewriting is exact)
    parser.ts         tokens → AST (Pratt parser, Excel precedence)
    compiler.ts       AST → JS closures, plus the formula's precedents
    functions.ts      built-in functions + registerFunction()
    workbook.ts       sheets, cells, dependency graph, recalculation, structure, JSON
    refs.ts           reference rewriting: copy/fill, insert/delete, renames
    format.ts         number/date formats and typed-input parsing
    values.ts         value model, errors, coercions
  ui/
    film-sheet.ts     the <film-sheet> element: properties, events, lifecycle
    controller.ts     commands, keyboard, clipboard, undo, persistence, chrome wiring
    grid.ts           canvas grid, scrolling, hit testing, accessible active-cell proxy
    editor.ts         in-cell editor, autocomplete, signature hints, reference pointing
    template.ts       the element's static markup
    styles.css        the element's styles (inlined at build time)
    axis.ts fill.ts history.ts io.ts   pure helpers
demo/                 demo page hosting one <film-sheet>
test/                 unit tests (Vitest, Node)
e2e/                  browser tests (Playwright) + embed.html test page
scripts/              package verification
```

## Engine

### Storage

A `Sheet` stores cells sparsely as `Map<column, Map<row, Cell>>`. Column-major
maps make range scans (`SUM(A:A)`) walk only stored cells, and give cheap
whole-column iteration. `maxRow`/`maxCol` bound scans of whole-row and
whole-column references; they only grow until a rebuild.

### Formulas

Typed text is parsed once and compiled into a tree of closures. Each closure
takes the evaluation context and returns a value, and references compile to
direct reads, so evaluating a formula never revisits its AST. Small arities
get specialised call paths to avoid per-call array allocation. Compiling also
yields the formula's precedents (cells and ranges it reads).

Range arguments are passed as lazy `RangeValue` objects (`get`, `forEach`,
`forEachOrdered`) rather than arrays, so aggregates stream over stored cells
and lookups stop early.

### Dependency graph

- Single-cell precedents go in `cellDeps: Map<cellKey, Set<Cell>>`, where
  `cellKey` packs the sheet id, row and column into one number.
- Range precedents are listeners in per-column buckets. Ranges wider than 64
  columns go into a per-sheet "wide" list instead.

An edit records its position. At recalc, `markDependents` runs one breadth
pass over all changed positions and marks dependents dirty. Each bucket's
listeners are copied into a per-call worklist the first time the bucket is
hit, and removed from it once marked, so a bucket costs O(listeners) per
recalc rather than per visited cell.

### Recalculation

Only dirty cells are evaluated, using an explicit stack (no recursion), so
dependency chains of any length are fine:

1. Pop a cell and evaluate its formula.
2. If it reads a cell that is still dirty, that read throws a `PENDING`
   sentinel. The formula is abandoned: no wasted work, and no IF branch runs
   on a placeholder value. The cell stays on the stack with its missing
   inputs pushed above it.
3. If the missing input is itself being evaluated, that's a cycle, and the
   cell becomes `#CIRC!`.

Volatile functions (`RAND`, `NOW`, `INDIRECT`, `OFFSET`) are re-marked dirty
on every recalc.

### Structural edits

Inserting or deleting rows or columns and renaming sheets rewrite formula
*text* through the tokenizer, then `rebuild()` recompiles everything. This is
simple and correct, but O(formulas). Shifting the dependency graph in place
is the known next step if structural edits on very large sheets need to be
faster.

### Persistence

`toJSON()`/`load()` use a compact JSON format (`[row, col, raw, style?]` per
cell). `load()` validates first, so bad input never clobbers a workbook. Sheet
ids survive round trips, and the UI's undo history relies on that.

## Component

### Element and controller

`<film-sheet>` is a Lit element with a static template and no bindings, so
Lit renders it once. A `SheetController` drives it imperatively. Every DOM
lookup, listener and focus check is scoped to the element's shadow root, which
is what lets several sheets share a page without interfering. The controller
is disposed on disconnect, and the workbook is stashed and restored if the
element is re-attached.

Page-level concerns (document title, theme selection, sample data) live in
the demo, not the component. The component follows Film's theme tokens and
builds its own modular size scale from `--film-sheet-size`, so it doesn't
depend on the host page's root font size.

### Grid

The grid is a single `<canvas>` under a transparent native scroller (for real
scrollbars and pointer input). Only visible cells are drawn, and formatted text
and text widths are cached. Content taller than browsers allow
(1,048,576 rows) maps onto a capped scroll area with a scale factor. Row and
column sizes are stored sparsely (`Axis`), with O(log k) offset and index
lookups.

### Accessibility

A canvas has no DOM for assistive tech, so the grid mirrors the active cell into a tiny
ARIA grid fragment: a header row plus one row containing its row header and
the cell. Two copies alternate, so every move is a real focus change that
screen readers announce, and the focused proxy sits over the cell for
magnifiers. A polite live region announces range selections, mode changes and
command results. F6 cycles through the page regions, and the toolbar uses a
roving tabindex.

### Undo

History entries are cell patches (`applyPatches` returns the inverse),
whole-workbook snapshots for structural edits (only the "before" state is
stored until first undo), row/column size changes, and groups.

## Testing

- **Unit tests** (`test/`, Vitest in Node): the engine and pure UI logic. A
  case table exercises every built-in function, and a guard test fails if a
  function lacks a case. Coverage thresholds apply to everything that runs in
  Node.
- **Browser tests** (`e2e/`, Playwright + Chromium): drive the real
  component with keyboard, mouse and clipboard. They assert what users and
  screen readers perceive (the focused cell's accessible name, the formula
  bar, the live region) and run an axe scan. `e2e/embed.html` hosts two
  sheets to test isolation and the public API.
- **Package verification** (`scripts/verify-package.mjs`): packs the
  tarball, installs it into a throwaway project, runs the engine in Node and
  typechecks both entry points with TypeScript 5 (`bundler` and `node16`).

## Tooling

- **TypeScript 7** (the native compiler) for typechecking and declaration
  emit. `stripInternal` keeps `@internal` members out of the public types.
  Relative imports carry `.js` extensions so the declarations also work under
  `moduleResolution: node16`.
- **oxlint** for linting, including type-aware rules. typescript-eslint needs
  TypeScript's JavaScript API, which TypeScript 7 doesn't provide.
- **Vite** builds the demo (`vite build`) and the library
  (`vite build --mode lib`, ES modules, preserved module structure). `lit`
  and `@mk/film` are external peer dependencies.

## Performance

These numbers are for one laptop and are indicative.

| Workload | Time |
|---|---|
| Load 1M cells (750k formulas, a 250k-deep chain), Node | ~1.7 s |
| Edit one input feeding a 250k-deep chain (250k cells re-evaluated), Node | ~130 ms |
| `SUM(A:A)` over 250k stored cells | ~5 ms |

Run the demo's *Performance test…* to measure in a browser.

## Future: large data

Datasets larger than an in-memory sheet comfortably holds (a few million
cells) are expected to come through a data-source plugin layer. Each plugin
would have its own loading strategy (streaming, windowed/virtual ranges,
external query engines such as DuckDB-WASM), lazily loaded so the core stays
small and synchronous. Keep cell storage behind the `Sheet` API so such a
layer can be added without changing formulas or the grid.

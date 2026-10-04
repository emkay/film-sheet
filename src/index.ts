/**
 * @mk/film-sheet — an Excel-like spreadsheet web component built on Film.
 *
 * Importing this module registers <film-sheet>. The page also needs Film's
 * components and theme:
 *
 * ```ts
 * import '@mk/film'
 * import '@mk/film/css/themes/default/index.css'
 * import '@mk/film-sheet'
 * ```
 *
 * The formula engine is also available on its own from `@mk/film-sheet/engine`.
 */
export { FilmSheet } from './ui/film-sheet.js'
export type { SheetChrome, SheetSelectDetail } from './ui/film-sheet.js'
export type { Workbook, WorkbookJSON, SheetJSON, Style } from './engine/index.js'

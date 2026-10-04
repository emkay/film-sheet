import { expect, type Locator, type Page } from '@playwright/test'

/**
 * Helpers for driving FilmSheet like a user. Assertions prefer what a user
 * (or assistive technology) can perceive — the focused grid cell's accessible
 * text, the formula bar, the name box, the status region — and fall back to
 * the app object only for coordinates the canvas doesn't expose to the DOM.
 */

export async function openApp (page: Page): Promise<void> {
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  page.on('console', (m) => {
    if (m.type() === 'error') errors.push(m.text())
  })
  ;(page as Page & { jsErrors: string[] }).jsErrors = errors
  await page.goto('/')
  await expect(page.locator('.grid[role="grid"]')).toBeVisible()
  await page.waitForFunction(() => !!document.querySelector('film-sheet')?._controller)
}

export function jsErrors (page: Page): string[] {
  return (page as Page & { jsErrors?: string[] }).jsErrors ?? []
}

/** File → New workbook, through the real menu. */
export async function newWorkbook (page: Page): Promise<void> {
  await page.getByRole('menuitem', { name: 'File' }).click()
  await page.getByRole('menuitem', { name: /New workbook/ }).click()
  await expect(page.getByRole('tab', { name: 'Sheet1' })).toBeVisible()
}

/** The proxy element that carries keyboard focus for the active cell. */
export function activeCell (page: Page): Locator {
  return page.locator('.cell-proxy:focus')
}

export async function expectActive (page: Page, text: string | RegExp): Promise<void> {
  await expect(activeCell(page)).toHaveText(text)
}

/** Screen position of a cell's centre (the grid is a canvas, so ask the app). */
export async function cellCenter (page: Page, address: string): Promise<{ x: number; y: number }> {
  return page.evaluate((addr) => {
    const m = /^([A-Z]+)(\d+)$/.exec(addr)!
    let col = 0
    for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64
    const app = document.querySelector('film-sheet')!._controller!
    const rect = app.grid.cellRect(Number(m[2]) - 1, col - 1)
    const host = app.grid.el.getBoundingClientRect()
    return { x: host.left + rect.x + rect.w / 2, y: host.top + rect.y + rect.h / 2 }
  }, address)
}

export async function clickCell (page: Page, address: string): Promise<void> {
  const { x, y } = await cellCenter(page, address)
  await page.mouse.click(x, y)
  await expect(page.locator('.name-box input').or(page.locator('.name-box'))).toBeAttached()
}

/** Go to a cell with the name box, as a keyboard user would (Ctrl+G). */
export async function goTo (page: Page, address: string): Promise<void> {
  const box = page.getByRole('textbox', { name: 'Name box' })
  await box.fill(address)
  await box.press('Enter')
  await expectActive(page, new RegExp(`^${address},`))
}

/** Type into the active cell and commit with Enter. */
export async function enter (page: Page, address: string, text: string): Promise<void> {
  await goTo(page, address)
  await page.keyboard.type(text)
  await page.keyboard.press('Enter')
}

export function formulaBar (page: Page): Locator {
  return page.getByRole('textbox', { name: 'Formula bar' })
}

export function status (page: Page): Locator {
  return page.locator('film-sheet .announcer')
}

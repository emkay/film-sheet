import { AxeBuilder } from '@axe-core/playwright'
import { expect, test } from '@playwright/test'
import { activeCell, enter, expectActive, goTo, jsErrors, newWorkbook, openApp, status } from './helpers'

test.beforeEach(async ({ page }) => {
  await openApp(page)
  await newWorkbook(page)
})

test.afterEach(({ page }) => {
  expect(jsErrors(page)).toEqual([])
})

test('the grid exposes the active cell with its row and column position', async ({ page }) => {
  const grid = page.getByRole('grid', { name: 'Sheet1' })
  await expect(grid).toHaveAttribute('aria-rowcount', '1048577')
  await expect(grid).toHaveAttribute('aria-colcount', '16385')
  await enter(page, 'C5', '=2+3')
  await goTo(page, 'C5')
  const cell = activeCell(page)
  await expect(cell).toHaveRole('gridcell')
  await expect(cell).toHaveAccessibleName('C5, 5')
  await expect(cell).toHaveAccessibleDescription(/Formula =2\+3/)
  await expect(cell).toHaveAttribute('aria-colindex', '4')
  await expect(page.getByRole('row').filter({ has: cell })).toHaveAttribute('aria-rowindex', '6')
  await expect(page.getByRole('columnheader', { name: 'C' })).toBeAttached()
  await expect(page.getByRole('rowheader', { name: '5' })).toBeAttached()
})

test('every arrow-key move puts focus on a fresh cell element (so it is announced)', async ({ page }) => {
  await goTo(page, 'A1')
  for (const [key, expected] of [['ArrowRight', 'B1, blank'], ['ArrowDown', 'B2, blank'], ['ArrowLeft', 'A2, blank']] as const) {
    await activeCell(page).evaluate((el) => el.setAttribute('data-was-focused', ''))
    await page.keyboard.press(key)
    await expectActive(page, expected)
    // A new element took focus: screen readers announce focus changes, not text edits.
    await expect(activeCell(page)).not.toHaveAttribute('data-was-focused')
    await page.locator('[data-was-focused]').evaluate((el) => el.removeAttribute('data-was-focused'))
  }
})

test('range selections, copy and undo are announced in the status region', async ({ page }) => {
  await goTo(page, 'A1')
  await page.keyboard.press('Shift+ArrowDown')
  await page.keyboard.press('Shift+ArrowRight')
  await expect(status(page)).toHaveText(/A1 to B2, 4 cells selected/)
  await page.keyboard.press('ControlOrMeta+c')
  await expect(status(page)).toHaveText('Copied A1:B2')
  await page.keyboard.press('ControlOrMeta+z')
  await expect(status(page)).toHaveText('Nothing to undo')
})

test('F6 cycles through the page regions and F10 opens the menu bar', async ({ page }) => {
  await goTo(page, 'A1')
  await page.keyboard.press('F6')
  await expect(page.getByRole('tab', { name: 'Sheet1' })).toBeFocused()
  await page.keyboard.press('F6')
  await expect(page.getByRole('menuitem', { name: 'File' })).toBeFocused()
  await page.keyboard.press('F6')
  await expect(page.getByRole('button', { name: 'Undo' })).toBeFocused()
  await page.keyboard.press('F6')
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toBeFocused()
  await page.keyboard.press('F6')
  await expect(activeCell(page)).toBeFocused()
  await page.keyboard.press('Shift+F6')
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toBeFocused()
  await page.keyboard.press('Escape')
  await expect(activeCell(page)).toBeFocused()
  await page.keyboard.press('F10')
  await expect(page.getByRole('menuitem', { name: 'File' })).toBeFocused()
})

test('Tab leaves the formula bar when not editing (no keyboard trap)', async ({ page }) => {
  await page.getByRole('textbox', { name: 'Formula bar' }).focus()
  await page.keyboard.press('Tab')
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).not.toBeFocused()
})

test('the toolbar is one tab stop with arrow-key navigation and real pressed states', async ({ page }) => {
  await goTo(page, 'A1')
  await page.keyboard.press('F6')
  await page.keyboard.press('F6')
  await page.keyboard.press('F6')
  const undo = page.getByRole('button', { name: 'Undo' })
  await expect(undo).toBeFocused()
  await page.keyboard.press('ArrowRight')
  await expect(page.getByRole('button', { name: 'Redo' })).toBeFocused()
  await expect(undo).toHaveAttribute('tabindex', '-1')
  await page.keyboard.press('Home')
  await expect(undo).toBeFocused()

  await page.keyboard.press('F6')
  await page.keyboard.press('F6')
  await page.keyboard.press('ControlOrMeta+b')
  await expect(page.getByRole('button', { name: 'Bold' })).toHaveAttribute('aria-pressed', 'true')
  await expect(page.getByRole('button', { name: 'Italic' })).toHaveAttribute('aria-pressed', 'false')
})

test('Shift+F10 opens the context menu at the cell; Escape returns focus to the grid', async ({ page }) => {
  await goTo(page, 'B2')
  await page.keyboard.press('Shift+F10')
  const menu = page.locator('#context-menu')
  await expect(menu).toBeVisible()
  await expect(menu.getByRole('menuitem', { name: /^Cut/ })).toBeFocused()
  await page.keyboard.press('Escape')
  await expect(menu).toBeHidden()
  await expect(activeCell(page)).toBeFocused()
  await expectActive(page, 'B2, blank')
})

test('Enter on Cancel in the rename dialog cancels', async ({ page }) => {
  await page.getByRole('menuitem', { name: 'Sheet', exact: true }).click()
  await page.getByRole('menuitem', { name: 'Rename…' }).click()
  await page.getByRole('textbox', { name: 'Sheet name' }).fill('Renamed')
  await page.getByRole('button', { name: 'Cancel' }).focus()
  await page.keyboard.press('Enter')
  await expect(page.getByRole('tab', { name: 'Sheet1' })).toBeVisible()
})

test('column width can be set from the keyboard', async ({ page }) => {
  await goTo(page, 'B1')
  await page.getByRole('menuitem', { name: 'Format' }).click()
  await page.getByRole('menuitem', { name: 'Column width…' }).click()
  const size = page.getByRole('textbox', { name: 'Size in pixels' })
  await expect(size).toHaveValue('100')
  await size.fill('180')
  await size.press('Enter')
  const width = await page.evaluate(() => document.querySelector('film-sheet')!._controller!.grid.cols.size(1))
  expect(width).toBe(180)
})

test('axe finds no serious or critical violations', async ({ page }) => {
  await page.getByRole('button', { name: 'Load sample data' }).click()
  await goTo(page, 'B3')
  const results = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21aa', 'wcag22aa']).analyze()
  const serious = results.violations
    .filter((v) => v.impact === 'serious' || v.impact === 'critical')
    .map((v) => `${v.id}: ${v.nodes.map((n) => n.target.join(' ')).join(', ')}`)
  expect(serious).toEqual([])
})

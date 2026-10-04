import { expect, test } from '@playwright/test'
import { activeCell, cellCenter, enter, expectActive, formulaBar, goTo, jsErrors, newWorkbook, openApp } from './helpers'

test.beforeEach(async ({ page }) => {
  await openApp(page)
  await newWorkbook(page)
})

test.afterEach(({ page }) => {
  expect(jsErrors(page)).toEqual([])
})

test('loads the sample workbook with calculated results', async ({ page }) => {
  await page.getByRole('button', { name: 'Load sample data' }).click()
  await goTo(page, 'B7')
  await expectActive(page, 'B7, $146,678.75')
  await expect(formulaBar(page)).toHaveValue('=SUM(B3:B6)')
})

test('typing, Enter and Tab commit and move like Excel', async ({ page }) => {
  await goTo(page, 'A1')
  await page.keyboard.type('10')
  await page.keyboard.press('Tab')
  await expectActive(page, 'B1, blank')
  await page.keyboard.type('=a1*2')
  await page.keyboard.press('Enter')
  await expectActive(page, 'B2, blank')
  await page.keyboard.press('ArrowUp')
  await expectActive(page, 'B1, 20')
  await expect(formulaBar(page)).toHaveValue('=A1*2')
})

test('editing an input recalculates dependents', async ({ page }) => {
  await enter(page, 'A1', '2')
  await enter(page, 'A2', '=A1^10')
  await enter(page, 'A3', '=SUM(A1:A2)')
  await enter(page, 'A1', '3')
  await goTo(page, 'A3')
  await expectActive(page, 'A3, 59052')
})

test('Escape cancels an edit; F2 edits in place', async ({ page }) => {
  await enter(page, 'A1', 'keep')
  await goTo(page, 'A1')
  await page.keyboard.type('discard')
  await page.keyboard.press('Escape')
  await expectActive(page, 'A1, keep')
  await page.keyboard.press('F2')
  await page.keyboard.type('er')
  await page.keyboard.press('Enter')
  await page.keyboard.press('ArrowUp')
  await expectActive(page, 'A1, keeper')
})

test('arrow keys point references into a formula; F4 makes them absolute', async ({ page }) => {
  await enter(page, 'A1', '5')
  await enter(page, 'A2', '6')
  await goTo(page, 'B3')
  await page.keyboard.type('=SUM(')
  await page.keyboard.press('ArrowLeft')
  await page.keyboard.press('ArrowUp')
  await page.keyboard.press('ArrowUp')
  await page.keyboard.press('Shift+ArrowDown')
  await page.keyboard.press('F4')
  await page.keyboard.press('Enter')
  await goTo(page, 'B3')
  await expect(formulaBar(page)).toHaveValue('=SUM($A$1:$A$2)')
  await expectActive(page, 'B3, 11')
})

test('clicking a cell while typing a formula inserts its reference', async ({ page }) => {
  await enter(page, 'C3', '42')
  await goTo(page, 'A1')
  await page.keyboard.type('=')
  const c3 = await cellCenter(page, 'C3')
  await page.mouse.click(c3.x, c3.y)
  await page.keyboard.type('+1')
  await page.keyboard.press('Enter')
  await goTo(page, 'A1')
  await expect(formulaBar(page)).toHaveValue('=C3+1')
  await expectActive(page, 'A1, 43')
})

test('function autocomplete is exposed to assistive tech and Tab inserts', async ({ page }) => {
  await goTo(page, 'A1')
  await page.keyboard.type('=SUMI')
  const editor = page.getByRole('textbox', { name: 'Edit A1' })
  await expect(editor).toHaveAttribute('aria-autocomplete', 'list')
  const list = page.getByRole('listbox', { name: 'Function suggestions' })
  await expect(list.getByRole('option')).toHaveCount(2)
  await page.keyboard.press('ArrowDown')
  await expect(editor).toHaveAttribute('aria-activedescendant', 'fn-opt-SUMIFS')
  await page.keyboard.press('Tab')
  await expect(editor).toHaveValue('=SUMIFS(')
  await page.keyboard.press('Escape')
})

test('Enter commits =SUM(A:A instead of accepting a suggestion', async ({ page }) => {
  await goTo(page, 'B1')
  await page.keyboard.type('=SUM(A:A')
  await page.keyboard.press('Enter')
  await goTo(page, 'B1')
  await expect(formulaBar(page)).toHaveValue('=SUM(A:A)')
})

test('typing in the formula bar edits the active cell', async ({ page }) => {
  await goTo(page, 'D4')
  await formulaBar(page).fill('=6*7')
  await formulaBar(page).press('Enter')
  await goTo(page, 'D4')
  await expectActive(page, 'D4, 42')
})

test('errors are shown and explained', async ({ page }) => {
  await enter(page, 'A1', '=1/0')
  await enter(page, 'A2', '=ROUND()')
  await goTo(page, 'A1')
  await expectActive(page, 'A1, #DIV/0!')
  await expect(page.getByText(/ROUND expects/)).toBeVisible()
  await expect(activeCell(page)).toBeFocused()
})

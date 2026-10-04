import { expect, test } from '@playwright/test'
import { activeCell, cellCenter, enter, expectActive, formulaBar, goTo, jsErrors, newWorkbook, openApp } from './helpers'

test.beforeEach(async ({ page }) => {
  await openApp(page)
  await newWorkbook(page)
})

test.afterEach(({ page }) => {
  expect(jsErrors(page)).toEqual([])
})

test('copy and paste shifts relative references and uses the system clipboard', async ({ page }) => {
  await enter(page, 'A1', '1')
  await enter(page, 'A2', '2')
  await enter(page, 'B1', '=A1*10')
  await goTo(page, 'B1')
  await page.keyboard.press('ControlOrMeta+c')
  expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('10')
  await page.keyboard.press('ArrowDown')
  await page.keyboard.press('ControlOrMeta+v')
  await expect(formulaBar(page)).toHaveValue('=A2*10')
  await expectActive(page, 'B2, 20')
})

test('copying a block with blanks pastes the blanks too', async ({ page }) => {
  await enter(page, 'D1', 'old')
  await enter(page, 'A1', 'x')
  await goTo(page, 'A1')
  await page.keyboard.press('Shift+ArrowRight')
  await page.keyboard.press('ControlOrMeta+c')
  await goTo(page, 'C1')
  await page.keyboard.press('ControlOrMeta+v')
  await goTo(page, 'D1')
  await expectActive(page, 'D1, blank')
})

test('pastes tab-separated text from other apps', async ({ page }) => {
  await page.evaluate(() => navigator.clipboard.writeText('x\t1\ny\t2\n'))
  await goTo(page, 'C3')
  await page.keyboard.press('ControlOrMeta+v')
  await goTo(page, 'D4')
  await expectActive(page, 'D4, 2')
  await goTo(page, 'C4')
  await expectActive(page, 'C4, y')
})

test('cut and paste moves cells; any edit in between cancels the cut', async ({ page }) => {
  await enter(page, 'A1', 'moved')
  await goTo(page, 'A1')
  await page.keyboard.press('ControlOrMeta+x')
  await goTo(page, 'C1')
  await page.keyboard.press('ControlOrMeta+v')
  await expectActive(page, 'C1, moved')
  await goTo(page, 'A1')
  await expectActive(page, 'A1, blank')

  await enter(page, 'B5', 'stay')
  await goTo(page, 'B5')
  await page.keyboard.press('ControlOrMeta+x')
  await enter(page, 'E5', 'unrelated edit')
  await goTo(page, 'B6')
  await page.keyboard.press('ControlOrMeta+v')
  await goTo(page, 'B5')
  await expectActive(page, 'B5, stay')
})

test('undo and redo, including formatting', async ({ page }) => {
  await enter(page, 'A1', 'first')
  await enter(page, 'A1', 'second')
  await goTo(page, 'A1')
  await page.keyboard.press('ControlOrMeta+b')
  await page.keyboard.press('ControlOrMeta+z')
  await page.keyboard.press('ControlOrMeta+z')
  await expectActive(page, 'A1, first')
  await page.keyboard.press('ControlOrMeta+y')
  await expectActive(page, 'A1, second')
  await page.keyboard.press('ControlOrMeta+y')
  await expect(activeCell(page)).toHaveAccessibleDescription(/bold/)
})

test('dragging the fill handle continues a series', async ({ page }) => {
  await enter(page, 'A1', '5')
  await enter(page, 'A2', '10')
  await goTo(page, 'A1')
  await page.keyboard.press('Shift+ArrowDown')
  const a2 = await cellCenter(page, 'A2')
  const a5 = await cellCenter(page, 'A5')
  // The handle is the small square at the selection's bottom-right corner.
  const handle = { x: a2.x + 50 - 1, y: a2.y + 12 - 1 }
  await page.mouse.move(handle.x, handle.y)
  await page.mouse.down()
  await page.mouse.move(a5.x, a5.y, { steps: 5 })
  await page.mouse.up()
  await goTo(page, 'A5')
  await expectActive(page, 'A5, 25')
})

test('Fill series from the menu works without a mouse', async ({ page }) => {
  await enter(page, 'B1', 'Item 1')
  await goTo(page, 'B1')
  for (let i = 0; i < 3; i++) await page.keyboard.press('Shift+ArrowDown')
  await page.getByRole('menuitem', { name: 'Edit' }).click()
  await page.getByRole('menuitem', { name: 'Fill series down' }).click()
  await goTo(page, 'B4')
  await expectActive(page, 'B4, Item 4')
})

test('inserting a row shifts formulas that reference below it', async ({ page }) => {
  await enter(page, 'A2', '7')
  await enter(page, 'B1', '=A2*2')
  await goTo(page, 'A2')
  await page.getByRole('menuitem', { name: 'Insert' }).click()
  await page.getByRole('menuitem', { name: 'Row above' }).click()
  await goTo(page, 'B1')
  await expect(formulaBar(page)).toHaveValue('=A3*2')
  await expectActive(page, 'B1, 14')
})

test('sheets: add, rename with F2, cross-sheet formulas follow renames', async ({ page }) => {
  await page.getByRole('button', { name: 'Add sheet' }).click()
  await expect(page.getByRole('tab', { name: 'Sheet2' })).toHaveAttribute('aria-selected', 'true')
  await enter(page, 'A1', '99')
  await page.getByRole('tab', { name: 'Sheet2' }).focus()
  await page.keyboard.press('F2')
  const name = page.getByRole('textbox', { name: 'Sheet name' })
  await name.fill('Data')
  await name.press('Enter')
  await expect(page.getByRole('tab', { name: 'Data' })).toBeVisible()
  await page.getByRole('tab', { name: 'Sheet1' }).click()
  await enter(page, 'A1', '=Data!A1+1')
  await goTo(page, 'A1')
  await expectActive(page, 'A1, 100')
})

test('the workbook survives a reload (autosave)', async ({ page }) => {
  await enter(page, 'C2', 'remember me')
  await expect(page.getByText('Saved')).toBeVisible()
  await page.reload()
  await page.waitForFunction(() => !!document.querySelector('film-sheet')?._controller)
  await goTo(page, 'C2')
  await expectActive(page, 'C2, remember me')
})

import { expect, test, type Page } from '@playwright/test'
import type { FilmSheet } from '../src/ui/film-sheet'
import { jsErrors } from './helpers'

/**
 * <film-sheet> as a library component, embedded in a page it doesn't own:
 * two instances, an unrelated input, and the public API.
 */

async function openEmbed (page: Page): Promise<void> {
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  ;(page as Page & { jsErrors: string[] }).jsErrors = errors
  await page.goto('/e2e/embed.html')
  await page.waitForFunction(() => [...document.querySelectorAll('film-sheet')].every((s) => s._controller))
}

const one = (page: Page) => page.locator('#one')
const two = (page: Page) => page.locator('#two')
const active = (sheet: ReturnType<typeof one>) => sheet.locator('.cell-proxy:focus')

test.beforeEach(async ({ page }) => openEmbed(page))
test.afterEach(({ page }) => {
  expect(jsErrors(page)).toEqual([])
})

test('does not touch the page: no focus grab, title, globals or storage', async ({ page }) => {
  await expect(page).toHaveTitle('Embed test page')
  expect(await page.evaluate(() => document.activeElement === document.body)).toBe(true)
  expect(await page.evaluate(() => ['filmSheet', 'app'].filter((k) => k in window))).toEqual([])
  await one(page).evaluate((el) => el.focus())
  await page.keyboard.type('42')
  await page.keyboard.press('Enter')
  expect(await page.evaluate(() => Object.keys(localStorage))).toEqual([])
  expect(await page.evaluate(() => document.documentElement.dataset.theme)).toBeUndefined()
})

test('two sheets on one page keep separate data, focus and keyboard', async ({ page }) => {
  await one(page).evaluate((el) => el.focus())
  await page.keyboard.type('first')
  await page.keyboard.press('Enter')
  await two(page).evaluate((el) => el.focus())
  await page.keyboard.type('second')
  await page.keyboard.press('Enter')
  const a1 = (id: string) => page.evaluate((sel) => document.querySelector<FilmSheet>(sel)!.workbook!.getValue(document.querySelector<FilmSheet>(sel)!.workbook!.sheets[0].id, 0, 0), id)
  expect(await a1('#one')).toBe('first')
  expect(await a1('#two')).toBe('second')
  // Only one grid shows the focused selection.
  await expect(active(two(page))).toHaveCount(1)
  await expect(active(one(page))).toHaveCount(0)
})

test('typing and clipboard shortcuts in the page are left alone', async ({ page }) => {
  const input = page.locator('#outside')
  await input.focus()
  await input.press('End')
  await page.keyboard.type('!')
  await expect(input).toHaveValue('page text!')
  await input.selectText()
  await page.keyboard.press('ControlOrMeta+c')
  expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('page text!')
  const value = await page.evaluate(() => document.querySelector<FilmSheet>('#one')!.workbook!.getValue(document.querySelector<FilmSheet>('#one')!.workbook!.sheets[0].id, 0, 0))
  expect(value).toBeNull()
})

test('load(), toJSON() and the data property round-trip; bad data is rejected', async ({ page }) => {
  const result = await page.evaluate(() => {
    const el = document.querySelector<FilmSheet>('#one')!
    el.load({ version: 1, sheets: [{ name: 'Budget', cells: [[0, 0, '2'], [1, 0, '=A1*21']] }] })
    const wb = el.workbook!
    const before = wb.getValue(wb.sheets[0].id, 1, 0)
    let error = ''
    try {
      el.load({ nope: true } as never)
    } catch (e) {
      error = (e as Error).message
    }
    const json = el.toJSON()!
    const other = document.querySelector<FilmSheet>('#two')!
    other.data = json
    return { before, error, names: json.sheets.map((s) => s.name), copied: other.workbook!.getValue(other.workbook!.sheets[0].id, 1, 0) }
  })
  expect(result).toEqual({ before: 42, error: expect.stringMatching(/Not a FilmSheet workbook/), names: ['Budget'], copied: 42 })
  await expect(one(page).getByRole('tab', { name: 'Budget' })).toBeVisible()
})

test('fires film-sheet-change and film-sheet-select', async ({ page }) => {
  await page.evaluate(() => {
    const el = document.querySelector<FilmSheet>('#one')!
    const log: string[] = ((window as unknown as { log: string[] }).log = [])
    el.addEventListener('film-sheet-change', () => log.push('change'))
    el.addEventListener('film-sheet-select', (e) => log.push(`select ${e.detail.sheet} ${e.detail.address} ${e.detail.range}`))
  })
  await one(page).evaluate((el) => el.focus())
  await page.keyboard.press('Shift+ArrowRight')
  await page.keyboard.type('x')
  await page.keyboard.press('Enter')
  const log = await page.evaluate(() => (window as unknown as { log: string[] }).log)
  expect(log).toContain('select Sheet1 A1 A1:B1')
  expect(log).toContain('change')
})

test('readonly blocks edits but keeps viewing, copying and searching', async ({ page }) => {
  await page.evaluate(() => document.querySelector<FilmSheet>('#one')!.load({ version: 1, sheets: [{ name: 'S', cells: [[0, 0, 'locked']] }] }))
  await one(page).evaluate((el) => {
    ;(el as FilmSheet).readonly = true
    el.focus()
  })
  await expect(one(page)).toHaveAttribute('readonly', '')
  await expect(one(page).locator('.toolbar')).toBeHidden()
  await page.keyboard.type('changed')
  await page.keyboard.press('Delete')
  await expect(one(page).locator('.announcer')).toHaveText('This sheet is read-only')
  await expect(active(one(page))).toHaveText('A1, locked')
  await page.keyboard.press('ControlOrMeta+c')
  expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('locked')
  await one(page).getByRole('menuitem', { name: 'Edit' }).click()
  await expect(one(page).getByRole('menuitem', { name: /^Paste/ }).first()).toHaveAttribute('disabled', '')
  await expect(one(page).getByRole('menuitem', { name: /^Copy/ })).not.toHaveAttribute('disabled', '')
  await page.keyboard.press('Escape')

  await one(page).evaluate((el) => {
    ;(el as FilmSheet).readonly = false
    el.focus()
  })
  await page.keyboard.type('unlocked')
  await page.keyboard.press('Enter')
  expect(await page.evaluate(() => document.querySelector<FilmSheet>('#one')!.workbook!.getValue(document.querySelector<FilmSheet>('#one')!.workbook!.sheets[0].id, 0, 0))).toBe('unlocked')
})

test('chrome="minimal" and "none" hide UI, and F6 skips what is hidden', async ({ page }) => {
  await expect(two(page).locator('film-menu-bar')).toBeHidden()
  await expect(two(page).locator('.toolbar')).toBeHidden()
  await expect(two(page).locator('.formula-bar')).toBeVisible()
  const gridHeight = async (): Promise<number> => (await two(page).locator('.grid-host').boundingBox())?.height ?? 0
  expect(await gridHeight()).toBeGreaterThan(150)
  await two(page).evaluate((el) => el.focus())
  await page.keyboard.press('F6')
  await expect(two(page).getByRole('tab', { name: 'Sheet1' })).toBeFocused()
  await page.keyboard.press('F6')
  await expect(two(page).getByRole('textbox', { name: 'Formula bar' })).toBeFocused()

  await two(page).evaluate((el) => {
    ;(el as FilmSheet).chrome = 'none'
  })
  await expect(two(page).locator('.formula-bar')).toBeHidden()
  await expect(two(page).locator('.sheet-bar')).toBeHidden()
  await expect(two(page).getByRole('grid')).toBeVisible()
  expect(await gridHeight()).toBeGreaterThan(250)
})

test('survives being removed and re-attached, keeping its data', async ({ page }) => {
  await one(page).evaluate((el) => el.focus())
  await page.keyboard.type('kept')
  await page.keyboard.press('Enter')
  const value = await page.evaluate(async () => {
    const el = document.querySelector<FilmSheet>('#one')!
    const parent = el.parentElement!
    el.remove()
    // Resizing while detached must not touch the disposed grid.
    window.dispatchEvent(new Event('resize'))
    await new Promise((r) => setTimeout(r, 50))
    parent.append(el)
    await new Promise((r) => setTimeout(r, 50))
    return el.workbook!.getValue(el.workbook!.sheets[0].id, 0, 0)
  })
  expect(value).toBe('kept')
  await expect(one(page).locator('.grid')).toHaveCount(1)
  await expect(one(page).locator('.cell-editor')).toHaveCount(1)
})

test('storage-key autosaves and restores', async ({ page }) => {
  await page.evaluate(() => {
    const el = document.querySelector<FilmSheet>('#one')!
    el.storageKey = 'embed-test'
    el.load({ version: 1, sheets: [{ name: 'Saved', cells: [[0, 0, 'persisted']] }] })
  })
  await expect.poll(() => page.evaluate(() => localStorage.getItem('embed-test'))).toContain('persisted')
  await page.evaluate(() => {
    const el = document.createElement('film-sheet')
    el.id = 'three'
    el.storageKey = 'embed-test'
    document.body.append(el)
  })
  await expect(page.locator('#three').getByRole('tab', { name: 'Saved' })).toBeVisible()
})

test('sizes itself independently of the page root font size', async ({ page }) => {
  const menuHeight = async (): Promise<number> => (await one(page).locator('film-menu-bar').boundingBox())!.height
  const before = await menuHeight()
  await page.evaluate(() => {
    document.documentElement.style.fontSize = '24px'
  })
  expect(await menuHeight()).toBeCloseTo(before, 0)
  await one(page).evaluate((el) => el.style.setProperty('--film-sheet-size', '18px'))
  expect(await menuHeight()).toBeGreaterThan(before)
})

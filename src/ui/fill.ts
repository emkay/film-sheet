import type { RangeAddress } from '../engine/address.js'
import { shiftFormula } from '../engine/refs.js'
import type { CellPatch, Sheet } from '../engine/workbook.js'

/**
 * Patches for dragging the fill handle from `src` out to `dest` (which
 * contains src). Numbers continue a linear trend, "Item 7" style text
 * counts up, dates step by a day, and formulas shift their relative
 * references — the behaviour people expect from Excel. `copyOnly` (Ctrl+D /
 * Ctrl+R) skips series detection and just copies.
 */
export function fillPatches (sheet: Sheet, src: RangeAddress, dest: RangeAddress, copyOnly = false): CellPatch[] {
  const patches: CellPatch[] = []
  const vertical = dest.r1 !== src.r1 || dest.r2 !== src.r2
  const lanes = vertical ? [src.c1, src.c2] : [src.r1, src.r2]
  const [from, to] = vertical ? [src.r1, src.r2] : [src.c1, src.c2]
  const [tFrom, tTo] = vertical ? [dest.r1, dest.r2] : [dest.c1, dest.c2]
  const n = to - from + 1

  for (let lane = lanes[0]; lane <= lanes[1]; lane++) {
    const get = (i: number) => (vertical ? sheet.get(i, lane) : sheet.get(lane, i))
    const sources = Array.from({ length: n }, (_, k) => get(from + k))
    const series = copyOnly ? undefined : numericSeries(sources)
    const textSeries = series || copyOnly ? undefined : countingText(sources)
    for (let t = tFrom; t <= tTo; t++) {
      if (t >= from && t <= to) continue
      const offset = t - from
      const k = ((offset % n) + n) % n
      const srcCell = sources[k]
      const raw = srcCell?.raw ?? ''
      let out = raw
      if (series) out = String(roundNice(series.intercept + series.slope * offset))
      else if (textSeries) out = textSeries.prefix + (textSeries.start + textSeries.step * offset) + textSeries.suffix
      else if (raw[0] === '=') out = shiftFormula(raw, vertical ? t - (from + k) : 0, vertical ? 0 : t - (from + k))
      const patch: CellPatch = { r: vertical ? t : lane, c: vertical ? lane : t, raw: out, style: srcCell?.style ?? null }
      patches.push(patch)
    }
  }
  return patches
}

type MaybeCell = ReturnType<Sheet['get']>

function numericSeries (cells: MaybeCell[]): { slope: number; intercept: number } | undefined {
  if (!cells.length) return undefined
  const ys: number[] = []
  for (const c of cells) {
    if (!c || c.formula || typeof c.value !== 'number') return undefined
    ys.push(c.value)
  }
  if (ys.length === 1) {
    // A lone date counts up by days; a lone number is copied.
    const fmt = cells[0]!.style?.fmt ?? ''
    if (/[dy]/i.test(fmt) && !/[0#]/.test(fmt)) return { slope: 1, intercept: ys[0] }
    return undefined
  }
  // Least-squares linear trend, as Excel does.
  const n = ys.length
  const mx = (n - 1) / 2
  const my = ys.reduce((a, b) => a + b, 0) / n
  let sxy = 0
  let sxx = 0
  ys.forEach((y, x) => {
    sxy += (x - mx) * (y - my)
    sxx += (x - mx) * (x - mx)
  })
  const slope = sxy / sxx
  return { slope, intercept: my - slope * mx }
}

function countingText (cells: MaybeCell[]): { prefix: string; suffix: string; start: number; step: number } | undefined {
  const parts = cells.map((c) => (c && !c.formula && typeof c.value === 'string' ? /^(.*?)(\d+)(\D*)$/.exec(c.value) : null))
  if (!parts.length || parts.some((p) => !p)) return undefined
  const [first] = parts as RegExpExecArray[]
  if (parts.some((p) => p![1] !== first[1] || p![3] !== first[3])) return undefined
  const nums = parts.map((p) => Number(p![2]))
  const step = nums.length > 1 ? nums[1] - nums[0] : 1
  for (let i = 2; i < nums.length; i++) if (nums[i] - nums[i - 1] !== step) return undefined
  return { prefix: first[1], suffix: first[3], start: nums[0], step }
}

function roundNice (v: number): number {
  return Number(v.toPrecision(12))
}

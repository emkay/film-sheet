/**
 * Sizes along one axis (row heights or column widths) for up to a million
 * entries. Only overridden sizes are stored; offsets come from
 * `index * default + (sum of overrides before index)`, found by binary
 * search over the sorted overrides, so lookups stay O(log n).
 */
export class Axis {
  private sorted: number[] = []
  private prefix: number[] = []
  private stale = true

  constructor (
    readonly count: number,
    readonly defaultSize: number,
    private sizes: Map<number, number>
  ) {}

  /** Point at a different sheet's size map. */
  use (sizes: Map<number, number>): void {
    this.sizes = sizes
    this.stale = true
  }

  invalidate (): void {
    this.stale = true
  }

  size (i: number): number {
    return this.sizes.get(i) ?? this.defaultSize
  }

  setSize (i: number, size: number): void {
    if (size === this.defaultSize) this.sizes.delete(i)
    else this.sizes.set(i, size)
    this.stale = true
  }

  private rebuild (): void {
    this.sorted = [...this.sizes.keys()].sort((a, b) => a - b)
    this.prefix = new Array(this.sorted.length)
    let acc = 0
    for (let k = 0; k < this.sorted.length; k++) {
      acc += this.sizes.get(this.sorted[k])! - this.defaultSize
      this.prefix[k] = acc
    }
    this.stale = false
  }

  /** Pixel offset of the start of entry i. */
  offset (i: number): number {
    if (this.stale) this.rebuild()
    // number of overrides with index < i
    let lo = 0
    let hi = this.sorted.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (this.sorted[mid] < i) lo = mid + 1
      else hi = mid
    }
    return i * this.defaultSize + (lo > 0 ? this.prefix[lo - 1] : 0)
  }

  total (): number {
    return this.offset(this.count)
  }

  /**
   * Index of the entry containing pixel offset px (clamped). Binary search
   * over the overridden entries only — O(log k) — then arithmetic across the
   * run of default-sized entries that follows.
   */
  indexAt (px: number): number {
    if (px <= 0) return 0
    if (this.stale) this.rebuild()
    const d = this.defaultSize
    const sorted = this.sorted
    // Last override whose start offset is ≤ px.
    let lo = 0
    let hi = sorted.length - 1
    let k = -1
    while (lo <= hi) {
      const mid = (lo + hi) >> 1
      const start = sorted[mid] * d + (mid > 0 ? this.prefix[mid - 1] : 0)
      if (start <= px) {
        k = mid
        lo = mid + 1
      } else hi = mid - 1
    }
    let i: number
    if (k < 0) i = Math.floor(px / d)
    else {
      const idx = sorted[k]
      const start = idx * d + (k > 0 ? this.prefix[k - 1] : 0)
      const end = start + this.sizes.get(idx)!
      i = px < end ? idx : idx + 1 + Math.floor((px - end) / d)
    }
    return Math.min(this.count - 1, i)
  }
}

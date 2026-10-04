import { describe, expect, it } from 'vitest'
import { Axis } from '../src/ui/axis'

describe('Axis', () => {
  it('indexAt inverts offset with mixed sizes', () => {
    const sizes = new Map<number, number>([[3, 50], [4, 5], [10, 200], [1000, 1]])
    const ax = new Axis(1_048_576, 24, sizes)
    for (let i = 0; i < 1200; i++) {
      const start = ax.offset(i)
      expect(ax.indexAt(start)).toBe(i)
      expect(ax.indexAt(start + ax.size(i) - 0.5)).toBe(i)
    }
    expect(ax.indexAt(1e12)).toBe(1_048_575)
  })
})

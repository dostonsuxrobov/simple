// src/imaging/random.ts (WP2)
// Deterministic pseudo-random numbers for noise, dithering and the Dissolve blend mode.
// Pure and DOM-free: the same seed gives the same pixels in the window, the worker and Node tests.

/**
 * mulberry32: a small, fast, seeded 32-bit generator. Returns a function that yields floats in [0, 1).
 * Use it for sequential streams; use hash2 when the value must depend only on a pixel position.
 */
export function mulberry32(seed: number): () => number {
  let state = seed >>> 0
  return () => {
    state = (state + 0x6d2b79f5) >>> 0
    let t = state
    t = Math.imul(t ^ (t >>> 15), t | 1)
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

/** murmur3 finaliser: full avalanche of a 32-bit value. */
function fmix32(value: number): number {
  let h = value | 0
  h ^= h >>> 16
  h = Math.imul(h, 0x85ebca6b)
  h ^= h >>> 13
  h = Math.imul(h, 0xc2b2ae35)
  h ^= h >>> 16
  return h
}

/**
 * Stateless hash of an integer position and a seed to a float in [0, 1). Coordinates are truncated to
 * 32-bit integers. hash2(x, y, s) never depends on evaluation order, so tiles, stripes and workers that
 * see the same position produce the same value.
 */
export function hash2(x: number, y: number, seed: number): number {
  let h = fmix32((seed | 0) + 0x9e3779b9)
  h = fmix32((h + Math.imul(x | 0, 0xcc9e2d51)) | 0)
  h = fmix32((h + Math.imul(y | 0, 0x1b873593)) | 0)
  return (h >>> 0) / 4294967296
}

/** Combines a base seed with extra integers (layer id hash, tile key, channel) into a new 32-bit seed. */
export function mixSeed(seed: number, ...values: readonly number[]): number {
  let h = fmix32((seed | 0) ^ 0x5bd1e995)
  for (const value of values) h = fmix32((h + Math.imul(value | 0, 0x9e3779b1)) | 0)
  return h >>> 0
}

/** 32-bit FNV-1a hash of a string (for seeds derived from layer ids). */
export function hashString(text: string): number {
  let h = 0x811c9dc5
  for (let index = 0; index < text.length; index += 1) {
    h ^= text.charCodeAt(index)
    h = Math.imul(h, 0x01000193)
  }
  return h >>> 0
}

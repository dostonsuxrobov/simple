export interface ScanReplacementRect { x: number; y: number; width: number; height: number }

export interface ScanReplacementInput {
  /** The line's words in reading order, aligned with the tokens of originalText. */
  words: Array<{ text: string; rect: ScanReplacementRect }>
  originalText: string
  newText: string
  /** Width of text (points) at the edit's size, font and horizontal scale. */
  measure: (text: string) => number
  spaceWidth: number
  /** Word extents along a slanted baseline (default: the boxes' x). */
  along?: Array<{ start: number; end: number }>
  /** How far the text may run at most (default: the last word's end). */
  lineEnd?: number
}

export type ScanReplacementPlan =
  | { kind: 'none' }
  | { kind: 'line' }
  | {
      kind: 'words'
      first: number
      /** Inclusive; first - 1 for a pure insertion. */
      last: number
      originalText: string
      text: string
      /** Union of the replaced words' boxes; null for a pure insertion. */
      rect: ScanReplacementRect | null
      start: number
      end: number
      /** Width the new text may use. */
      room: number
    }

export const MAX_FIT_COMPRESSION: number
export function replacementTokens(text: string): string[]
export function planScanReplacement(input: ScanReplacementInput): ScanReplacementPlan

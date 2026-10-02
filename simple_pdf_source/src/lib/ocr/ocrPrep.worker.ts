// Module worker: turns a rendered page (RGBA) into Tesseract's input (PGM) off
// the main thread. The pipeline itself lives in electron/ocr-preprocess.mjs so
// Node tests exercise exactly the same code.
//
// It also prepares edits of scanned text ('retouch-line', see scanEdit.ts):
// the printed font of a line is measured (electron/scan-style.mjs) and its
// glyphs are retouched out of the scan (electron/ocr-retouch.mjs).
import { encodePgm, preparePage, sha256Hex } from '../../../electron/ocr-preprocess.mjs'
import { composePatch, retouchLine } from '../../../electron/ocr-retouch.mjs'
import { estimateScanStyle } from '../../../electron/scan-style.mjs'
import type { OcrPrepRequest, OcrPrepResponse } from './types'
import type { ScanRetouchRequest, ScanRetouchResponse } from './scanEdit'

// The project compiles against the DOM library, where `self` is a Window.
interface PrepWorkerScope {
  addEventListener(type: 'message', listener: (event: MessageEvent<OcrPrepRequest | ScanRetouchRequest>) => void): void
  postMessage(message: OcrPrepResponse | ScanRetouchResponse, options?: { transfer?: Transferable[] }): void
}
const scope = self as unknown as PrepWorkerScope

async function preparePageMessage(request: OcrPrepRequest) {
  try {
    const pixels = new Uint8ClampedArray(request.rgba)
    const prepared = preparePage(pixels, request.width, request.height, { channels: 4, dpi: request.dpi })
    const pgm = encodePgm(prepared.image)
    const imageSha256 = await sha256Hex(pgm)
    scope.postMessage({
      type: 'prepared',
      id: request.id,
      blank: prepared.blank,
      width: prepared.image.width,
      height: prepared.image.height,
      pgm: prepared.blank ? null : pgm,
      imageSha256,
      deskew: prepared.deskew,
      stats: prepared.stats,
    }, { transfer: prepared.blank ? [] : [pgm.buffer as ArrayBuffer] })
  } catch (error) {
    scope.postMessage({ type: 'error', id: request.id, message: error instanceof Error ? error.message : String(error) })
  }
}

// FileReaderSync exists only in workers (not in the DOM library this project compiles against).
const FileReaderSyncClass = (globalThis as { FileReaderSync?: new () => { readAsDataURL(blob: Blob): string } }).FileReaderSync

/** PNG data URL of RGBA pixels (OffscreenCanvas + FileReaderSync), or null where the worker cannot encode. */
async function encodePng(rgba: Uint8ClampedArray, width: number, height: number): Promise<string | null> {
  try {
    if (typeof OffscreenCanvas === 'undefined' || !FileReaderSyncClass) return null
    const canvas = new OffscreenCanvas(width, height)
    const context = canvas.getContext('2d')
    if (!context) return null
    context.putImageData(new ImageData(new Uint8ClampedArray(rgba), width, height), 0, 0)
    const blob = await canvas.convertToBlob({ type: 'image/png' })
    return new FileReaderSyncClass().readAsDataURL(blob)
  } catch {
    return null
  }
}

async function retouchLineMessage(request: ScanRetouchRequest) {
  try {
    const pixels = new Uint8ClampedArray(request.rgba)
    const common = {
      channels: 4 as const,
      dpi: request.dpi,
      baseline: request.baseline,
      length: request.length,
      fontSize: request.fontSize,
      text: request.text,
    }
    const { style, features } = estimateScanStyle(pixels, request.width, request.height, common)
    const result = retouchLine(pixels, request.width, request.height, {
      ...common,
      xHeight: Number.isFinite(features.xHeight) ? features.xHeight : undefined,
      fontClass: style.fontClass,
      targets: request.targets?.length ? request.targets : undefined,
      segment: !request.targets?.length,
      seed: request.seed,
    })
    const full = composePatch(result, null)
    const dataUrl = full ? await encodePng(full.rgba, full.width, full.height) : null
    const transfer: Transferable[] = [result.filled.buffer as ArrayBuffer, result.labels.buffer as ArrayBuffer, result.blocked.buffer as ArrayBuffer]
    if (full && !dataUrl) transfer.push(full.rgba.buffer as ArrayBuffer)
    scope.postMessage({
      type: 'retouched',
      id: request.id,
      style,
      words: result.words.map((word) => ({ text: word.text, labels: word.labels, box: word.box })),
      segmented: result.segmented,
      inkBox: result.inkBox,
      patch: full
        ? { x: full.x, y: full.y, width: full.width, height: full.height, ...(dataUrl ? { dataUrl } : { rgba: full.rgba.buffer as ArrayBuffer }) }
        : null,
      retouch: {
        width: result.width,
        height: result.height,
        bilevel: result.bilevel,
        paperColor: result.paperColor,
        filled: result.filled.buffer as ArrayBuffer,
        labels: result.labels.buffer as ArrayBuffer,
        blocked: result.blocked.buffer as ArrayBuffer,
      },
    }, { transfer })
  } catch (error) {
    scope.postMessage({ type: 'retouch-error', id: request.id, message: error instanceof Error ? error.message : String(error) })
  }
}

scope.addEventListener('message', (event) => {
  const request = event.data
  if (request?.type === 'prepare-page') void preparePageMessage(request)
  else if (request?.type === 'retouch-line') void retouchLineMessage(request)
})

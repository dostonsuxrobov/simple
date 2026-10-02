// src/shared/worker-ops/index.ts (WP1)
// Routing table and message runtime of the imaging worker. DOM-free: the module worker
// (../imaging.worker.ts), the inline fallback of ../workerClient.ts and Node tests all load it.
//
// Protocol (main thread <-> worker), one job per worker at a time:
//   worker -> main  { type: 'ready' }                                  once, after the module graph loaded
//   main -> worker  { type: 'run', id, op, input }                     input buffers transferred
//   main -> worker  { type: 'abort', id }                              flips the job's AbortSignal
//   worker -> main  { type: 'progress', id, fraction }                 0..1, throttled to 1% steps
//   worker -> main  { type: 'result', id, output }                     output buffers transferred
//   worker -> main  { type: 'error', id, name, message }               name 'AbortError' after an abort
import type { ImagingHandlers, ImagingOp, OpOptions } from '../../imaging/types.ts'
import { handlers as colorHandlers } from './color.ts'
import { handlers as maskHandlers } from './mask.ts'

/** The per-owner handler tables (WP2: color.ts, WP4: mask.ts). Their op sets must be disjoint. */
export const HANDLER_TABLES: Readonly<Record<'color' | 'mask', ImagingHandlers>> = Object.freeze({
  color: colorHandlers,
  mask: maskHandlers,
})

/** Every available handler, keyed by op. */
export const HANDLERS: ImagingHandlers = Object.freeze({ ...colorHandlers, ...maskHandlers })

export type ImagingWorkerRequest =
  | { readonly type: 'run'; readonly id: number; readonly op: ImagingOp; readonly input: unknown }
  | { readonly type: 'abort'; readonly id: number }

export type ImagingWorkerResponse =
  | { readonly type: 'ready' }
  | { readonly type: 'progress'; readonly id: number; readonly fraction: number }
  | { readonly type: 'result'; readonly id: number; readonly output: unknown }
  | { readonly type: 'error'; readonly id: number; readonly name: string; readonly message: string }

export type PostImagingResponse = (message: ImagingWorkerResponse, transfer: Transferable[]) => void

/** The parts of a DedicatedWorkerGlobalScope the runtime uses (tests pass a fake). */
export interface ImagingWorkerScope {
  postMessage(message: ImagingWorkerResponse, transfer: Transferable[]): void
  addEventListener(type: 'message', listener: (event: { readonly data: unknown }) => void): void
}

export interface ImagingWorkerRuntime {
  /** Feed one message received from the main thread. Unknown messages are ignored. */
  handle(message: unknown): void
  /** Jobs currently running. */
  readonly active: number
}

export function namedError(name: string, message: string): Error {
  const error = new Error(message)
  error.name = name
  return error
}

/** ArrayBuffers reachable from plain objects, arrays and typed-array views, deduplicated. */
export function collectTransferables(value: unknown, found: Set<ArrayBuffer> = new Set(), depth = 0): Set<ArrayBuffer> {
  if (!value || typeof value !== 'object' || depth > 8) return found
  if (ArrayBuffer.isView(value)) {
    if (value.buffer instanceof ArrayBuffer && value.buffer.byteLength > 0) found.add(value.buffer)
    return found
  }
  if (value instanceof ArrayBuffer) {
    if (value.byteLength > 0) found.add(value)
    return found
  }
  for (const child of Array.isArray(value) ? value : Object.values(value)) collectTransferables(child, found, depth + 1)
  return found
}

function describeError(error: unknown, aborted: boolean): { name: string; message: string } {
  if (aborted) return { name: 'AbortError', message: 'The operation was cancelled.' }
  if (error instanceof Error) return { name: error.name || 'Error', message: error.message || 'The imaging operation failed.' }
  return { name: 'Error', message: typeof error === 'string' && error ? error : 'The imaging operation failed.' }
}

type AnyHandler = (input: unknown, options: OpOptions) => unknown

export function createImagingWorkerRuntime(handlers: ImagingHandlers, post: PostImagingResponse): ImagingWorkerRuntime {
  const running = new Map<number, AbortController>()

  const run = async (id: number, op: ImagingOp, input: unknown) => {
    const controller = new AbortController()
    running.set(id, controller)
    let reported = -1
    const onProgress = (fraction: number) => {
      if (controller.signal.aborted) return
      const value = Math.min(1, Math.max(0, Number(fraction) || 0))
      if (value < 1 && value - reported < 0.01) return
      if (value <= reported) return
      reported = value
      post({ type: 'progress', id, fraction: value }, [])
    }
    try {
      const handler = handlers[op] as AnyHandler | undefined
      if (typeof handler !== 'function') {
        throw namedError('NotSupportedError', `The imaging operation "${String(op)}" is not available in this build.`)
      }
      const output = await handler(input, { signal: controller.signal, onProgress })
      if (controller.signal.aborted) throw namedError('AbortError', 'The operation was cancelled.')
      post({ type: 'result', id, output }, [...collectTransferables(output)])
    } catch (error) {
      const failure = describeError(error, controller.signal.aborted)
      post({ type: 'error', id, name: failure.name, message: failure.message }, [])
    } finally {
      running.delete(id)
    }
  }

  return {
    get active() {
      return running.size
    },
    handle(message: unknown) {
      if (!message || typeof message !== 'object') return
      const request = message as { type?: unknown; id?: unknown; op?: unknown; input?: unknown }
      if (typeof request.id !== 'number') return
      if (request.type === 'run') void run(request.id, request.op as ImagingOp, request.input)
      else if (request.type === 'abort') running.get(request.id)?.abort()
    },
  }
}

/** Wires a worker global scope to the runtime and announces readiness. */
export function attachImagingWorker(scope: ImagingWorkerScope, handlers: ImagingHandlers = HANDLERS): ImagingWorkerRuntime {
  const runtime = createImagingWorkerRuntime(handlers, (message, transfer) => scope.postMessage(message, transfer))
  scope.addEventListener('message', (event) => runtime.handle(event.data))
  scope.postMessage({ type: 'ready' }, [])
  return runtime
}

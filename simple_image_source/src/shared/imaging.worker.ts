// src/shared/imaging.worker.ts (WP1)
// Module worker entry, created by workerClient.ts with
//   new Worker(new URL('./imaging.worker.ts', import.meta.url), { type: 'module' })
// Vite bundles it as its own chunk (vite.config.ts: worker.format 'es'). Routing and the message
// protocol live in ./worker-ops/index.ts so Node tests can drive them without a real worker.
import { attachImagingWorker } from './worker-ops/index.ts'
import type { ImagingWorkerScope } from './worker-ops/index.ts'

const scope = globalThis as unknown as Partial<ImagingWorkerScope> & { document?: unknown }

// Attach only inside a dedicated worker; importing this file from a window or from Node is inert.
if (typeof scope.postMessage === 'function' && typeof scope.addEventListener === 'function' && typeof scope.document === 'undefined') {
  attachImagingWorker(scope as ImagingWorkerScope)
}

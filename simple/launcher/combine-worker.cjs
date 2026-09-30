const { parentPort, workerData } = require('node:worker_threads')
const { combineFiles } = require('./combine-service.cjs')

combineFiles(workerData, (progress) => parentPort.postMessage({ progress })).then(
  (result) => parentPort.postMessage({ result }),
  (error) => parentPort.postMessage({ error: error.message }),
)

'use strict'

const $ = (selector) => document.querySelector(selector)
const status = $('#status')
const dropTarget = $('#drop-target')
const dropOverlay = $('#drop-overlay')
let dragDepth = 0

function showStatus(message, kind = 'ok') {
  status.textContent = message
  status.dataset.kind = kind
  status.classList.add('visible')
  clearTimeout(showStatus.timer)
  showStatus.timer = setTimeout(() => status.classList.remove('visible'), 4200)
}

async function run(action, successMessage) {
  try {
    const result = await action()
    if (successMessage) showStatus(typeof successMessage === 'function' ? successMessage(result) : successMessage)
    return result
  } catch (error) {
    showStatus(error?.message || 'Something went wrong.', 'error')
    return null
  }
}

$('#open').addEventListener('click', () => run(
  () => window.simpleLauncher.open(),
  (result) => result?.opened ? `Opened ${result.opened} file${result.opened === 1 ? '' : 's'}.` : 'No file selected.',
))

document.querySelectorAll('[data-mode]').forEach((button) => {
  button.addEventListener('click', () => run(
    () => window.simpleLauncher.launchMode(button.dataset.mode),
    'Workspace opened.',
  ))
})

$('#register').addEventListener('click', () => run(
  () => window.simpleLauncher.registerFileTypes(),
  (result) => `simple is available for ${result.extensionCount} file types. Choose your defaults next.`,
))
$('#defaults').addEventListener('click', () => run(() => window.simpleLauncher.openDefaultApps()))
$('#unregister').addEventListener('click', () => run(
  () => window.simpleLauncher.unregisterFileTypes(),
  'simple file support was removed.',
))

$('#minimize').addEventListener('click', () => window.simpleLauncher.minimize())
$('#close').addEventListener('click', () => window.simpleLauncher.close())

dropTarget.addEventListener('dragenter', (event) => {
  event.preventDefault()
  dragDepth += 1
  dropOverlay.hidden = false
})
dropTarget.addEventListener('dragover', (event) => event.preventDefault())
dropTarget.addEventListener('dragleave', (event) => {
  event.preventDefault()
  dragDepth -= 1
  if (dragDepth <= 0) {
    dragDepth = 0
    dropOverlay.hidden = true
  }
})
dropTarget.addEventListener('drop', (event) => {
  event.preventDefault()
  dragDepth = 0
  dropOverlay.hidden = true
  const paths = Array.from(event.dataTransfer.files, (file) => window.simpleLauncher.pathForFile(file)).filter(Boolean)
  void run(
    () => window.simpleLauncher.launchPaths(paths),
    (result) => result?.opened ? `Opened ${result.opened} file${result.opened === 1 ? '' : 's'}.` : 'That file type is not supported.',
  )
})

void window.simpleLauncher.info().then((info) => {
  document.documentElement.dataset.packaged = String(info.packaged)
})

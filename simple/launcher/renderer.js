'use strict'

const $ = (selector) => document.querySelector(selector)
const status = $('#status')
const dropTarget = $('#drop-target')
const dropOverlay = $('#drop-overlay')
let dragDepth = 0
const combineDialog = $('#combine-dialog')
let combineEntries = []
let combining = false

function combineMessage(message, kind = 'ok') {
  $('#combine-message').textContent = message
  $('#combine-message').dataset.kind = kind
}

function renderCombine() {
  const list = $('#combine-list')
  list.replaceChildren()
  combineEntries.forEach((entry, index) => {
    const row = document.createElement('li')
    row.className = 'combine-row'
    const name = document.createElement('div')
    name.className = 'combine-name'
    const strong = document.createElement('strong')
    strong.textContent = `${index + 1}. ${entry.name}`
    strong.title = entry.path
    const detail = document.createElement('small')
    detail.textContent = entry.size > 1024 * 1024 ? `${(entry.size / 1024 / 1024).toFixed(1)} MB` : `${Math.ceil(entry.size / 1024)} KB`
    name.append(strong, detail)
    const pages = document.createElement('input')
    pages.value = entry.pages
    pages.placeholder = 'All pages'
    pages.setAttribute('aria-label', `Pages from ${entry.name}`)
    pages.disabled = combining || /\.(png|jpe?g)$/i.test(entry.path)
    pages.addEventListener('input', () => { entry.pages = pages.value })
    row.append(name, pages)
    const addAction = (label, title, action, disabled = false) => {
      const button = document.createElement('button')
      button.className = 'text-button'
      button.textContent = label
      button.title = title
      button.setAttribute('aria-label', title)
      button.disabled = combining || disabled
      button.addEventListener('click', action)
      row.append(button)
    }
    addAction('Preview', `Preview ${entry.name}`, () => run(() => window.simpleLauncher.launchPaths([entry.path])))
    const move = (offset) => {
      const [item] = combineEntries.splice(index, 1)
      combineEntries.splice(index + offset, 0, item)
      renderCombine()
      $('#combine-list').children[index + offset]?.querySelector('[title^="Preview"]')?.focus()
    }
    addAction('↑', `Move up ${entry.name}`, () => move(-1), index === 0)
    addAction('↓', `Move down ${entry.name}`, () => move(1), index === combineEntries.length - 1)
    addAction('×', `Remove ${entry.name}`, () => { combineEntries.splice(index, 1); renderCombine() })
    list.append(row)
  })
  $('#combine-empty').hidden = combineEntries.length > 0
  $('#combine-count').textContent = `${combineEntries.length} file${combineEntries.length === 1 ? '' : 's'}`
  $('#combine-save').disabled = combining || combineEntries.length < 2
  $('#combine-add').disabled = combining
  $('#combine-close').disabled = false
  $('#combine-save').textContent = combining ? 'Combining…' : 'Save combined PDF'
}

async function addCombineFiles(paths) {
  if (combining) return
  try {
    const entries = await window.simpleLauncher.combineAdd(paths)
    if (combineEntries.length + entries.length > 100) throw new Error('Combine up to 100 files at a time.')
    combineEntries.push(...entries)
    combineMessage('')
    renderCombine()
  } catch (error) { combineMessage(error.message || 'Could not add these files.', 'error') }
}

$('#combine').addEventListener('click', () => { renderCombine(); combineDialog.showModal(); $('#combine-add').focus() })
$('#combine-add').addEventListener('click', () => void addCombineFiles())
$('#combine-close').addEventListener('click', () => combineDialog.close())
combineDialog.addEventListener('dragover', (event) => event.preventDefault())
combineDialog.addEventListener('drop', (event) => {
  event.preventDefault()
  const paths = Array.from(event.dataTransfer.files, (file) => window.simpleLauncher.pathForFile(file)).filter(Boolean)
  void addCombineFiles(paths)
})
$('#combine-save').addEventListener('click', async () => {
  if (combining || combineEntries.length < 2) return
  combining = true
  renderCombine()
  combineMessage('Choose where to save the combined PDF…')
  try {
    const result = await window.simpleLauncher.combineSave(combineEntries)
    if (result.canceled) combineMessage('')
    else {
      const message = `Saved ${result.name} — ${result.pageCount} pages. Opened in PDF for review.`
      combineMessage(message)
      if (!combineDialog.open) showStatus(message)
    }
  } catch (error) {
    combineMessage(error.message || 'Could not combine these files.', 'error')
    if (!combineDialog.open) showStatus(error.message || 'Could not combine these files.', 'error')
  }
  finally { combining = false; renderCombine() }
})
window.simpleLauncher.onCombineProgress(({ index, total, name }) => combineMessage(`Preparing file ${index + 1} of ${total}: ${name}`))

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
document.addEventListener('keydown', (event) => {
  if (event.ctrlKey && event.key.toLowerCase() === 'o' && !combineDialog.open) {
    event.preventDefault()
    $('#open').click()
  }
})

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

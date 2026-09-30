'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { spawn, spawnSync } = require('node:child_process')

function argument(name) {
  const prefix = `--${name}=`
  return process.argv.find((value) => value.startsWith(prefix))?.slice(prefix.length) || null
}

function writeMinimalPdf(targetPath) {
  const content = 'BT /F1 18 Tf 72 720 Td (Simple print smoke) Tj ET\n'
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${Buffer.byteLength(content, 'ascii')} >>\nstream\n${content}endstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ]
  let output = '%PDF-1.4\n'
  const offsets = [0]
  objects.forEach((object, index) => {
    offsets.push(Buffer.byteLength(output, 'ascii'))
    output += `${index + 1} 0 obj\n${object}\nendobj\n`
  })
  const xrefOffset = Buffer.byteLength(output, 'ascii')
  output += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  output += offsets.slice(1).map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`).join('')
  output += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`
  fs.writeFileSync(targetPath, output, 'ascii')
}

function writeVideoFixture(targetPath) {
  const ffmpeg = argument('ffmpeg') || 'ffmpeg'
  const generated = spawnSync(ffmpeg, [
    '-hide_banner', '-loglevel', 'error', '-y',
    '-f', 'lavfi', '-i', 'color=c=0x356a9f:s=320x180:d=1:r=12',
    '-c:v', 'libvpx-vp9', '-pix_fmt', 'yuv420p', '-an', targetPath,
  ], { encoding: 'utf8', windowsHide: true })
  if (generated.error || generated.status !== 0) {
    throw new Error(`Could not generate the WebM print fixture with ffmpeg. ${generated.error?.message || generated.stderr || ''}`.trim())
  }
}

const executable = path.resolve(argument('exe') || path.join(__dirname, '..', 'release', 'simple.exe'))
const firstPort = Number(argument('port') || 9599)
const onlyMode = argument('only')
const supportedModes = ['image', 'docs', 'calc', 'pdf', 'video']
const imagePath = path.resolve(argument('image') || path.join(__dirname, '..', '..', 'simple_image_source', 'public', 'brand-icon.png'))
const documentPath = path.resolve(argument('doc') || path.join(__dirname, '..', '..', 'simple_doc_source', 'qa', 'fixtures', 'simple-docs-roundtrip-fixture.docx'))
const spreadsheetPath = path.resolve(argument('sheet') || path.join(__dirname, '..', '..', 'simple_calc_source', 'scripts', 'fixtures', 'print-smoke.csv'))
const generatedPdfDirectory = argument('pdf') ? null : fs.mkdtempSync(path.join(os.tmpdir(), 'simple-pdf-print-smoke-'))
const pdfPath = path.resolve(argument('pdf') || path.join(generatedPdfDirectory, 'print-smoke.pdf'))
const generatedVideoDirectory = argument('video') || (onlyMode && onlyMode !== 'video') ? null : fs.mkdtempSync(path.join(os.tmpdir(), 'simple-video-print-smoke-'))
const videoPath = argument('video') ? path.resolve(argument('video')) : generatedVideoDirectory ? path.join(generatedVideoDirectory, 'print-smoke.webm') : null
if (generatedPdfDirectory && (!onlyMode || onlyMode === 'pdf')) writeMinimalPdf(pdfPath)
if (generatedVideoDirectory) writeVideoFixture(videoPath)

assert.ok(!onlyMode || supportedModes.includes(onlyMode), `Unknown --only mode: ${onlyMode}`)
const fixtures = [
  ['packaged executable', executable, true],
  ['image fixture', imagePath, !onlyMode || onlyMode === 'image'],
  ['DOCX fixture', documentPath, !onlyMode || onlyMode === 'docs'],
  ['spreadsheet fixture', spreadsheetPath, !onlyMode || onlyMode === 'calc'],
  ['PDF fixture', pdfPath, !onlyMode || onlyMode === 'pdf'],
  ['video fixture', videoPath, !onlyMode || onlyMode === 'video'],
]
for (const [label, filePath, required] of fixtures) {
  if (!required) continue
  assert.ok(fs.existsSync(filePath) && fs.statSync(filePath).isFile(), `Missing ${label}: ${filePath}`)
}
assert.ok(Number.isInteger(firstPort) && firstPort >= 1024 && firstPort <= 65_525, `Invalid first debugging port: ${firstPort}`)

const delay = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

async function targets(port) {
  return (await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()).filter((item) => item.type === 'page')
}

async function workspaceTarget(port, mode) {
  const started = Date.now()
  while (Date.now() - started < 60_000) {
    try {
      const target = (await targets(port)).find((item) => item.url.includes(`/modules/${mode}/dist/index.html`))
      if (target?.webSocketDebuggerUrl) return target
    } catch {}
    await delay(150)
  }
  throw new Error(`The packaged ${mode} workspace did not start.`)
}

class CdpClient {
  constructor(url) {
    this.url = url
    this.socket = null
    this.nextId = 0
    this.pending = new Map()
  }

  async connect() {
    this.socket = new WebSocket(this.url)
    await new Promise((resolve, reject) => {
      this.socket.addEventListener('open', resolve, { once: true })
      this.socket.addEventListener('error', reject, { once: true })
    })
    this.socket.addEventListener('message', (event) => {
      const message = JSON.parse(String(event.data))
      if (!message.id || !this.pending.has(message.id)) return
      const { resolve, reject } = this.pending.get(message.id)
      this.pending.delete(message.id)
      if (message.error) reject(new Error(message.error.message))
      else resolve(message.result)
    })
    await this.call('Runtime.enable')
  }

  call(method, params = {}) {
    const id = ++this.nextId
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject })
      this.socket.send(JSON.stringify({ id, method, params }))
    })
  }

  async evaluate(expression) {
    const response = await this.call('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true, userGesture: true })
    if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text)
    return response.result.value
  }

  async waitFor(expression, timeout = 30_000) {
    const started = Date.now()
    while (Date.now() - started < timeout) {
      try { if (await this.evaluate(expression)) return } catch {}
      await delay(100)
    }
    throw new Error(`Timed out waiting for: ${expression}`)
  }

  close() {
    try { this.socket?.close() } catch {}
  }
}

async function triggerPrint(client, mode, trigger) {
  if (mode === 'video') {
    await client.waitFor(`document.querySelector('video')?.readyState >= 2 && document.querySelector('video')?.videoWidth > 0 && Boolean(document.querySelector('.print-toolbar-button'))`, 45_000)
    if (trigger === 'toolbar') {
      const visible = await client.evaluate(`(() => { const button = document.querySelector('.print-toolbar-button'); const style = getComputedStyle(button); return !button.disabled && style.display !== 'none' && style.visibility !== 'hidden' && button.getBoundingClientRect().width > 0 })()`)
      assert.equal(visible, true, 'The Video Print toolbar button is not discoverable.')
      await client.evaluate(`document.querySelector('.print-toolbar-button').click()`)
    } else {
      await client.evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', code: 'KeyP', ctrlKey: true, bubbles: true, cancelable: true }))`)
    }
    return
  }

  if (mode === 'image') {
    await client.waitFor(`document.querySelector('canvas')?.width > 0 && document.querySelector('button[title^="Print"]') && !document.querySelector('button[title^="Print"]').disabled`)
    if (trigger === 'toolbar') {
      const visible = await client.evaluate(`(() => { const button = document.querySelector('button[title^="Print"]'); const style = getComputedStyle(button); return style.display !== 'none' && style.visibility !== 'hidden' && button.getBoundingClientRect().width > 0 })()`)
      assert.equal(visible, true, 'The image Print toolbar button is not discoverable.')
      await client.evaluate(`document.querySelector('button[title^="Print"]').click()`)
    } else {
      await client.evaluate(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', ctrlKey: true, bubbles: true, cancelable: true }))`)
    }
    return
  }

  if (mode === 'calc') {
    try {
      await client.waitFor(`Boolean(document.querySelector('button[aria-label="Print (Ctrl+P)"]'))`)
    } catch (error) {
      const state = await client.evaluate(`({ title: document.title, text: document.body.innerText.slice(0, 1200), buttons: Array.from(document.querySelectorAll('button')).map((button) => button.getAttribute('aria-label') || button.title || button.textContent.trim()).filter(Boolean).slice(0, 80) })`)
      throw new Error(`${error.message}\nCalc renderer state: ${JSON.stringify(state)}`)
    }
    if (trigger === 'toolbar') {
      const visible = await client.evaluate(`(() => { const button = document.querySelector('button[aria-label="Print (Ctrl+P)"]'); const style = getComputedStyle(button); return style.display !== 'none' && style.visibility !== 'hidden' && button.getBoundingClientRect().width > 0 })()`)
      assert.equal(visible, true, 'The Calc Print toolbar button is not discoverable.')
      await client.evaluate(`document.querySelector('button[aria-label="Print (Ctrl+P)"]').click()`)
    } else {
      await client.evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', ctrlKey: true, bubbles: true, cancelable: true }))`)
    }
    return
  }

  if (mode === 'pdf') {
    try {
      await client.waitFor(`Boolean(document.querySelector('.page-canvas')?.width > 300 && document.querySelector('button[aria-label="Print (Ctrl+P)"]'))`)
    } catch (error) {
      const state = await client.evaluate(`({ title: document.title, text: document.body.innerText.slice(0, 1200), canvas: (() => { const canvas = document.querySelector('.page-canvas'); return canvas ? { width: canvas.width, height: canvas.height, cssWidth: canvas.getBoundingClientRect().width } : null })(), buttons: Array.from(document.querySelectorAll('button')).map((button) => button.getAttribute('aria-label') || button.title || button.textContent.trim()).filter(Boolean).slice(0, 80) })`)
      throw new Error(`${error.message}\nPDF renderer state: ${JSON.stringify(state)}`)
    }
    if (trigger === 'toolbar') {
      const visible = await client.evaluate(`(() => { const button = document.querySelector('button[aria-label="Print (Ctrl+P)"]'); const style = getComputedStyle(button); return style.display !== 'none' && style.visibility !== 'hidden' && button.getBoundingClientRect().width > 0 })()`)
      assert.equal(visible, true, 'The PDF Print toolbar button is not discoverable.')
      await client.evaluate(`document.querySelector('button[aria-label="Print (Ctrl+P)"]').click()`)
    } else {
      await client.evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', ctrlKey: true, bubbles: true, cancelable: true }))`)
    }
    return
  }

  await client.waitFor(`document.querySelector('#welcome')?.hidden && document.querySelector('#loading-overlay')?.hidden && document.querySelector('#print-button')`)
  if (trigger === 'toolbar') {
    await client.evaluate(`document.querySelector('#more-button').click()`)
    const visible = await client.evaluate(`(() => { const button = document.querySelector('#print-button'); const style = getComputedStyle(button); return !button.closest('[hidden]') && style.display !== 'none' && style.visibility !== 'hidden' && button.getBoundingClientRect().width > 0 })()`)
    assert.equal(visible, true, 'The Docs Print menu command is not discoverable.')
    await client.evaluate(`document.querySelector('#print-button').click()`)
  } else {
    await client.evaluate(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', ctrlKey: true, bubbles: true, cancelable: true }))`)
  }
}

function assertTwoPaneAudit(mode, audit) {
  assert.equal(audit.visible, true, `The ${mode} print dialog is not visible.`)
  assert.equal(audit.twoPane, true, `The ${mode} print options and live preview are not side by side.`)
  assert.equal(audit.contained, true, `The ${mode} print dialog extends outside the packaged window.`)
  assert.ok(audit.leftWidth >= 250, `The ${mode} print options pane is too narrow (${audit.leftWidth}px).`)
  assert.ok(audit.rightWidth >= 340, `The ${mode} live preview pane is too narrow (${audit.rightWidth}px).`)
  assert.equal(audit.controls, true, `The ${mode} print dialog is missing required page-layout controls.`)
  assert.equal(audit.preview, true, `The ${mode} print dialog did not produce a functional live preview.`)
  assert.equal(audit.submitReady, true, `The ${mode} print action did not become ready after preview rendering.`)
}

async function assertPrintDialog(client, mode) {
  if (mode === 'video') {
    await client.waitFor(`Boolean(document.querySelector('.video-print-dialog'))`, 45_000)
    await client.waitFor(`Boolean(document.querySelector('.video-print-preview[aria-busy="false"]') && document.querySelector('.video-print-preview-frame')?.srcdoc.length > 500 && !document.querySelector('.video-print-action')?.disabled)`, 45_000)
    await client.evaluate(`(() => {
      const change = (label, value) => { const control = document.querySelector('[aria-label="' + label + '"]'); control.value = value; control.dispatchEvent(new Event('change', { bubbles: true })); };
      change('Print paper size', 'a4');
      change('Print orientation', 'portrait');
      change('Print margins', 'narrow');
      change('Print scaling', 'fill');
      change('Print color mode', 'grayscale');
    })()`)
    await client.waitFor(`document.querySelector('.video-print-preview[aria-busy="false"]') && document.querySelector('.video-print-preview-frame')?.srcdoc.includes('grayscale(1)') && document.querySelector('.video-print-preview-toolbar')?.textContent.includes('A4') && document.querySelector('.video-print-preview-toolbar')?.textContent.includes('Portrait') && !document.querySelector('.video-print-action')?.disabled`, 45_000)
    const audit = await client.evaluate(`(() => {
      const root = document.querySelector('.video-print-dialog');
      const left = document.querySelector('.video-print-controls');
      const right = document.querySelector('.video-print-preview');
      const rootRect = root.getBoundingClientRect();
      const leftRect = left.getBoundingClientRect();
      const rightRect = right.getBoundingClientRect();
      const frame = document.querySelector('.video-print-preview-frame');
      return {
        visible: rootRect.width > 0,
        twoPane: leftRect.right <= rightRect.left + 1 && leftRect.top < rightRect.bottom && rightRect.top < leftRect.bottom,
        contained: rootRect.left >= 0 && rootRect.top >= 0 && rootRect.right <= innerWidth + 1 && rootRect.bottom <= innerHeight + 1,
        leftWidth: Math.round(leftRect.width), rightWidth: Math.round(rightRect.width),
        controls: document.querySelector('[aria-label="Print paper size"]').value === 'a4'
          && document.querySelector('[aria-label="Print orientation"]').value === 'portrait'
          && document.querySelector('[aria-label="Print margins"]').value === 'narrow'
          && document.querySelector('[aria-label="Print scaling"]').value === 'fill'
          && document.querySelector('[aria-label="Print color mode"]').value === 'grayscale',
        preview: frame?.srcdoc.length > 500 && frame.srcdoc.includes('grayscale(1)')
          && document.querySelector('.video-print-preview').getAttribute('aria-busy') === 'false'
          && document.querySelector('.video-print-preview-toolbar').textContent.includes('A4')
          && document.querySelector('.video-print-warning')?.textContent.includes('outside the printable area'),
        submitReady: !document.querySelector('.video-print-action').disabled,
      };
    })()`)
    assertTwoPaneAudit(mode, audit)
    return audit
  }

  if (mode === 'docs') {
    await client.waitFor(`!document.querySelector('#print-modal')?.hidden`)
    await client.waitFor(`Boolean(!document.querySelector('#print-submit')?.disabled && !document.querySelector('#print-preview-pdf')?.hidden && document.querySelector('#print-preview-pdf')?.src.startsWith('blob:'))`, 45_000)
    await client.evaluate(`(() => {
      const change = (selector, value) => { const control = document.querySelector(selector); control.value = value; control.dispatchEvent(new Event('input', { bubbles: true })); control.dispatchEvent(new Event('change', { bubbles: true })); };
      change('#print-paper', 'A4');
      document.querySelector('input[name="print-orientation"][value="landscape"]').click();
      change('#print-margins', 'normal');
      document.querySelector('input[name="print-pages"][value="custom"]').click();
      change('#print-page-range', '1');
      if (!document.querySelector('#print-title-checkbox').checked) document.querySelector('#print-title-checkbox').click();
      if (!document.querySelector('#print-page-numbers').checked) document.querySelector('#print-page-numbers').click();
    })()`)
    await client.waitFor(`document.querySelector('#print-preview-status')?.textContent.startsWith('1 page') && !document.querySelector('#print-submit')?.disabled`, 45_000)
    const audit = await client.evaluate(`(() => {
      const root = document.querySelector('.print-dialog');
      const left = document.querySelector('.print-options');
      const right = document.querySelector('.print-preview-pane');
      const rootRect = root.getBoundingClientRect();
      const leftRect = left.getBoundingClientRect();
      const rightRect = right.getBoundingClientRect();
      const preview = document.querySelector('#print-preview-pdf');
      return {
        visible: !document.querySelector('#print-modal').hidden && rootRect.width > 0,
        twoPane: leftRect.right <= rightRect.left + 1 && leftRect.top < rightRect.bottom && rightRect.top < leftRect.bottom,
        contained: rootRect.left >= 0 && rootRect.top >= 0 && rootRect.right <= innerWidth + 1 && rootRect.bottom <= innerHeight + 1,
        leftWidth: Math.round(leftRect.width), rightWidth: Math.round(rightRect.width),
        controls: document.querySelector('#print-paper').value === 'A4'
          && document.querySelector('input[name="print-orientation"]:checked').value === 'landscape'
          && document.querySelector('#print-margins').value === 'normal'
          && document.querySelector('#print-page-range').value === '1'
          && document.querySelectorAll('input[name="print-pages"]').length >= 2
          && document.querySelectorAll('#print-title-checkbox, #print-page-numbers').length === 2,
        preview: preview.src.startsWith('blob:') && !preview.hidden
          && document.querySelector('#print-preview-status').textContent.startsWith('1 page')
          && document.querySelector('#print-preview-stage').getAttribute('aria-busy') === 'false',
        submitReady: !document.querySelector('#print-submit').disabled,
      };
    })()`)
    assertTwoPaneAudit(mode, audit)
    return audit
  }

  if (mode === 'calc') {
    await client.waitFor(`document.querySelector('#print-dialog-title')?.textContent.includes('Print spreadsheet')`)
    await client.waitFor(`Boolean(document.querySelector('.print-preview-frame')?.srcdoc.length > 500 && document.querySelector('.print-preview-pane')?.getAttribute('aria-busy') === 'false' && !document.querySelector('.print-controls-actions .primary-action')?.disabled)`, 45_000)
    await client.evaluate(`(() => {
      const saved = document.querySelector('[aria-label="Use saved page layout"]');
      if (saved?.checked) saved.click();
    })()`)
    await client.waitFor(`!document.querySelector('[aria-label="Print paper size"]')?.matches(':disabled')`)
    await client.evaluate(`(() => {
      const change = (label, value) => { const control = document.querySelector('[aria-label="' + label + '"]'); control.value = value; control.dispatchEvent(new Event('change', { bubbles: true })); };
      change('Print paper size', 'a4');
      change('Print orientation', 'landscape');
      change('Print margins', 'narrow');
      change('Print scaling', 'fit-sheet');
      const headings = Array.from(document.querySelectorAll('.print-detail-options label')).find((label) => label.textContent.includes('Row and column headings'))?.querySelector('input');
      if (headings && !headings.checked) headings.click();
    })()`)
    await client.waitFor(`document.querySelector('.print-preview-summary')?.textContent.includes('A4') && document.querySelector('.print-preview-summary')?.textContent.includes('Landscape') && document.querySelector('.print-preview-pane')?.getAttribute('aria-busy') === 'false' && !document.querySelector('.print-controls-actions .primary-action')?.disabled`, 45_000)
    const audit = await client.evaluate(`(() => {
      const root = document.querySelector('.print-layout-dialog');
      const left = document.querySelector('.print-controls-pane');
      const right = document.querySelector('.print-preview-pane');
      const rootRect = root.getBoundingClientRect();
      const leftRect = left.getBoundingClientRect();
      const rightRect = right.getBoundingClientRect();
      const frame = document.querySelector('.print-preview-frame');
      return {
        visible: rootRect.width > 0 && document.querySelector('#print-dialog-title').textContent.includes('Print spreadsheet'),
        twoPane: leftRect.right <= rightRect.left + 1 && leftRect.top < rightRect.bottom && rightRect.top < leftRect.bottom,
        contained: rootRect.left >= 0 && rootRect.top >= 0 && rootRect.right <= innerWidth + 1 && rootRect.bottom <= innerHeight + 1,
        leftWidth: Math.round(leftRect.width), rightWidth: Math.round(rightRect.width),
        controls: document.querySelector('[aria-label="Print scope"]')
          && document.querySelector('[aria-label="Print paper size"]').value === 'a4'
          && document.querySelector('[aria-label="Print orientation"]').value === 'landscape'
          && document.querySelector('[aria-label="Print margins"]').value === 'narrow'
          && document.querySelector('[aria-label="Print scaling"]').value === 'fit-sheet'
          && document.querySelectorAll('.print-detail-options input[type="checkbox"]').length >= 2,
        preview: frame?.srcdoc.length > 500
          && document.querySelector('.print-preview-summary').textContent.includes('A4')
          && document.querySelector('.print-preview-summary').textContent.includes('Landscape')
          && document.querySelector('.print-preview-pane').getAttribute('aria-busy') === 'false',
        submitReady: !document.querySelector('.print-controls-actions .primary-action').disabled,
      };
    })()`)
    audit.controls = Boolean(audit.controls)
    assertTwoPaneAudit(mode, audit)
    return audit
  }

  if (mode === 'pdf') {
    await client.waitFor(`Boolean(document.querySelector('.print-dialog.print-layout-dialog[aria-label="Print"]'))`, 45_000)
    await client.waitFor(`Boolean(document.querySelector('.print-page-image')?.complete && document.querySelector('.print-page-image')?.naturalWidth > 0 && !document.querySelector('.print-dialog-footer .button-primary')?.disabled)`, 45_000)
    await client.evaluate(`(() => {
      const field = (name) => Array.from(document.querySelectorAll('.print-options-pane .print-field')).find((label) => label.querySelector(':scope > span')?.textContent.trim() === name)?.querySelector('select');
      const change = (name, value) => { const control = field(name); control.value = value; control.dispatchEvent(new Event('change', { bubbles: true })); };
      change('Paper', 'A4');
      change('Orientation', 'landscape');
      change('Margins', 'minimum');
      change('Color', 'bw');
      change('Page sizing', 'actual');
    })()`)
    await client.waitFor(`document.querySelector('.print-preview-toolbar')?.textContent.includes('A4 landscape') && document.querySelector('.print-page-image')?.complete && document.querySelector('.print-page-image')?.naturalWidth > 0`, 45_000)
    const audit = await client.evaluate(`(() => {
      const root = document.querySelector('.print-dialog.print-layout-dialog');
      const left = document.querySelector('.print-options-pane');
      const right = document.querySelector('.print-preview-pane');
      const rootRect = root.getBoundingClientRect();
      const leftRect = left.getBoundingClientRect();
      const rightRect = right.getBoundingClientRect();
      const labels = Array.from(left.querySelectorAll('.print-field > span')).map((span) => span.textContent.trim());
      const submit = document.querySelector('.print-dialog-footer .button-primary');
      return {
        visible: rootRect.width > 0,
        twoPane: leftRect.right <= rightRect.left + 1 && leftRect.top < rightRect.bottom && rightRect.top < leftRect.bottom,
        contained: rootRect.left >= 0 && rootRect.top >= 0 && rootRect.right <= innerWidth + 1 && rootRect.bottom <= innerHeight + 1,
        leftWidth: Math.round(leftRect.width), rightWidth: Math.round(rightRect.width),
        controls: ['Printer', 'Copies', 'Pages', 'Paper', 'Orientation', 'Margins', 'Color', 'Page sizing'].every((name) => labels.includes(name))
          && left.querySelectorAll('input[name="print-range"]').length >= 3,
        preview: document.querySelector('.print-page-image')?.naturalWidth > 0
          && document.querySelector('.print-preview-toolbar').textContent.includes('A4 landscape')
          && document.querySelector('.print-preview-summary').textContent.includes('PDF page')
          && !document.body.innerText.includes('System preview'),
        submitReady: Boolean(submit && submit.textContent.trim() === 'Print' && !submit.disabled),
      };
    })()`)
    assertTwoPaneAudit(mode, audit)
    return audit
  }

  await client.waitFor(`Boolean(document.querySelector('.image-print-dialog'))`, 45_000)
  await client.evaluate(`(() => {
    const change = (selector, value) => { const control = document.querySelector(selector); control.value = value; control.dispatchEvent(new Event('change', { bubbles: true })); };
    change('[data-print-setting="paper"]', 'a4');
    change('[data-print-setting="orientation"]', 'landscape');
    change('[data-print-setting="margin-preset"]', '6.35');
    change('[data-print-setting="scale-mode"]', 'fill');
    change('[data-print-setting="position"]', 'top-right');
    const grayscale = document.querySelector('[data-print-setting="grayscale"]');
    if (!grayscale.checked) grayscale.click();
  })()`)
  await client.waitFor(`Boolean(document.querySelector('.print-preview-sheet[data-paper="a4"][data-orientation="landscape"]') && document.querySelector('.print-preview-image-frame img')?.complete && document.querySelector('.print-preview-image-frame img')?.naturalWidth > 0 && !document.querySelector('.image-print-dialog .print-submit')?.disabled)`, 45_000)
  const audit = await client.evaluate(`(() => {
    const root = document.querySelector('.image-print-dialog');
    const left = document.querySelector('.image-print-controls');
    const right = document.querySelector('.image-print-preview');
    const rootRect = root.getBoundingClientRect();
    const leftRect = left.getBoundingClientRect();
    const rightRect = right.getBoundingClientRect();
    const sheet = document.querySelector('.print-preview-sheet');
    const image = document.querySelector('.print-preview-image-frame img');
    return {
      visible: rootRect.width > 0,
      twoPane: leftRect.right <= rightRect.left + 1 && leftRect.top < rightRect.bottom && rightRect.top < leftRect.bottom,
      contained: rootRect.left >= 0 && rootRect.top >= 0 && rootRect.right <= innerWidth + 1 && rootRect.bottom <= innerHeight + 1,
      leftWidth: Math.round(leftRect.width), rightWidth: Math.round(rightRect.width),
      controls: document.querySelector('[data-print-setting="paper"]').value === 'a4'
        && document.querySelector('[data-print-setting="orientation"]').value === 'landscape'
        && document.querySelector('[data-print-setting="margin-preset"]').value === '6.35'
        && document.querySelector('[data-print-setting="scale-mode"]').value === 'fill'
        && document.querySelector('[data-print-setting="position"]').value === 'top-right'
        && document.querySelector('[data-print-setting="grayscale"]').checked,
      preview: sheet?.dataset.paper === 'a4' && sheet?.dataset.orientation === 'landscape'
        && sheet.getBoundingClientRect().width > 200 && image?.naturalWidth > 0
        && document.querySelector('.print-preview-summary').textContent.includes('Effective resolution'),
      submitReady: !document.querySelector('.image-print-dialog .print-submit').disabled,
    };
  })()`)
  assertTwoPaneAudit(mode, audit)
  return audit
}

async function probe(mode, filePath, trigger, port) {
  const profileDirectory = fs.mkdtempSync(path.join(os.tmpdir(), `simple-${mode}-print-profile-`))
  const child = spawn(executable, [`--remote-debugging-port=${port}`, `--user-data-dir=${profileDirectory}`, filePath], {
    env: { ...process.env, APPDATA: profileDirectory },
    stdio: 'ignore',
    windowsHide: true,
  })
  let client
  try {
    const workspace = await workspaceTarget(port, mode)
    client = new CdpClient(workspace.webSocketDebuggerUrl)
    await client.connect()
    await triggerPrint(client, mode, trigger)
    const audit = await assertPrintDialog(client, mode)
    process.stdout.write(`${mode} ${trigger}: verified safe two-pane preview (${audit.leftWidth}px + ${audit.rightWidth}px); final submission intentionally skipped.\n`)
  } finally {
    if (client) {
      try { await Promise.race([client.call('Browser.close'), delay(1_000)]) } catch {}
      client.close()
    }
    spawnSync('taskkill.exe', ['/PID', String(child.pid), '/T', '/F'], { stdio: 'ignore', windowsHide: true })
    await delay(300)
    try { fs.rmSync(profileDirectory, { recursive: true, force: true }) } catch {}
  }
}

async function main() {
  const probes = [
    ['image', imagePath, 'toolbar'],
    ['image', imagePath, 'shortcut'],
    ['docs', documentPath, 'toolbar'],
    ['docs', documentPath, 'shortcut'],
    ['calc', spreadsheetPath, 'toolbar'],
    ['calc', spreadsheetPath, 'shortcut'],
    ['pdf', pdfPath, 'toolbar'],
    ['pdf', pdfPath, 'shortcut'],
    ['video', videoPath, 'toolbar'],
    ['video', videoPath, 'shortcut'],
  ].filter(([mode]) => !onlyMode || mode === onlyMode)
  for (let index = 0; index < probes.length; index += 1) {
    const [mode, filePath, trigger] = probes[index]
    await probe(mode, filePath, trigger, firstPort + index)
  }
}

main().catch((error) => {
  console.error(error.stack || error)
  process.exitCode = 1
}).finally(() => {
  if (generatedPdfDirectory) {
    try { fs.rmSync(generatedPdfDirectory, { recursive: true, force: true }) } catch {}
  }
  if (generatedVideoDirectory) {
    try { fs.rmSync(generatedVideoDirectory, { recursive: true, force: true }) } catch {}
  }
})

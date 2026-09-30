'use strict'

const assert = require('node:assert/strict')
const crypto = require('node:crypto')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { spawn, spawnSync } = require('node:child_process')

const projectDirectory = path.resolve(__dirname, '..')
const electronPath = require('electron')
const port = Number(process.env.SIMPLE_VIDEO_QA_PORT || 9875)

function delay(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds))
}

async function waitForDebugger() {
  const deadline = Date.now() + 20_000
  while (Date.now() < deadline) {
    try {
      const pages = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
      const page = pages.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
      if (page) return page
    } catch {}
    await delay(125)
  }
  throw new Error('Timed out waiting for the Simple Video renderer.')
}

async function connectCdp(url) {
  const socket = new WebSocket(url)
  const pending = new Map()
  let sequence = 0
  await new Promise((resolve, reject) => {
    socket.addEventListener('open', resolve, { once: true })
    socket.addEventListener('error', reject, { once: true })
  })
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(String(event.data))
    if (!message.id || !pending.has(message.id)) return
    const { resolve, reject } = pending.get(message.id)
    pending.delete(message.id)
    if (message.error) reject(new Error(message.error.message))
    else resolve(message.result)
  })
  return {
    close: () => socket.close(),
    send(method, params = {}) {
      const id = ++sequence
      return new Promise((resolve, reject) => {
        pending.set(id, { resolve, reject })
        socket.send(JSON.stringify({ id, method, params }))
      })
    },
  }
}

async function sha256(filePath) {
  return crypto.createHash('sha256').update(await fs.readFile(filePath)).digest('hex')
}

async function main() {
  const temporaryRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-video-export-'))
  const fixturePath = path.join(temporaryRoot, 'export fixture.webm')
  const exportDirectory = path.join(temporaryRoot, 'exports')
  const printDirectory = path.join(temporaryRoot, 'prints')
  const qaArtifactDirectory = path.join(projectDirectory, 'tmp', 'video-print-qa')
  const profileDirectory = path.join(temporaryRoot, 'profile')
  let child
  let cdp
  let stderr = ''

  try {
    const generated = spawnSync('ffmpeg', [
      '-hide_banner', '-loglevel', 'error', '-y',
      '-f', 'lavfi', '-i', 'color=c=0x356a9f:s=320x180:d=1:r=12',
      '-f', 'lavfi', '-i', 'color=c=0xe63020:s=320x180:d=1:r=12',
      '-filter_complex', '[0:v][1:v]concat=n=2:v=1:a=0[v]', '-map', '[v]',
      '-c:v', 'libvpx-vp9', '-pix_fmt', 'yuv420p', '-an', fixturePath,
    ], { encoding: 'utf8' })
    if (generated.error || generated.status !== 0) {
      throw new Error(`Could not generate the WebM QA fixture. ${generated.error?.message || generated.stderr || ''}`.trim())
    }
    await fs.mkdir(printDirectory, { recursive: true })
    await fs.mkdir(qaArtifactDirectory, { recursive: true })

    child = spawn(electronPath, [
      `--remote-debugging-port=${port}`,
      `--user-data-dir=${profileDirectory}`,
      projectDirectory,
      fixturePath,
    ], {
      cwd: projectDirectory,
      env: {
        ...process.env,
        ELECTRON_DISABLE_SECURITY_WARNINGS: 'true',
        SIMPLE_VIDEO_QA_EXPORT_DIRECTORY: exportDirectory,
        SIMPLE_VIDEO_QA_PRINT_DIRECTORY: printDirectory,
        SIMPLE_VIDEO_QA_PRINT_DELAY_MS: '350',
      },
      stdio: ['ignore', 'ignore', 'pipe'],
      windowsHide: true,
    })
    child.stderr.on('data', (chunk) => { stderr = `${stderr}${chunk}`.slice(-12_000) })

    const page = await waitForDebugger()
    cdp = await connectCdp(page.webSocketDebuggerUrl)
    await cdp.send('Runtime.enable')

    const evaluate = async (expression) => {
      const response = await cdp.send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true })
      if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text)
      return response.result?.value
    }
    const waitFor = async (expression, label) => {
      const deadline = Date.now() + 15_000
      while (Date.now() < deadline) {
        if (await evaluate(expression)) return
        await delay(100)
      }
      throw new Error(`Timed out waiting for ${label}.`)
    }
    const openExport = async () => {
      await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'E', code: 'KeyE', ctrlKey: true, shiftKey: true, bubbles: true }))`)
      await waitFor(`Boolean(document.querySelector('.export-modal'))`, 'the Export As dialog')
    }
    const clickExport = async (label) => {
      await evaluate(`(() => {
        const button = [...document.querySelectorAll('.export-options button')].find((item) => item.textContent.includes(${JSON.stringify(label)}))
        if (!button || button.disabled) throw new Error(${JSON.stringify(`${label} export is unavailable.`)})
        button.click()
        return true
      })()`)
      await waitFor(`!document.querySelector('.export-modal')`, `${label} export to complete`)
    }

    await waitFor(`document.querySelector('video')?.readyState >= 2 && document.querySelector('video')?.videoWidth === 320`, 'the decoded video frame')
    const toolbar = await evaluate(`(() => {
      const button = [...document.querySelectorAll('.top-toolbar button')].find((item) => item.textContent.includes('Export As'))
      return { exists: Boolean(button), visible: Boolean(button && button.getBoundingClientRect().width && button.getBoundingClientRect().height), title: button?.title || '' }
    })()`)
    assert.deepEqual(toolbar, { exists: true, visible: true, title: 'Export As (Ctrl+Shift+E)' })

    const printToolbar = await evaluate(`(() => {
      const button = [...document.querySelectorAll('.top-toolbar button')].find((item) => item.textContent.includes('Print'))
      return { exists: Boolean(button), visible: Boolean(button && button.getBoundingClientRect().width && button.getBoundingClientRect().height), title: button?.title || '' }
    })()`)
    assert.deepEqual(printToolbar, { exists: true, visible: true, title: 'Print current frame (Ctrl+P)' })

    await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', code: 'KeyP', ctrlKey: true, bubbles: true, cancelable: true }))`)
    await waitFor(`Boolean(document.querySelector('.video-print-dialog'))`, 'the current-frame print dialog')
    await waitFor(`document.querySelector('.video-print-preview-frame')?.contentDocument?.querySelector('.print-page') && document.querySelector('.video-print-preview-frame')?.contentDocument?.querySelector('img')?.complete && document.querySelector('.video-print-preview')?.getAttribute('aria-busy') === 'false'`, 'the live video-frame paper preview')
    const printDialog = await evaluate(`(() => ({
      paper: document.querySelector('[aria-label="Print paper size"]')?.value,
      orientation: document.querySelector('[aria-label="Print orientation"]')?.value,
      margins: document.querySelector('[aria-label="Print margins"]')?.value,
      scaling: document.querySelector('[aria-label="Print scaling"]')?.value,
      color: document.querySelector('[aria-label="Print color mode"]')?.value,
      metadata: document.querySelector('.video-print-check input')?.checked,
      previewHasPage: document.querySelector('.video-print-preview-frame')?.srcdoc.includes('class="print-page"'),
      previewHasFrame: document.querySelector('.video-print-preview-frame')?.srcdoc.includes('data:image/png;base64,'),
      previewHasPrintCss: document.querySelector('.video-print-preview-frame')?.srcdoc.includes('@media print'),
      controlWidth: Math.round(document.querySelector('.video-print-controls').getBoundingClientRect().width),
      previewWidth: Math.round(document.querySelector('.video-print-preview').getBoundingClientRect().width),
      actionDisabled: document.querySelector('.video-print-action')?.disabled,
      source: document.querySelector('.video-print-source')?.textContent.replace(/\s+/g, ' ').trim(),
    }))()`)
    assert.equal(printDialog.paper, 'letter')
    assert.equal(printDialog.orientation, 'landscape')
    assert.equal(printDialog.margins, 'normal')
    assert.equal(printDialog.scaling, 'fit')
    assert.equal(printDialog.color, 'color')
    assert.equal(printDialog.metadata, true)
    assert.equal(printDialog.previewHasPage, true)
    assert.equal(printDialog.previewHasFrame, true)
    assert.equal(printDialog.previewHasPrintCss, true)
    assert.equal(printDialog.actionDisabled, false)
    assert.match(printDialog.source, /320 × 180px/)
    assert(printDialog.previewWidth > printDialog.controlWidth, 'the paper preview must occupy the second pane')

    const printShortcutBlocked = await evaluate(`(() => {
      const video = document.querySelector('video')
      const before = video.currentTime
      const event = new KeyboardEvent('keydown', { key: 'ArrowRight', code: 'ArrowRight', bubbles: true, cancelable: true })
      window.dispatchEvent(event)
      return { prevented: event.defaultPrevented, before, after: video.currentTime }
    })()`)
    assert.equal(printShortcutBlocked.prevented, true)
    assert.equal(printShortcutBlocked.after, printShortcutBlocked.before)

    const printFocusWrap = await evaluate(`(() => {
      const dialog = document.querySelector('.video-print-dialog')
      const controls = [...dialog.querySelectorAll('button:not([disabled]), select:not([disabled]), input:not([disabled])')].filter((item) => item.getClientRects().length)
      controls[0].focus()
      controls[0].dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', code: 'Tab', shiftKey: true, bubbles: true, cancelable: true }))
      const reverse = document.activeElement === controls.at(-1)
      controls.at(-1).dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', code: 'Tab', bubbles: true, cancelable: true }))
      return reverse && document.activeElement === controls[0]
    })()`)
    assert.equal(printFocusWrap, true)

    await evaluate(`(() => {
      const scaling = document.querySelector('[aria-label="Print scaling"]')
      scaling.value = 'custom'
      scaling.dispatchEvent(new Event('change', { bubbles: true }))
    })()`)
    await waitFor(`!document.querySelector('[aria-label="Custom scale percentage"]')?.disabled`, 'the custom print scale controls')
    await evaluate(`(() => {
      const scale = document.querySelector('[aria-label="Custom scale percentage"]')
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
      setter.call(scale, '125')
      scale.dispatchEvent(new Event('input', { bubbles: true }))
      scale.dispatchEvent(new Event('change', { bubbles: true }))
    })()`)
    await waitFor(`document.querySelector('.video-print-preview-frame')?.srcdoc.includes('data-scale-mode="custom"') && document.querySelector('.video-print-preview-toolbar')?.textContent.includes('125%') && document.querySelector('.video-print-preview')?.getAttribute('aria-busy') === 'false'`, 'the precise custom-scale preview')

    await evaluate(`(() => {
      const orientation = document.querySelector('[aria-label="Print orientation"]')
      orientation.value = 'portrait'
      orientation.dispatchEvent(new Event('change', { bubbles: true }))
      const color = document.querySelector('[aria-label="Print color mode"]')
      color.value = 'grayscale'
      color.dispatchEvent(new Event('change', { bubbles: true }))
      const scaling = document.querySelector('[aria-label="Print scaling"]')
      scaling.value = 'fill'
      scaling.dispatchEvent(new Event('change', { bubbles: true }))
    })()`)
    await waitFor(`document.querySelector('.video-print-preview-frame')?.srcdoc.includes('data-orientation="portrait"') && document.querySelector('.video-print-preview-frame')?.srcdoc.includes('data-color-mode="grayscale"') && document.querySelector('.video-print-preview-frame')?.contentDocument?.querySelector('img')?.complete && document.querySelector('.video-print-preview')?.getAttribute('aria-busy') === 'false'`, 'the updated portrait black-and-white preview')
    assert.equal(await evaluate(`document.querySelector('.video-print-preview-frame')?.srcdoc.includes('filter:grayscale(1)')`), true)

    const previewScreenshot = await cdp.send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: false })
    await fs.writeFile(path.join(qaArtifactDirectory, 'video-print-two-pane.png'), Buffer.from(previewScreenshot.data, 'base64'))
    await evaluate(`document.querySelector('.video-print-action').click()`)
    await waitFor(`document.querySelector('.video-print-action')?.disabled && document.querySelector('.video-print-action')?.textContent.includes('Printing')`, 'the active print job state')
    const escapeDuringPrint = await evaluate(`(() => {
      const event = new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', bubbles: true, cancelable: true })
      window.dispatchEvent(event)
      return event.defaultPrevented
    })()`)
    assert.equal(escapeDuringPrint, true)
    await delay(75)
    assert.equal(await evaluate(`Boolean(document.querySelector('.video-print-dialog'))`), true, 'Escape must not tear down an active print session')
    await waitFor(`!document.querySelector('.video-print-dialog')`, 'the QA print-to-PDF operation to complete')
    const printPdf = await fs.readFile(path.join(printDirectory, 'video-frame-print.pdf'))
    assert.equal(printPdf.subarray(0, 5).toString('ascii'), '%PDF-')
    assert.equal((printPdf.toString('latin1').match(/\/Type\s*\/Page(?!s)\b/g) || []).length, 1, 'the final print model should contain one physical page')

    await openExport()
    const modalFocus = await evaluate(`(() => ({
      inside: Boolean(document.activeElement?.closest('.export-modal')),
      label: document.activeElement?.textContent?.replace(/\\s+/g, ' ').trim() || ''
    }))()`)
    assert.equal(modalFocus.inside, true)
    assert.match(modalFocus.label, /PNG image/)

    const beforeShortcutIsolation = await evaluate(`({
      time: document.querySelector('video').currentTime,
      dialog: Boolean(document.querySelector('.export-modal'))
    })`)
    await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', code: 'ArrowRight', bubbles: true, cancelable: true }))`)
    await delay(100)
    const afterShortcutIsolation = await evaluate(`({
      time: document.querySelector('video').currentTime,
      dialog: Boolean(document.querySelector('.export-modal'))
    })`)
    assert.equal(beforeShortcutIsolation.dialog, true)
    assert.equal(afterShortcutIsolation.dialog, true)
    assert.equal(afterShortcutIsolation.time, beforeShortcutIsolation.time)

    const openShortcutBlocked = await evaluate(`(() => {
      const event = new KeyboardEvent('keydown', { key: 'o', code: 'KeyO', ctrlKey: true, bubbles: true, cancelable: true })
      window.dispatchEvent(event)
      return event.defaultPrevented
    })()`)
    assert.equal(openShortcutBlocked, true)

    const focusWrap = await evaluate(`(() => {
      const dialog = document.querySelector('.export-modal')
      const buttons = [...dialog.querySelectorAll('button:not(:disabled)')]
      buttons.at(-1).focus()
      buttons.at(-1).dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', code: 'Tab', bubbles: true, cancelable: true }))
      return document.activeElement === buttons[0]
    })()`)
    assert.equal(focusWrap, true)

    const options = await evaluate(`[...document.querySelectorAll('.export-options button')].map((item) => ({ text: item.textContent.replace(/\\s+/g, ' ').trim(), disabled: item.disabled }))`)
    assert.equal(options.length, 3)
    assert.equal(options.every((option) => !option.disabled), true)
    assert.match(options[0].text, /PNG image/)
    assert.match(options[1].text, /JPEG image/)
    assert.match(options[2].text, /Original video copy/)

    await clickExport('PNG image')
    await openExport()
    await clickExport('JPEG image')
    await openExport()
    await clickExport('Original video copy')

    const pngPath = path.join(exportDirectory, 'export fixture frame 00-00-00.png')
    const jpegPath = path.join(exportDirectory, 'export fixture frame 00-00-00.jpg')
    const copyPath = path.join(exportDirectory, 'export fixture copy.webm')
    const [png, jpeg] = await Promise.all([fs.readFile(pngPath), fs.readFile(jpegPath)])
    assert.equal(png.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])), true)
    assert.equal(jpeg[0], 0xff)
    assert.equal(jpeg[1], 0xd8)
    assert.equal(jpeg[jpeg.length - 2], 0xff)
    assert.equal(jpeg[jpeg.length - 1], 0xd9)
    assert.equal(await sha256(copyPath), await sha256(fixturePath))

    // The second half is red; a stale capture of the previously presented frame
    // is blue. Seek and export in the same event turn to exercise pending decode.
    await openExport()
    const pendingSeek = await evaluate(`(() => {
      const video=document.querySelector('video'); video.currentTime=1.5;
      const seeking=video.seeking;
      [...document.querySelectorAll('.export-options button')].find(item=>item.textContent.includes('PNG image')).click();
      return seeking;
    })()`)
    assert.equal(pendingSeek,true,'fixture must exercise an in-flight seek')
    await waitFor(`!document.querySelector('.export-modal')`, 'pending-seek export')
    const seekPng = await fs.readFile(path.join(exportDirectory,'export fixture frame 00-00-01.500.png'))
    const pixel = await evaluate(`(async () => {
      const image=new Image(); image.src='data:image/png;base64,${seekPng.toString('base64')}'; await image.decode();
      const canvas=document.createElement('canvas'); canvas.width=320; canvas.height=180;
      const context=canvas.getContext('2d'); context.drawImage(image,0,0);
      return Array.from(context.getImageData(160,90,1,1).data);
    })()`)
    assert.ok(pixel[0]>180&&pixel[1]<90&&pixel[2]<90,`expected the red 1.5s frame, got ${pixel}`)
    assert.equal(await evaluate(`document.querySelector('video').currentTime`),1.5,'capturing must not advance the playhead')
    await evaluate(`(() => { const video=document.querySelector('video'); video.currentTime=0.2; window.dispatchEvent(new KeyboardEvent('keydown',{key:'p',code:'KeyP',ctrlKey:true,bubbles:true,cancelable:true})); })()`)
    await waitFor(`Boolean(document.querySelector('.video-print-dialog'))`, 'pending-seek print capture')
    await waitFor(`document.querySelector('.video-print-preview-frame')?.contentDocument?.querySelector('img')?.complete`, 'pending-seek print image')
    const printPixel = await evaluate(`(() => {
      const image=document.querySelector('.video-print-preview-frame').contentDocument.querySelector('img');
      const canvas=document.createElement('canvas');canvas.width=320;canvas.height=180;
      const context=canvas.getContext('2d');context.drawImage(image,0,0);return Array.from(context.getImageData(160,90,1,1).data);
    })()`)
    assert.ok(printPixel[2]>110&&printPixel[0]<100,`expected the blue 0.2s print frame, got ${printPixel}`)
    assert.equal(await evaluate(`document.querySelector('video').currentTime`),0.2)

    const result = {
      toolbar,
      printToolbar,
      printPreview: { paper: 'Letter', twoPane: true, capturedResolution: '320 × 180', finalPdfBytes: printPdf.length },
      options: options.map((option) => option.text),
      pngBytes: png.length,
      jpegBytes: jpeg.length,
      copyMatchesOriginal: true,
      decodedResolution: '320 × 180',
      pendingSeekCapture: { exportTime:1.5, exportedPixel:pixel, printTime:0.2, printedPixel:printPixel, playheadUnchanged:true },
    }
    process.stdout.write(`${JSON.stringify(result, null, 2)}\n`)
    await cdp.send('Browser.close').catch(() => {})
  } catch (error) {
    if (stderr) process.stderr.write(stderr)
    throw error
  } finally {
    cdp?.close()
    if (child && child.exitCode === null) {
      child.kill()
      await Promise.race([
        new Promise((resolve) => child.once('exit', resolve)),
        delay(2_000),
      ])
    }
    const safePrefix = path.join(os.tmpdir(), 'simple-video-export-')
    if (temporaryRoot.startsWith(safePrefix)) await fs.rm(temporaryRoot, { recursive: true, force: true, maxRetries: 5, retryDelay: 120 })
  }
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`)
  process.exitCode = 1
})

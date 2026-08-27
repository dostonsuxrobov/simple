'use strict'

const assert = require('node:assert/strict')
const crypto = require('node:crypto')
const fs = require('node:fs')
const path = require('node:path')
const { spawn, spawnSync } = require('node:child_process')

function argument(name) {
  const prefix = `--${name}=`
  return process.argv.find((value) => value.startsWith(prefix))?.slice(prefix.length) || null
}

const executable = path.resolve(argument('exe') || path.join(__dirname, '..', 'release', 'simple.exe'))
const imagePath = path.resolve(argument('image') || '')
const videoPath = path.resolve(argument('video') || '')
const oversizedImagePath = argument('oversized-image') ? path.resolve(argument('oversized-image')) : null
const invalidVideoPath = argument('invalid-video') ? path.resolve(argument('invalid-video')) : null

for (const [label, filePath] of [['portable executable', executable], ['image fixture', imagePath], ['video fixture', videoPath]]) {
  assert.ok(filePath && fs.statSync(filePath).isFile(), `Missing ${label}: ${filePath}`)
}
if (oversizedImagePath) assert.ok(fs.statSync(oversizedImagePath).isFile(), `Missing oversized image fixture: ${oversizedImagePath}`)
if (invalidVideoPath) assert.ok(fs.statSync(invalidVideoPath).isFile(), `Missing invalid video fixture: ${invalidVideoPath}`)

function delay(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds))
}

async function pollTarget(port, mode, timeout = 60_000) {
  const started = Date.now()
  while (Date.now() - started < timeout) {
    try {
      const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
      const target = targets.find((item) => item.type === 'page' && item.url.includes(`/modules/${mode}/dist/index.html`))
      if (target?.webSocketDebuggerUrl) return target
    } catch {}
    await delay(150)
  }
  throw new Error(`Timed out waiting for the packaged ${mode} workspace.`)
}

class CdpClient {
  constructor(url) {
    this.url = url
    this.nextId = 0
    this.pending = new Map()
    this.socket = null
  }

  async connect() {
    this.socket = new WebSocket(this.url)
    await new Promise((resolve, reject) => {
      const onOpen = () => { cleanup(); resolve() }
      const onError = () => { cleanup(); reject(new Error('Could not connect to the packaged workspace.')) }
      const cleanup = () => {
        this.socket.removeEventListener('open', onOpen)
        this.socket.removeEventListener('error', onError)
      }
      this.socket.addEventListener('open', onOpen)
      this.socket.addEventListener('error', onError)
    })
    this.socket.addEventListener('message', (event) => {
      const message = JSON.parse(String(event.data))
      if (!message.id || !this.pending.has(message.id)) return
      const { resolve, reject } = this.pending.get(message.id)
      this.pending.delete(message.id)
      if (message.error) reject(new Error(message.error.message))
      else resolve(message.result)
    })
    this.socket.addEventListener('close', () => {
      for (const { reject } of this.pending.values()) reject(new Error('The packaged workspace closed.'))
      this.pending.clear()
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
    const response = await this.call('Runtime.evaluate', {
      expression,
      awaitPromise: true,
      returnByValue: true,
      userGesture: true,
    })
    if (response.exceptionDetails) {
      throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text || 'Renderer evaluation failed.')
    }
    return response.result.value
  }

  async waitFor(expression, timeout = 15_000) {
    const started = Date.now()
    while (Date.now() - started < timeout) {
      try {
        if (await this.evaluate(expression)) return
      } catch {}
      await delay(100)
    }
    throw new Error(`Timed out waiting for renderer state: ${expression}`)
  }

  close() {
    try { this.socket?.close() } catch {}
  }
}

async function waitForExit(child, timeout = 20_000) {
  if (child.exitCode !== null) return
  await Promise.race([
    new Promise((resolve) => child.once('exit', resolve)),
    delay(timeout).then(() => { throw new Error('The portable workspace did not exit cleanly.') }),
  ])
}

async function withPortableMode(mode, filePath, port, action) {
  const child = spawn(executable, [`--remote-debugging-port=${port}`, filePath], {
    stdio: 'ignore',
    windowsHide: true,
  })
  let client = null
  try {
    const target = await pollTarget(port, mode)
    client = new CdpClient(target.webSocketDebuggerUrl)
    await client.connect()
    const result = await action(client)
    await client.evaluate(mode === 'image' ? 'window.simpleImage.confirmClose()' : 'window.simpleVideo.close()').catch(() => {})
    client.close()
    await waitForExit(child)
    return result
  } finally {
    client?.close()
    if (child.exitCode === null) {
      spawnSync('taskkill.exe', ['/PID', String(child.pid), '/T', '/F'], { stdio: 'ignore', windowsHide: true })
    }
  }
}

async function testImage(port) {
  const beforeHash = crypto.createHash('sha256').update(fs.readFileSync(imagePath)).digest('hex')
  return withPortableMode('image', imagePath, port, async (client) => {
    await client.waitFor('document.querySelector("canvas")?.width > 0 && document.querySelector(".inspector") !== null')
    const initial = await client.evaluate(`(() => {
      const canvas = document.querySelector('canvas')
      return { width: canvas.width, height: canvas.height, details: document.querySelector('.inspector').innerText }
    })()`)
    assert.match(initial.details, /PNG/)

    await client.evaluate(`document.querySelector('button[title="Crop"]').click()`)
    await client.waitFor('document.querySelector(".crop-box") !== null')
    await client.evaluate(`Array.from(document.querySelectorAll('button')).find((button) => button.textContent.includes('Apply crop')).click()`)
    await client.waitFor(`document.querySelector('canvas').width < ${initial.width}`)
    const cropped = await client.evaluate(`({ width: document.querySelector('canvas').width, height: document.querySelector('canvas').height })`)

    await client.evaluate(`document.querySelector('button[title="Rotate right"]').click()`)
    await client.waitFor(`document.querySelector('canvas').width === ${cropped.height}`)
    const rotated = await client.evaluate(`({ width: document.querySelector('canvas').width, height: document.querySelector('canvas').height })`)
    assert.deepEqual(rotated, { width: cropped.height, height: cropped.width })

    const stroke = await client.evaluate(`(() => {
      document.querySelector('button[title="Brush"]').click()
      const color = document.querySelector('input[type="color"]')
      const setValue = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
      setValue.call(color, '#ff00ff')
      color.dispatchEvent(new Event('input', { bubbles: true }))
      color.dispatchEvent(new Event('change', { bubbles: true }))
      const size = document.querySelector('input[type="range"]')
      setValue.call(size, '40')
      size.dispatchEvent(new Event('input', { bubbles: true }))
      size.dispatchEvent(new Event('change', { bubbles: true }))
      const canvas = document.querySelector('canvas')
      const bounds = canvas.getBoundingClientRect()
      return { x1: bounds.left + bounds.width * .43, x2: bounds.left + bounds.width * .57, y: bounds.top + bounds.height * .5 }
    })()`)
    await client.waitFor(`document.querySelector('.paint-values')?.innerText.includes('40 px') && getComputedStyle(document.querySelector('.paint-preview span')).backgroundColor === 'rgb(255, 0, 255)'`)
    await client.call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: stroke.x1, y: stroke.y })
    await client.call('Input.dispatchMouseEvent', { type: 'mousePressed', x: stroke.x1, y: stroke.y, button: 'left', buttons: 1, clickCount: 1 })
    await client.call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: (stroke.x1 + stroke.x2) / 2, y: stroke.y, button: 'left', buttons: 1 })
    await client.call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: stroke.x2, y: stroke.y, button: 'left', buttons: 1 })
    await client.call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: stroke.x2, y: stroke.y, button: 'left', buttons: 0, clickCount: 1 })
    await client.waitFor(`document.querySelector('.save-state')?.textContent === 'Modified'`)
    const painted = await client.evaluate(`(() => {
      const canvas = document.querySelector('canvas')
      const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data
      let magenta = 0
      for (let index = 0; index < pixels.length; index += 4) {
        if (pixels[index] > 200 && pixels[index + 1] < 80 && pixels[index + 2] > 200 && pixels[index + 3] > 200) magenta += 1
      }
      return magenta
    })()`)
    assert.ok(painted > 100, `Brush stroke did not add enough magenta pixels: ${painted}`)

    await client.evaluate(`document.querySelector('.primary-button').click()`)
    await client.waitFor(`document.querySelector('.save-state')?.textContent === 'Saved'`, 20_000)
    const afterHash = crypto.createHash('sha256').update(fs.readFileSync(imagePath)).digest('hex')
    assert.notEqual(afterHash, beforeHash, 'Saving did not update the edited image fixture.')
    assert.deepEqual([...fs.readFileSync(imagePath).subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10])
    return { initial, cropped, rotated, painted: true, saved: true }
  })
}

async function testVideo(port) {
  return withPortableMode('video', videoPath, port, async (client) => {
    await client.waitFor(`document.querySelector('video')?.readyState >= 1 && document.querySelector('video').duration > 2`)
    const metadata = await client.evaluate(`(() => {
      const video = document.querySelector('video')
      return { duration: video.duration, width: video.videoWidth, height: video.videoHeight }
    })()`)
    assert.equal(metadata.width, 640)
    assert.equal(metadata.height, 360)

    await client.evaluate(`document.querySelector('video').play()`)
    await client.waitFor(`document.querySelector('video').currentTime > .2`)
    await client.evaluate(`document.querySelector('video').pause()`)
    await client.evaluate(`(() => {
      const select = document.querySelector('.rate-control select')
      select.value = '1.5'
      select.dispatchEvent(new Event('change', { bubbles: true }))
    })()`)
    await client.waitFor(`document.querySelector('video').playbackRate === 1.5`)
    await client.evaluate(`window.simpleVideo.toggleFullscreen()`)
    await client.waitFor(`document.querySelector('.app-shell')?.classList.contains('is-fullscreen')`)
    await client.evaluate(`document.querySelector('button[title="More"]').click()`)
    await client.evaluate(`Array.from(document.querySelectorAll('.compact-menu button')).find((button) => button.textContent.includes('Close video')).click()`)
    await client.waitFor(`document.querySelector('.welcome') !== null && !document.querySelector('.app-shell')?.classList.contains('is-fullscreen')`)
    await client.evaluate(`document.querySelector('.recent-row').click()`)
    await client.waitFor(`document.querySelector('video')?.readyState >= 1 && document.querySelector('video').duration > 2`)
    return { ...metadata, playback: true, rate: 1.5, fullscreenExit: true, recentReopen: true }
  })
}

async function testOversizedImage(port) {
  return withPortableMode('image', oversizedImagePath, port, async (client) => {
    await client.waitFor(`document.querySelector('.toast')?.textContent.includes('too large to edit safely')`)
    assert.equal(await client.evaluate(`document.querySelector('.canvas-stage').hidden && document.querySelector('.welcome') !== null`), true)
    return { rejectedSafely: true }
  })
}

async function testInvalidVideo(port) {
  return withPortableMode('video', invalidVideoPath, port, async (client) => {
    await client.waitFor(`document.querySelector('.media-error') !== null`)
    const message = await client.evaluate(`document.querySelector('.media-error').innerText`)
    assert.match(message, /codec|damaged|could not be (?:read|decoded)/i)
    return { persistentError: true, message }
  })
}

async function main() {
  const basePort = 9400 + (process.pid % 100)
  const image = await testImage(basePort)
  const video = await testVideo(basePort + 1)
  const oversizedImage = oversizedImagePath ? await testOversizedImage(basePort + 2) : null
  const invalidVideo = invalidVideoPath ? await testInvalidVideo(basePort + 3) : null
  process.stdout.write(`${JSON.stringify({ executable, image, video, oversizedImage, invalidVideo }, null, 2)}\n`)
}

main().catch((error) => {
  console.error(error.stack || error)
  process.exitCode = 1
})

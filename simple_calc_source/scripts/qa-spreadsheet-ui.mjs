import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import process from 'node:process'
import ExcelJS from 'exceljs'

const port = Number(process.env.SIMPLE_CALC_QA_PORT || 9385)
const appUrl = process.env.SIMPLE_CALC_QA_URL || 'http://127.0.0.1:5173/'
const outputDirectory = path.resolve(process.cwd(), 'tmp', 'simple-calc-qa')
const fixturePath = path.join(outputDirectory, 'ui-fixture.xlsx')
const deadline = Date.now() + 180_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

await fs.mkdir(outputDirectory, { recursive: true })
const fixture = new ExcelJS.Workbook()
const summary = fixture.addWorksheet('Summary', { views: [{ state: 'frozen', xSplit: 1, ySplit: 1, topLeftCell: 'B2' }] })
summary.getCell('A1').value = 'Item'
summary.getCell('B1').value = 'Amount'
summary.getCell('A2').value = 'North'
summary.getCell('B2').value = 1200
summary.getCell('A3').value = 'South'
summary.getCell('B3').value = 800
summary.getCell('A4').value = 'Total'
summary.getCell('B4').value = { formula: 'SUM(B2:B3)', result: 2000 }
summary.getCell('B4').numFmt = '$#,##0.00'
summary.getCell('A1').font = { bold: true, color: { argb: 'FFFFFFFF' } }
summary.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF476B57' } }
summary.getCell('A1').border = { bottom: { style: 'double', color: { argb: 'FF20362B' } } }
summary.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }
summary.getCell('B1').font = { bold: true, color: { argb: 'FFFFFFFF' } }
summary.getCell('B1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF476B57' } }
summary.getCell('B2').font = { color: { theme: 4, tint: -0.2 } }
summary.getCell('B2').border = { right: { style: 'dashed', color: { indexed: 10 } } }
summary.getCell('D2').value = 'Gradient'
summary.getCell('D2').fill = {
  type: 'gradient',
  gradient: 'angle',
  degree: 45,
  stops: [
    { position: 0, color: { argb: 'FFFFFFFF' } },
    { position: 1, color: { argb: 'FF4F81BD' } },
  ],
}
summary.getCell('D4').value = 'Rotate'
summary.getCell('D4').alignment = { textRotation: 25, vertical: 'middle' }
summary.getCell('E2').value = 123456789012345
summary.getCell('E2').numFmt = '0'
summary.getCell('A6').value = 'Line one\nLine two\nLine three'
summary.getCell('A6').alignment = { wrapText: true, vertical: 'top' }
summary.getCell('C1').value = 'hidden column'
summary.columns = [{ width: 22 }, { width: 16 }, { width: 12, hidden: true }, { width: 18 }, { width: 8 }]
summary.getRow(3).hidden = true
// Exercise preserved imported dimensions far down the sheet. The scroll path
// must not linearly rescan thousands of preceding row metrics per cell.
for (let row = 5; row <= 5000; row += 1) summary.getRow(row).height = 14 + (row % 5)
await fixture.xlsx.writeFile(fixturePath)

let targets
let target
while (Date.now() < deadline) {
  try {
    targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && (item.title === 'simple_calc' || item.url === 'about:blank' || item.url.startsWith(appUrl)))
    if (target?.webSocketDebuggerUrl) break
  } catch {}
  await pause(250)
}
if (!target?.webSocketDebuggerUrl) throw new Error('No simple_calc QA page target is available.')

const socket = new WebSocket(target.webSocketDebuggerUrl)
await new Promise((resolve, reject) => {
  socket.addEventListener('open', resolve, { once: true })
  socket.addEventListener('error', reject, { once: true })
})

let requestId = 0
const pending = new Map()
const runtimeErrors = []
socket.addEventListener('message', (event) => {
  const message = JSON.parse(String(event.data))
  if (message.method === 'Runtime.exceptionThrown') runtimeErrors.push(message.params?.exceptionDetails?.exception?.description || message.params?.exceptionDetails?.text || 'Runtime exception')
  const request = pending.get(message.id)
  if (!request) return
  pending.delete(message.id)
  if (message.error) request.reject(new Error(message.error.message))
  else request.resolve(message.result)
})

function call(method, params = {}) {
  const id = ++requestId
  socket.send(JSON.stringify({ id, method, params }))
  return new Promise((resolve, reject) => pending.set(id, { resolve, reject }))
}

async function evaluate(expression) {
  const result = await call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text)
  return result.result?.value
}

async function waitFor(expression, label) {
  while (Date.now() < deadline) {
    if (await evaluate(`Boolean(${expression})`)) return
    await pause(100)
  }
  throw new Error(`Timed out waiting for ${label}`)
}

async function click(selector) {
  const clicked = await evaluate(`(() => { const element = document.querySelector(${JSON.stringify(selector)}); if (!element) return false; element.click(); return true })()`)
  assert.equal(clicked, true, `Missing clickable ${selector}`)
}

async function selectCell(address) {
  const expression = JSON.stringify(address)
  const selected = await evaluate(`(() => {
    const cell = [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-' + ${expression}));
    if (!cell) return false;
    cell.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
    return true;
  })()`)
  assert.equal(selected, true, `Could not select ${address}`)
  await pause(60)
}

async function usePromptResponses(responses) {
  await evaluate(`(() => {
    window.__simpleCalcQAPrompt = window.prompt;
    const responses = ${JSON.stringify(responses)};
    window.prompt = () => responses.shift() ?? null;
  })()`)
}

async function restorePrompt() {
  await evaluate(`(() => {
    if (window.__simpleCalcQAPrompt) window.prompt = window.__simpleCalcQAPrompt;
    delete window.__simpleCalcQAPrompt;
  })()`)
}

async function editCell(address, value) {
  const expression = JSON.stringify(address)
  const selected = await evaluate(`(() => {
    const cell = [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-' + ${expression}));
    if (!cell) return false;
    cell.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
    return true;
  })()`)
  assert.equal(selected, true, `Could not select ${address}`)
  await pause(40)
  const started = await evaluate(`(() => {
    const cell = [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-' + ${expression}));
    if (!cell) return false;
    cell.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }));
    return true;
  })()`)
  assert.equal(started, true, `Could not start editing ${address}`)
  await waitFor("document.querySelector('.cell-editor')", `${address} editor`)
  await evaluate(`(() => { const editor = document.querySelector('.cell-editor'); editor.focus(); editor.select(); })()`)
  await call('Input.insertText', { text: value })
  await pause(80)
  await evaluate("document.querySelector('.cell-editor').blur()")
  await waitFor("!document.querySelector('.cell-editor')", `${address} commit`)
}

async function dragBetween(fromSelector, toSelector) {
  const points = await evaluate(`(() => {
    const from = document.querySelector(${JSON.stringify(fromSelector)});
    const to = document.querySelector(${JSON.stringify(toSelector)});
    if (!from || !to) return null;
    const a = from.getBoundingClientRect();
    const b = to.getBoundingClientRect();
    return {
      from: { x: a.left + a.width / 2, y: a.top + a.height / 2 },
      to: { x: b.left + b.width / 2, y: b.top + b.height / 2 },
    };
  })()`)
  assert(points, `Could not drag from ${fromSelector} to ${toSelector}`)
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: points.from.x, y: points.from.y, button: 'left', buttons: 1, clickCount: 1 })
  await pause(80)
  await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: points.to.x, y: points.to.y, button: 'left', buttons: 1 })
  await pause(80)
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: points.to.x, y: points.to.y, button: 'left', buttons: 0, clickCount: 1 })
  await pause(80)
}

async function dragBy(selector, deltaX, deltaY) {
  const point = await evaluate(`(() => {
    const element = document.querySelector(${JSON.stringify(selector)});
    if (!element) return null;
    const box = element.getBoundingClientRect();
    return { x: box.left + box.width / 2, y: box.top + box.height / 2 };
  })()`)
  assert(point, `Could not find drag handle ${selector}`)
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x, y: point.y, button: 'left', buttons: 1, clickCount: 1 })
  await pause(80)
  await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: point.x + deltaX, y: point.y + deltaY, button: 'left', buttons: 1 })
  await pause(80)
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x + deltaX, y: point.y + deltaY, button: 'left', buttons: 0, clickCount: 1 })
  await pause(100)
}

async function dragFreezeHandle(axis, boundarySelector) {
  const points = await evaluate(`(() => {
    const handle = document.querySelector('[data-freeze-handle="${axis}"]');
    const boundary = document.querySelector(${JSON.stringify(boundarySelector)})?.parentElement;
    if (!handle || !boundary) return null;
    const source = handle.getBoundingClientRect();
    const target = boundary.getBoundingClientRect();
    return {
      from: { x: source.left + source.width / 2, y: source.top + source.height / 2 },
      to: ${axis === 'rows'
        ? '{ x: source.left + source.width / 2, y: target.bottom - 1 }'
        : '{ x: target.right - 1, y: source.top + source.height / 2 }'},
    };
  })()`)
  assert(points, `Could not drag the ${axis} freeze handle`)
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: points.from.x, y: points.from.y, button: 'left', buttons: 1, clickCount: 1 })
  for (let step = 1; step <= 5; step += 1) {
    const progress = step / 5
    await call('Input.dispatchMouseEvent', {
      type: 'mouseMoved',
      x: points.from.x + (points.to.x - points.from.x) * progress,
      y: points.from.y + (points.to.y - points.from.y) * progress,
      button: 'left',
      buttons: 1,
    })
    await pause(20)
  }
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: points.to.x, y: points.to.y, button: 'left', buttons: 0, clickCount: 1 })
  await pause(140)
}

async function dragFreezeDividerToOrigin(axis) {
  const points = await evaluate(`(() => {
    const divider = document.querySelector('[data-freeze-divider="${axis}"]');
    const corner = document.querySelector('.grid-corner');
    if (!divider || !corner) return null;
    const source = divider.getBoundingClientRect();
    const target = corner.getBoundingClientRect();
    return {
      from: { x: source.left + source.width / 2, y: source.top + source.height / 2 },
      to: ${axis === 'rows'
        ? '{ x: source.left + source.width / 2, y: target.bottom - 1 }'
        : '{ x: target.right - 1, y: source.top + source.height / 2 }'},
    };
  })()`)
  assert(points, `Could not drag the ${axis} freeze divider to the origin`)
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: points.from.x, y: points.from.y, button: 'left', buttons: 1, clickCount: 1 })
  for (let step = 1; step <= 5; step += 1) {
    const progress = step / 5
    await call('Input.dispatchMouseEvent', {
      type: 'mouseMoved',
      x: points.from.x + (points.to.x - points.from.x) * progress,
      y: points.from.y + (points.to.y - points.from.y) * progress,
      button: 'left',
      buttons: 1,
    })
    await pause(20)
  }
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: points.to.x, y: points.to.y, button: 'left', buttons: 0, clickCount: 1 })
  await pause(140)
}

async function measureScrollJitter() {
  const center = await evaluate(`(() => {
    const viewport = document.querySelector('.sheet-viewport');
    viewport.scrollTo({ left: 0, top: 0 });
    const box = viewport.getBoundingClientRect();
    window.__freezeJitter = { running: true, samples: [] };
    const sample = () => {
      const viewportBox = viewport.getBoundingClientRect();
      const column = document.querySelector('.column-header')?.getBoundingClientRect();
      const row = document.querySelector('.row-header')?.getBoundingClientRect();
      const corner = document.querySelector('.grid-corner')?.getBoundingClientRect();
      const frozen = [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-A1'))?.getBoundingClientRect();
      const moving = [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-D20'))?.getBoundingClientRect();
      window.__freezeJitter.samples.push({
        scrollLeft: viewport.scrollLeft,
        scrollTop: viewport.scrollTop,
        columnTop: column ? column.top - viewportBox.top : null,
        rowLeft: row ? row.left - viewportBox.left : null,
        cornerTop: corner ? corner.top - viewportBox.top : null,
        cornerLeft: corner ? corner.left - viewportBox.left : null,
        frozenTop: frozen ? frozen.top - viewportBox.top : null,
        frozenLeft: frozen ? frozen.left - viewportBox.left : null,
        movingTop: moving ? moving.top - viewportBox.top : null,
        movingLeft: moving ? moving.left - viewportBox.left : null,
      });
      if (window.__freezeJitter.running) requestAnimationFrame(sample);
    };
    requestAnimationFrame(sample);
    return { x: box.left + box.width * 0.72, y: box.top + box.height * 0.72 };
  })()`)
  assert(center, 'Could not find the grid viewport for jitter QA')
  await pause(60)
  for (let step = 0; step < 24; step += 1) {
    await call('Input.dispatchMouseEvent', { type: 'mouseWheel', x: center.x, y: center.y, deltaX: 8, deltaY: 14 })
    await pause(12)
  }
  await pause(100)
  const result = await evaluate(`(() => {
    window.__freezeJitter.running = false;
    return { samples: window.__freezeJitter.samples, scrollLeft: document.querySelector('.sheet-viewport').scrollLeft, scrollTop: document.querySelector('.sheet-viewport').scrollTop };
  })()`)
  const range = (key) => {
    const values = result.samples.map((sample) => sample[key]).filter(Number.isFinite)
    return values.length ? Math.max(...values) - Math.min(...values) : Number.POSITIVE_INFINITY
  }
  assert(result.scrollTop > 150, `Vertical jitter QA did not scroll far enough: ${result.scrollTop}px`)
  assert(result.scrollLeft > 100, `Horizontal jitter QA did not scroll far enough: ${result.scrollLeft}px`)
  for (const key of ['columnTop', 'rowLeft', 'cornerTop', 'cornerLeft', 'frozenTop', 'frozenLeft']) {
    assert(range(key) <= 2, `${key} shook by ${range(key)}px during scrolling`)
  }
  assert(range('movingTop') > 100, 'Non-frozen rows did not move during jitter QA')
  assert(range('movingLeft') > 50, 'Non-frozen columns did not move during jitter QA')
  return result
}

const percentile = (values, fraction) => {
  const sorted = values.filter(Number.isFinite).sort((left, right) => left - right)
  return sorted.length ? sorted[Math.min(sorted.length - 1, Math.ceil(sorted.length * fraction) - 1)] : Number.POSITIVE_INFINITY
}

async function performanceMetrics() {
  const { metrics = [] } = await call('Performance.getMetrics')
  return Object.fromEntries(metrics.map(({ name, value }) => [name, value]))
}

async function measureScrollResponsiveness() {
  await call('Emulation.setCPUThrottlingRate', { rate: 2 })
  try {
    await evaluate(`(() => {
      const viewport = document.querySelector('.sheet-viewport');
      viewport.scrollTo({ left: 180, top: Math.max(1000, viewport.scrollHeight - viewport.clientHeight - 3000) });
    })()`)
    await pause(250)
    const baselineFrames = await evaluate(`new Promise((resolve) => {
      const intervals = [];
      let previous = performance.now();
      const sample = (now) => {
        intervals.push(now - previous);
        previous = now;
        if (intervals.length < 24) requestAnimationFrame(sample);
        else resolve(intervals.slice(1));
      };
      requestAnimationFrame(sample);
    })`)
    const setup = await evaluate(`(() => {
      const viewport = document.querySelector('.sheet-viewport');
      const box = viewport.getBoundingClientRect();
      const state = window.__scrollResponse = {
        running: true,
        frames: [],
        wheels: [],
        scrollEvents: 0,
        mutations: 0,
        longTasks: [],
      };
      viewport.addEventListener('wheel', (event) => state.wheels.push({
        t: performance.now(),
        deltaY: event.deltaY,
        beforeTop: viewport.scrollTop,
      }), { passive: true, signal: (() => { const controller = new AbortController(); state.controller = controller; return controller.signal })() });
      viewport.addEventListener('scroll', () => { state.scrollEvents += 1; }, { passive: true, signal: state.controller.signal });
      const observer = new MutationObserver((records) => {
        for (const record of records) state.mutations += record.addedNodes.length + record.removedNodes.length;
      });
      observer.observe(document.querySelector('.sheet-canvas'), { childList: true, subtree: true });
      state.mutationObserver = observer;
      if (PerformanceObserver.supportedEntryTypes.includes('longtask')) {
        const longTaskObserver = new PerformanceObserver((list) => {
          state.longTasks.push(...list.getEntries().map((entry) => entry.duration));
        });
        longTaskObserver.observe({ type: 'longtask', buffered: false });
        state.longTaskObserver = longTaskObserver;
      }
      const frame = (now) => {
        state.frames.push({ t: now, top: viewport.scrollTop });
        if (state.running) requestAnimationFrame(frame);
      };
      requestAnimationFrame(frame);
      return {
        x: box.left + box.width * 0.72,
        y: box.top + box.height * 0.72,
        startTop: viewport.scrollTop,
      };
    })()`)
    const before = await performanceMetrics()
    const inputs = []
    const dispatchLatencies = []
    const wheelBurst = async (deltaY) => {
      for (let index = 0; index < 24; index += 1) {
        inputs.push({ deltaY })
        const dispatchStarted = performance.now()
        await call('Input.dispatchMouseEvent', {
          type: 'mouseWheel',
          x: setup.x,
          y: setup.y,
          deltaX: 0,
          deltaY,
        })
        dispatchLatencies.push(performance.now() - dispatchStarted)
        await pause(12)
      }
    }
    await wheelBurst(44)
    await pause(120)
    const downTop = await evaluate("document.querySelector('.sheet-viewport').scrollTop")
    await wheelBurst(-44)
    await pause(220)
    const result = await evaluate(`(() => {
      const state = window.__scrollResponse;
      state.running = false;
      state.controller.abort();
      state.mutationObserver.disconnect();
      state.longTaskObserver?.disconnect();
      return {
        frames: state.frames,
        wheels: state.wheels,
        scrollEvents: state.scrollEvents,
        mutations: state.mutations,
        longTasks: state.longTasks,
        finalTop: document.querySelector('.sheet-viewport').scrollTop,
      };
    })()`)
    const after = await performanceMetrics()
    const visualLatencies = []
    for (const wheel of result.wheels) {
      const direction = Math.sign(wheel.deltaY)
      const movingFrame = result.frames.find((frame) => frame.t >= wheel.t && direction * (frame.top - wheel.beforeTop) > 0)
      if (movingFrame) visualLatencies.push(movingFrame.t - wheel.t)
    }
    const frameGaps = result.frames.slice(1).map((frame, index) => frame.t - result.frames[index].t)
    const metricDelta = (name) => Math.max(0, (after[name] || 0) - (before[name] || 0)) * 1000
    const summary = {
      inputs: inputs.length,
      wheels: result.wheels.length,
      downTravel: downTop - setup.startTop,
      upTravel: downTop - result.finalTop,
      returnError: Math.abs(result.finalTop - setup.startTop),
      dispatchP95: percentile(dispatchLatencies, 0.95),
      dispatchMax: Math.max(0, ...dispatchLatencies),
      visualP95: percentile(visualLatencies, 0.95),
      visualMax: Math.max(0, ...visualLatencies),
      frameP95: percentile(frameGaps, 0.95),
      frameMax: Math.max(0, ...frameGaps),
      baselineP95: percentile(baselineFrames, 0.95),
      longTaskCount: result.longTasks.length,
      longTaskTotal: result.longTasks.reduce((sum, value) => sum + value, 0),
      taskPerInput: metricDelta('TaskDuration') / inputs.length,
      scriptPerInput: metricDelta('ScriptDuration') / inputs.length,
      styleLayoutPerInput: (metricDelta('LayoutDuration') + metricDelta('RecalcStyleDuration')) / inputs.length,
      mutations: result.mutations,
      scrollEvents: result.scrollEvents,
    }
    assert(summary.wheels >= 46, `Only ${summary.wheels} of 48 wheel inputs reached the grid`)
    assert(summary.downTravel >= 950 && summary.upTravel >= 950, `Scroll travel was incomplete: ${summary.downTravel}px down, ${summary.upTravel}px up`)
    assert(summary.returnError <= 8, `Direction reversal returned ${summary.returnError}px from its start`)
    assert(summary.dispatchP95 <= 55 && summary.dispatchMax <= 100, `Wheel dispatch latency was ${summary.dispatchP95}ms p95 / ${summary.dispatchMax}ms max`)
    assert(summary.visualP95 <= 65 && summary.visualMax <= 130, `Visual scroll latency was ${summary.visualP95}ms p95 / ${summary.visualMax}ms max`)
    assert(summary.frameP95 <= Math.max(38, summary.baselineP95 + 18), `Scroll frame p95 was ${summary.frameP95}ms versus ${summary.baselineP95}ms baseline`)
    assert(summary.frameMax <= 90, `A scroll frame stalled for ${summary.frameMax}ms`)
    assert(summary.longTaskCount <= 1 && summary.longTaskTotal <= 90, `Scroll produced ${summary.longTaskCount} long tasks totaling ${summary.longTaskTotal}ms`)
    assert(summary.taskPerInput <= 22, `Main-thread scroll work averaged ${summary.taskPerInput}ms per input`)
    assert(summary.scriptPerInput <= 11, `Scroll script work averaged ${summary.scriptPerInput}ms per input`)
    assert(summary.styleLayoutPerInput <= 4, `Scroll style/layout work averaged ${summary.styleLayoutPerInput}ms per input`)
    return summary
  } finally {
    await call('Emulation.setCPUThrottlingRate', { rate: 1 })
  }
}

async function doubleClick(selector) {
  const point = await evaluate(`(() => {
    const element = document.querySelector(${JSON.stringify(selector)});
    if (!element) return null;
    const box = element.getBoundingClientRect();
    return { x: box.left + box.width / 2, y: box.top + box.height / 2 };
  })()`)
  assert(point, `Could not find double-click target ${selector}`)
  for (const clickCount of [1, 2]) {
    await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x, y: point.y, button: 'left', buttons: 1, clickCount })
    await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x, y: point.y, button: 'left', buttons: 0, clickCount })
    await pause(60)
  }
  await pause(100)
}

async function screenshot(fileName) {
  const result = await call('Page.captureScreenshot', { format: 'png', captureBeyondViewport: false })
  await fs.writeFile(path.join(outputDirectory, fileName), Buffer.from(result.data, 'base64'))
}

await call('Runtime.enable')
await call('Page.enable')
await call('Performance.enable')
await evaluate("localStorage.removeItem('simple-calc:recent:v1'); window.__simpleCalcQAReloading = true")
await call('Page.reload', { ignoreCache: true })
await waitFor("document.readyState === 'complete' && !window.__simpleCalcQAReloading", 'initial page reload')
await waitFor("document.querySelector('.welcome')", 'welcome screen')
const welcomeMetrics = await evaluate(`(() => ({
  title: document.title,
  heading: document.querySelector('.welcome h1')?.textContent,
  buttons: document.querySelectorAll('.hero-actions button').length,
  width: innerWidth,
  height: innerHeight,
}))()`)
assert.equal(welcomeMetrics.title, 'simple_calc')
assert.equal(welcomeMetrics.heading, 'Numbers, without the noise.')
assert.equal(welcomeMetrics.buttons, 2)
await screenshot('welcome.png')

await click('.primary-action')
await waitFor("document.querySelector('.sheet-viewport')", 'new workbook grid')
const gridMetrics = await evaluate(`(() => ({
  cells: document.querySelectorAll('.grid-cell').length,
  tabs: document.querySelectorAll('.sheet-tab').length,
  formulaBar: Boolean(document.querySelector('.formula-bar input')),
  menuTriggers: [...document.querySelectorAll('[data-menu-trigger]')].map((item) => item.textContent.trim()),
  toolbarControls: document.querySelectorAll('.command-bar button, .command-bar select').length,
  active: document.querySelector('.name-box')?.textContent,
}))()`)
assert(gridMetrics.cells > 20 && gridMetrics.cells < 1500, `Unexpected virtualized cell count: ${gridMetrics.cells}`)
assert.equal(gridMetrics.tabs, 1)
assert.equal(gridMetrics.formulaBar, true)
assert.equal(gridMetrics.active, 'A1')
assert.deepEqual(gridMetrics.menuTriggers, ['View', 'Insert', 'Format', 'Data', 'Tools'])
assert(gridMetrics.toolbarControls >= 25, `Expected the expanded toolbar, found ${gridMetrics.toolbarControls} controls`)

await editCell('A1', '3')
await editCell('B1', '5')
await editCell('A2', '=A1+B1')
const formulaState = await evaluate(`(() => {
  const cell = [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-A2'));
  cell.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
  return { display: cell.textContent.trim() };
})()`)
await pause(60)
const formulaBarValue = await evaluate("document.querySelector('.formula-bar input').value")
if (formulaState.display !== '8') {
  const debugState = await evaluate(`(() => ({
    cells: [...document.querySelectorAll('.grid-cell')].filter((item) => /-(A1|B1|A2)$/.test(item.id)).map((item) => ({ id: item.id, text: item.textContent, title: item.title })),
    formulaBar: document.querySelector('.formula-bar input')?.value,
    active: document.querySelector('.name-box')?.textContent,
  }))()`)
  console.error('Formula debug state:', debugState)
}
assert.equal(formulaState.display, '8')
assert.equal(formulaBarValue, '=A1+B1')

// Structural menu commands shift both values and formula references and remain undoable.
await selectCell('A1')
await click('[data-menu-trigger="insert"]')
await click('[data-menu-action="insert-rows-above"]')
await waitFor("[...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-A3'))?.textContent.trim() === '8'", 'inserted row formula result')
const insertedRowState = await evaluate(`(() => ({
  a2: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-A2'))?.textContent.trim(),
  b2: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-B2'))?.textContent.trim(),
  a3: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-A3'))?.textContent.trim(),
  formula: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-A3'))?.title,
}))()`)
assert.deepEqual(insertedRowState, { a2: '3', b2: '5', a3: '8', formula: '=A2+B2' })
await evaluate("window.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true, cancelable: true }))")
await waitFor("[...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-A2'))?.textContent.trim() === '8'", 'undo inserted row')
assert.equal(await evaluate("[...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-A2'))?.title"), '=A1+B1')

await editCell('C1', '1')
await editCell('C2', '2')
await dragBetween('[id$="-C1"]', '[id$="-C2"]')
await dragBetween('.fill-handle', '[id$="-C5"]')
try {
  await waitFor("[...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-C5'))?.textContent.trim() === '5'", 'numeric series autofill')
} catch (error) {
  const fillDebug = await evaluate(`(() => ({
    active: document.querySelector('.name-box')?.textContent,
    selection: getComputedStyle(document.querySelector('.selection-outline')).cssText,
    values: ['C1', 'C2', 'C3', 'C4', 'C5'].map((address) => {
      const item = [...document.querySelectorAll('.grid-cell')].find((cell) => cell.id.endsWith('-' + address));
      return item ? { address, text: item.textContent.trim(), box: item.getBoundingClientRect().toJSON() } : null;
    }),
  }))()`)
  console.error('Autofill debug state:', JSON.stringify(fillDebug))
  throw error
}
const autofillState = await evaluate(`(() => ({
  values: ['C1', 'C2', 'C3', 'C4', 'C5'].map((address) => [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-' + address))?.textContent.trim()),
  aggregate: document.querySelector('.aggregate-picker output')?.textContent.trim(),
  aggregateLabel: document.querySelector('.aggregate-picker select')?.selectedOptions[0]?.textContent,
  aggregateMode: document.querySelector('.aggregate-picker select')?.value,
}))()`)
assert.deepEqual(autofillState.values, ['1', '2', '3', '4', '5'])
assert.equal(autofillState.aggregateMode, 'sum')
assert.equal(autofillState.aggregateLabel, 'Sum')
assert.equal(autofillState.aggregate, '15')
await evaluate(`(() => {
  const select = document.querySelector('.aggregate-picker select');
  select.value = 'count';
  select.dispatchEvent(new Event('change', { bubbles: true }));
})()`)
await waitFor("document.querySelector('.aggregate-picker select')?.value === 'count' && document.querySelector('.aggregate-picker output')?.textContent.trim() === '5'", 'count aggregate selection')
await screenshot('autofill-summary.png')
await evaluate(`(() => {
  const format = document.querySelector('.number-select');
  format.value = '0%';
  format.dispatchEvent(new Event('change', { bubbles: true }));
  const aggregate = document.querySelector('.aggregate-picker select');
  aggregate.value = 'average';
  aggregate.dispatchEvent(new Event('change', { bubbles: true }));
})()`)
await waitFor("document.querySelector('.aggregate-picker output')?.textContent.trim() === '300%'", 'formatted percentage aggregate')
await evaluate(`(() => {
  const format = document.querySelector('.number-select');
  format.value = 'General';
  format.dispatchEvent(new Event('change', { bubbles: true }));
  const aggregate = document.querySelector('.aggregate-picker select');
  aggregate.value = 'count';
  aggregate.dispatchEvent(new Event('change', { bubbles: true }));
})()`)
await waitFor("document.querySelector('.aggregate-picker output')?.textContent.trim() === '5'", 'aggregate reset')

await evaluate("window.dispatchEvent(new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true, cancelable: true }))")
await waitFor("document.querySelector('.search-panel input') === document.activeElement", 'focused find panel')
await call('Input.insertText', { text: '8' })
await waitFor("document.querySelector('.name-box')?.textContent === 'A2'", 'formula search result')
const searchState = await evaluate(`(() => ({
  query: document.querySelector('.search-panel input')?.value,
  count: document.querySelector('.search-count')?.textContent.trim(),
}))()`)
assert.equal(searchState.query, '8')
assert.equal(searchState.count, '1 of 1')
await screenshot('find-panel.png')
await evaluate("document.querySelector('.search-panel input').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }))")
await waitFor("!document.querySelector('.search-panel')", 'find panel close')

await click('.sheet-actions .tool-button')
await waitFor("document.querySelectorAll('.sheet-tab').length === 2", 'second sheet')
await screenshot('editor.png')

// Reload with a real recent-file entry and exercise the native open IPC path.
await evaluate(`localStorage.setItem('simple-calc:recent:v1', JSON.stringify([{ path: ${JSON.stringify(fixturePath)}, name: 'ui-fixture.xlsx', format: 'xlsx', openedAt: Date.now() }]))`)
await call('Page.reload', { ignoreCache: true })
await waitFor("document.querySelector('.recent-main')", 'recent fixture')
await click('.recent-main')
await waitFor("document.querySelector('.sheet-viewport')", 'imported workbook grid')
await waitFor("[...document.querySelectorAll('.grid-cell')].some((item) => item.id.endsWith('-B4') && item.textContent.includes('2,000'))", 'imported formula display')
const imported = await evaluate(`(() => ({
  title: document.querySelector('.document-title')?.textContent,
  tab: document.querySelector('.sheet-tab.is-active')?.textContent,
  cells: document.querySelectorAll('.grid-cell').length,
  formula: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-B4'))?.title,
  frozenDividers: document.querySelectorAll('.freeze-divider').length,
  hiddenRowPresent: [...document.querySelectorAll('.grid-cell')].some((item) => item.id.endsWith('-A3')),
  hiddenColumnPresent: [...document.querySelectorAll('.grid-cell')].some((item) => item.id.endsWith('-C1')),
  headerStyle: (() => {
    const item = [...document.querySelectorAll('.grid-cell')].find((cell) => cell.id.endsWith('-A1'));
    const style = item && getComputedStyle(item);
    return style ? { backgroundColor: style.backgroundColor, color: style.color, fontWeight: style.fontWeight, borderBottomStyle: style.borderBottomStyle } : null;
  })(),
  themeStyle: (() => {
    const item = [...document.querySelectorAll('.grid-cell')].find((cell) => cell.id.endsWith('-B2'));
    const style = item && getComputedStyle(item);
    return style ? { color: style.color, borderRightStyle: style.borderRightStyle } : null;
  })(),
  gradient: (() => {
    const item = [...document.querySelectorAll('.grid-cell')].find((cell) => cell.id.endsWith('-D2'));
    return item ? getComputedStyle(item).backgroundImage : '';
  })(),
  rotation: (() => {
    const item = [...document.querySelectorAll('.grid-cell')].find((cell) => cell.id.endsWith('-D4'))?.querySelector('.cell-content');
    return item ? getComputedStyle(item).transform : '';
  })(),
}))()`)
assert.equal(imported.title, 'ui-fixture.xlsx')
assert.equal(imported.tab, 'Summary')
assert.equal(imported.formula, '=SUM(B2:B3)')
assert(imported.cells < 1500)
assert.equal(imported.frozenDividers, 2)
assert.equal(imported.hiddenRowPresent, false)
assert.equal(imported.hiddenColumnPresent, false)
assert.deepEqual(imported.headerStyle, {
  backgroundColor: 'rgb(71, 107, 87)',
  color: 'rgb(255, 255, 255)',
  fontWeight: '700',
  borderBottomStyle: 'double',
})
assert.equal(imported.themeStyle.borderRightStyle, 'dashed')
assert.notEqual(imported.themeStyle.color, 'rgb(29, 34, 32)')
assert.match(imported.gradient, /gradient/i)
assert.notEqual(imported.rotation, 'none')

for (const menu of ['view', 'insert', 'format']) {
  await click(`[data-menu-trigger="${menu}"]`)
  await waitFor(`document.querySelector('.sheet-menu-panel[aria-label="${menu[0].toUpperCase()}${menu.slice(1)} menu"]')`, `${menu} menu`)
  await screenshot(`${menu}-menu.png`)
  await click(`[data-menu-trigger="${menu}"]`)
}

const initialDimensions = await evaluate(`(() => ({
  columnB: document.querySelector('[data-resize-column="B"]')?.parentElement.getBoundingClientRect().width,
  columnE: document.querySelector('[data-resize-column="E"]')?.parentElement.getBoundingClientRect().width,
  row4: document.querySelector('[data-resize-row="4"]')?.parentElement.getBoundingClientRect().height,
  row6: document.querySelector('[data-resize-row="6"]')?.parentElement.getBoundingClientRect().height,
}))()`)
assert(initialDimensions.columnB > 0 && initialDimensions.columnE > 0)
assert(initialDimensions.row4 > 0 && initialDimensions.row6 > 0)

await dragBy('[data-resize-column="B"]', 48, 0)
await dragBy('[data-resize-row="4"]', 0, 28)
const manualDimensions = await evaluate(`(() => ({
  columnB: document.querySelector('[data-resize-column="B"]')?.parentElement.getBoundingClientRect().width,
  row4: document.querySelector('[data-resize-row="4"]')?.parentElement.getBoundingClientRect().height,
}))()`)
assert(Math.abs(manualDimensions.columnB - initialDimensions.columnB - 48) <= 3, `Column drag delta was ${manualDimensions.columnB - initialDimensions.columnB}px`)
assert(Math.abs(manualDimensions.row4 - initialDimensions.row4 - 28) <= 3, `Row drag delta was ${manualDimensions.row4 - initialDimensions.row4}px`)

await doubleClick('[data-resize-column="E"]')
await doubleClick('[data-resize-row="6"]')
const fittedDimensions = await evaluate(`(() => {
  const numberCell = [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-E2'));
  const wrappedCell = [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-A6'));
  const numberContent = numberCell?.querySelector('.cell-content');
  const wrappedContent = wrappedCell?.querySelector('.cell-content');
  return {
    columnB: document.querySelector('[data-resize-column="B"]')?.parentElement.getBoundingClientRect().width,
    columnE: document.querySelector('[data-resize-column="E"]')?.parentElement.getBoundingClientRect().width,
    row4: document.querySelector('[data-resize-row="4"]')?.parentElement.getBoundingClientRect().height,
    row6: document.querySelector('[data-resize-row="6"]')?.parentElement.getBoundingClientRect().height,
    numberText: numberCell?.textContent.trim(),
    numberFits: numberContent ? numberContent.scrollWidth <= numberContent.clientWidth + 1 : false,
    wrappedText: wrappedCell?.textContent,
    wrappedFits: wrappedContent && wrappedCell ? wrappedContent.scrollHeight <= wrappedCell.clientHeight + 1 : false,
  };
})()`)
assert(fittedDimensions.columnE > initialDimensions.columnE + 25, `Column auto-fit width was ${fittedDimensions.columnE}px`)
assert(fittedDimensions.row6 > initialDimensions.row6 + 20, `Row auto-fit height was ${fittedDimensions.row6}px`)
assert.equal(fittedDimensions.numberText, '123456789012345')
assert.equal(fittedDimensions.numberFits, true)
assert.equal(fittedDimensions.wrappedText, 'Line one\nLine two\nLine three')
assert.equal(fittedDimensions.wrappedFits, true)
await screenshot('resized-autofit-workbook.png')

// Exercise screenshot-requested Insert, Format, and View menu behavior.
await selectCell('F2')
await click('[data-menu-trigger="insert"]')
await click('[data-menu-action="insert-checkbox"]')
await waitFor("[...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-F2'))?.querySelector('.cell-checkbox')", 'inserted checkbox')
await evaluate("[...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-F2')).querySelector('.cell-checkbox').click()")
await waitFor("[...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-F2'))?.querySelector('.cell-checkbox')?.checked === true", 'checked checkbox')

await selectCell('G2')
await usePromptResponses(['Low, Medium, High'])
await click('[data-menu-trigger="insert"]')
await click('[data-menu-action="insert-dropdown"]')
await restorePrompt()
await waitFor("[...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-G2'))?.querySelector('.cell-dropdown')", 'inserted dropdown')
await evaluate(`(() => {
  const dropdown = [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-G2')).querySelector('.cell-dropdown');
  dropdown.value = 'High';
  dropdown.dispatchEvent(new Event('change', { bubbles: true }));
})()`)
await waitFor("[...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-G2'))?.querySelector('.cell-dropdown')?.value === 'High'", 'selected dropdown value')

await selectCell('H2')
await usePromptResponses(['Persistent note'])
await click('[data-menu-trigger="insert"]')
await click('[data-menu-action="insert-note"]')
await restorePrompt()
await waitFor("[...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-H2'))?.classList.contains('has-note')", 'inserted note')

await selectCell('I2')
await usePromptResponses(['https://example.com', 'Example'])
await click('[data-menu-trigger="insert"]')
await click('[data-menu-action="insert-link"]')
await restorePrompt()
await waitFor("[...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-I2'))?.querySelector('.cell-hyperlink')?.textContent === 'Example'", 'inserted hyperlink')

await editCell('A8', 'Merged title')
await dragBetween('[id$="-A8"]', '[id$="-B8"]')
await click('[data-menu-trigger="format"]')
await click('[data-menu-action="merge-all"]')
await waitFor("![...document.querySelectorAll('.grid-cell')].some((item) => item.id.endsWith('-B8'))", 'merged range')

await dragFreezeHandle('rows', '[data-resize-row="2"]')
await dragFreezeHandle('columns', '[data-resize-column="B"]')
await waitFor("document.querySelector('[data-freeze-handle=\"rows\"]')?.getAttribute('aria-valuenow') === '2' && document.querySelector('[data-freeze-handle=\"columns\"]')?.getAttribute('aria-valuenow') === '2'", 'dragged row and column freeze handles')
const freezeHandleState = await evaluate(`(() => {
  const row = document.querySelector('[data-resize-row="2"]').parentElement.getBoundingClientRect();
  const column = document.querySelector('[data-resize-column="B"]').parentElement.getBoundingClientRect();
  const horizontal = document.querySelector('.freeze-divider.is-horizontal').getBoundingClientRect();
  const vertical = document.querySelector('.freeze-divider.is-vertical').getBoundingClientRect();
  return {
    rows: Number(document.querySelector('[data-freeze-handle="rows"]').getAttribute('aria-valuenow')),
    columns: Number(document.querySelector('[data-freeze-handle="columns"]').getAttribute('aria-valuenow')),
    rowAlignment: Math.abs(horizontal.top - row.bottom),
    columnAlignment: Math.abs(vertical.left - column.right),
    horizontalThickness: horizontal.height,
    verticalThickness: vertical.width,
  };
})()`)
assert.equal(freezeHandleState.rows, 2)
assert.equal(freezeHandleState.columns, 2)
assert(freezeHandleState.rowAlignment <= 2, `Row freeze line missed its boundary by ${freezeHandleState.rowAlignment}px`)
assert(freezeHandleState.columnAlignment <= 2, `Column freeze line missed its boundary by ${freezeHandleState.columnAlignment}px`)
assert(freezeHandleState.horizontalThickness >= 3)
assert(freezeHandleState.verticalThickness >= 3)

await dragFreezeDividerToOrigin('rows')
await dragFreezeDividerToOrigin('columns')
await waitFor("document.querySelector('[data-freeze-handle=\"rows\"]')?.getAttribute('aria-valuenow') === '0' && document.querySelector('[data-freeze-handle=\"columns\"]')?.getAttribute('aria-valuenow') === '0' && !document.querySelector('.freeze-divider')", 'dragged freeze dividers back to the corner')
await dragFreezeHandle('rows', '[data-resize-row="2"]')
await dragFreezeHandle('columns', '[data-resize-column="B"]')
await waitFor("document.querySelector('[data-freeze-handle=\"rows\"]')?.getAttribute('aria-valuenow') === '2' && document.querySelector('[data-freeze-handle=\"columns\"]')?.getAttribute('aria-valuenow') === '2'", 'restored dragged row and column freezes')

await click('[data-menu-trigger="view"]')
await click('[data-menu-action="view-show-gridlines"]')
await waitFor("document.querySelector('.sheet-viewport')?.classList.contains('hides-gridlines')", 'hidden gridlines')
const insertedFeatureState = await evaluate(`(() => ({
  checkbox: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-F2'))?.querySelector('.cell-checkbox')?.checked,
  dropdown: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-G2'))?.querySelector('.cell-dropdown')?.value,
  note: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-H2'))?.title,
  link: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-I2'))?.querySelector('.cell-hyperlink')?.textContent,
  mergedChildPresent: [...document.querySelectorAll('.grid-cell')].some((item) => item.id.endsWith('-B8')),
  gridlinesHidden: document.querySelector('.sheet-viewport')?.classList.contains('hides-gridlines'),
  frozenDividers: document.querySelectorAll('.freeze-divider').length,
  frozenRows: Number(document.querySelector('[data-freeze-handle="rows"]')?.getAttribute('aria-valuenow')),
  frozenColumns: Number(document.querySelector('[data-freeze-handle="columns"]')?.getAttribute('aria-valuenow')),
}))()`)
assert.deepEqual(insertedFeatureState, {
  checkbox: true,
  dropdown: 'High',
  note: 'Persistent note',
  link: 'Example',
  mergedChildPresent: false,
  gridlinesHidden: true,
  frozenDividers: 2,
  frozenRows: 2,
  frozenColumns: 2,
})
await screenshot('menus-and-cell-features.png')
await measureScrollJitter()
await screenshot('drag-freeze-and-scroll.png')
const scrollPerformance = await measureScrollResponsiveness()
console.log(`Scroll QA: ${scrollPerformance.visualP95.toFixed(1)}ms visual p95, ${scrollPerformance.frameP95.toFixed(1)}ms frame p95, ${scrollPerformance.taskPerInput.toFixed(1)}ms main-thread/input, ${scrollPerformance.mutations} DOM mutations.`)
await evaluate("document.querySelector('.sheet-viewport').scrollTo({ left: 0, top: 0 })")
await waitFor("document.querySelector('.sheet-viewport').scrollLeft === 0 && document.querySelector('.sheet-viewport').scrollTop === 0", 'grid scroll reset')

// Exercise the actual Electron save bridge against an opened source workbook.
// This verifies that an ordinary edit uses source-backed XLSX overlay rather
// than rebuilding the workbook and shedding its imported structure.
await editCell('B2', '1300')
await evaluate("window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true, cancelable: true }))")
await waitFor("document.querySelector('.toast')?.textContent.includes('Saved ui-fixture.xlsx')", 'source-backed workbook save')
const savedFixture = new ExcelJS.Workbook()
await savedFixture.xlsx.readFile(fixturePath)
const savedSummary = savedFixture.getWorksheet('Summary')
assert(savedSummary)
assert.equal(savedSummary.getCell('B2').value, 1300)
assert.equal(savedSummary.getCell('B4').formula, 'SUM(B2:B3)')
assert.equal(savedSummary.getCell('A1').fill.fgColor.argb, 'FF476B57')
assert.equal(savedSummary.getCell('A1').border.bottom.style, 'double')
assert.equal(savedSummary.getRow(3).hidden, true)
assert.equal(savedSummary.getColumn(3).hidden, true)
assert.equal(savedSummary.views[0].state, 'frozen')
assert.equal(savedSummary.views[0].xSplit, 2)
assert.equal(savedSummary.views[0].ySplit, 2)
assert.equal(savedSummary.views[0].topLeftCell, 'C3')
assert.equal(savedSummary.views[0].showGridLines, false)
assert.equal(savedSummary.getCell('F2').value, true)
assert.equal(savedSummary.getCell('F2').dataValidation.type, 'list')
assert.deepEqual(savedSummary.getCell('F2').dataValidation.formulae, ['"TRUE,FALSE"'])
assert.equal(savedSummary.getCell('G2').value, 'High')
assert.equal(savedSummary.getCell('G2').dataValidation.type, 'list')
assert.deepEqual(savedSummary.getCell('G2').dataValidation.formulae, ['"Low,Medium,High"'])
assert(savedSummary.getCell('H2').note, 'Cell note did not survive save')
assert.equal(savedSummary.getCell('I2').hyperlink, 'https://example.com')
assert.equal(savedSummary.getCell('A8').value, 'Merged title')
assert.equal(savedSummary.getCell('B8').isMerged, true)
assert(Math.abs(savedSummary.getColumn(2).width - ((fittedDimensions.columnB - 5) / 7)) < 0.2)
assert(Math.abs(savedSummary.getColumn(5).width - ((fittedDimensions.columnE - 5) / 7)) < 0.2)
assert(Math.abs(savedSummary.getRow(4).height - fittedDimensions.row4 * 0.75) < 0.2)
assert(Math.abs(savedSummary.getRow(6).height - fittedDimensions.row6 * 0.75) < 0.2)

// Reopen the saved workbook and verify the native XLSX dimensions are restored.
await call('Page.reload', { ignoreCache: true })
await waitFor("document.querySelector('.recent-main')", 'saved workbook recent entry')
await click('.recent-main')
await waitFor("[...document.querySelectorAll('.grid-cell')].some((item) => item.id.endsWith('-E2') && item.textContent.trim() === '123456789012345')", 'reopened resized workbook')
const reopenedDimensions = await evaluate(`(() => ({
  columnB: document.querySelector('[data-resize-column="B"]')?.parentElement.getBoundingClientRect().width,
  columnE: document.querySelector('[data-resize-column="E"]')?.parentElement.getBoundingClientRect().width,
  row4: document.querySelector('[data-resize-row="4"]')?.parentElement.getBoundingClientRect().height,
  row6: document.querySelector('[data-resize-row="6"]')?.parentElement.getBoundingClientRect().height,
  numberTitle: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-E2'))?.title,
  checkbox: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-F2'))?.querySelector('.cell-checkbox')?.checked,
  dropdown: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-G2'))?.querySelector('.cell-dropdown')?.value,
  note: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-H2'))?.title,
  link: [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-I2'))?.querySelector('.cell-hyperlink')?.textContent,
  mergedChildPresent: [...document.querySelectorAll('.grid-cell')].some((item) => item.id.endsWith('-B8')),
  gridlinesHidden: document.querySelector('.sheet-viewport')?.classList.contains('hides-gridlines'),
  frozenDividers: document.querySelectorAll('.freeze-divider').length,
  frozenRows: Number(document.querySelector('[data-freeze-handle="rows"]')?.getAttribute('aria-valuenow')),
  frozenColumns: Number(document.querySelector('[data-freeze-handle="columns"]')?.getAttribute('aria-valuenow')),
}))()`)
for (const key of ['columnB', 'columnE', 'row4', 'row6']) {
  assert(Math.abs(reopenedDimensions[key] - fittedDimensions[key]) <= 2, `${key} changed after reopen`)
}
assert.equal(reopenedDimensions.numberTitle, '123456789012345')
assert.equal(reopenedDimensions.checkbox, true)
assert.equal(reopenedDimensions.dropdown, 'High')
assert.equal(reopenedDimensions.note, 'Persistent note')
assert.equal(reopenedDimensions.link, 'Example')
assert.equal(reopenedDimensions.mergedChildPresent, false)
assert.equal(reopenedDimensions.gridlinesHidden, true)
assert.equal(reopenedDimensions.frozenDividers, 2)
assert.equal(reopenedDimensions.frozenRows, 2)
assert.equal(reopenedDimensions.frozenColumns, 2)
await screenshot('reopened-resized-workbook.png')

assert.deepEqual(runtimeErrors, [], `Renderer exceptions: ${runtimeErrors.join('\n')}`)
socket.close()
process.stdout.write(`UI QA passed. Screenshots: ${outputDirectory}\n`)

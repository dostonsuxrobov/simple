import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import process from "node:process";

const require = createRequire(import.meta.url);
const electronPath = require("electron");
const projectRoot = path.resolve(new URL("..", import.meta.url).pathname.slice(1));
const fixture = path.join(projectRoot, "qa", "fixtures", "simple-docs-roundtrip-fixture.docx");
const profile = await mkdtemp(path.join(tmpdir(), "simple-docs-ui-profile-"));
const port = 19_000 + Math.floor(Math.random() * 900);
const child = spawn(electronPath, [
  `--remote-debugging-port=${port}`,
  `--user-data-dir=${profile}`,
  ".",
  fixture,
], { cwd: projectRoot, stdio: ["ignore", "pipe", "pipe"] });

let stderr = "";
child.stderr.on("data", (chunk) => { stderr += chunk.toString(); });

async function pollJson() {
  const deadline = Date.now() + 25_000;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(`http://127.0.0.1:${port}/json/list`);
      const targets = await response.json();
      const target = targets.find((item) => item.type === "page" && !item.url.startsWith("devtools:"));
      if (target?.webSocketDebuggerUrl) return target;
    } catch {}
    await new Promise((resolve) => setTimeout(resolve, 120));
  }
  throw new Error(`Electron did not expose its renderer. ${stderr}`);
}

class Cdp {
  constructor(url) {
    this.id = 0;
    this.pending = new Map();
    this.socket = new WebSocket(url);
  }
  async open() {
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(String(event.data));
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }
  send(method, params = {}) {
    const id = ++this.id;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }
  async evaluate(expression) {
    const result = await this.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.text);
    return result.result.value;
  }
  close() { this.socket.close(); }
}

async function waitFor(cdp, expression, label, timeout = 30_000) {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    if (await cdp.evaluate(expression)) return;
    await new Promise((resolve) => setTimeout(resolve, 140));
  }
  throw new Error(`Timed out waiting for ${label}.`);
}

let cdp;
try {
  const target = await pollJson();
  cdp = new Cdp(target.webSocketDebuggerUrl);
  await cdp.open();
  await cdp.send("Runtime.enable");
  await waitFor(cdp, `document.title.includes("simple-docs-roundtrip-fixture") && document.querySelector("#editor.is-active")`, "fixture document");
  await cdp.send("Input.dispatchKeyEvent", { type: "keyDown", key: "p", code: "KeyP", windowsVirtualKeyCode: 80, modifiers: 2 });
  await cdp.send("Input.dispatchKeyEvent", { type: "keyUp", key: "p", code: "KeyP", windowsVirtualKeyCode: 80, modifiers: 2 });
  await waitFor(cdp, `!document.querySelector("#print-modal").hidden`, "print dialog");
  await waitFor(cdp, `!document.querySelector("#print-submit").disabled && document.querySelector("#print-preview-pdf").src.startsWith("blob:")`, "initial vector preview", 45_000);

  const documentPaperAudit = await cdp.evaluate(`(() => ({
    orientationDisabled: [...document.querySelectorAll('input[name="print-orientation"]')].every((input) => input.disabled),
    noteVisible: !document.querySelector("#print-document-paper-note").hidden,
    note: document.querySelector("#print-document-paper-note").textContent,
  }))()`);
  if (!documentPaperAudit.orientationDisabled || !documentPaperAudit.noteVisible || !/original size and orientation/i.test(documentPaperAudit.note)) {
    throw new Error(`Document-paper behavior is not explicit.\n${JSON.stringify(documentPaperAudit, null, 2)}`);
  }

  const disabledDuringDebounce = await cdp.evaluate(`(() => {
    const set = (id, value) => { const node = document.querySelector(id); node.value = value; node.dispatchEvent(new Event("input", { bubbles: true })); };
    set("#print-paper", "A4");
    const landscape = document.querySelector('input[name="print-orientation"][value="landscape"]'); landscape.checked = true; landscape.dispatchEvent(new Event("input", { bubbles: true }));
    set("#print-margins", "normal");
    const customPages = document.querySelector('input[name="print-pages"][value="custom"]'); customPages.checked = true; customPages.dispatchEvent(new Event("input", { bubbles: true }));
    set("#print-page-range", "2");
    document.querySelector("#print-title-checkbox").click();
    document.querySelector("#print-page-numbers").click();
    return document.querySelector("#print-submit").disabled;
  })()`);
  if (!disabledDuringDebounce) throw new Error("Print remained enabled while a newer debounced preview was pending.");
  await waitFor(cdp, `document.querySelector("#print-preview-status").textContent.startsWith("1 page") && !document.querySelector("#print-submit").disabled`, "updated one-page preview", 45_000);

  const audit = await cdp.evaluate(`(() => {
    const dialog = document.querySelector(".print-dialog").getBoundingClientRect();
    const options = document.querySelector(".print-options").getBoundingClientRect();
    const preview = document.querySelector(".print-preview-pane").getBoundingClientRect();
    const embed = document.querySelector("#print-preview-pdf");
    return {
      visible: !document.querySelector("#print-modal").hidden,
      twoPane: options.right <= preview.left && options.width >= 260 && preview.width >= 360,
      contained: dialog.left >= 0 && dialog.top >= 0 && dialog.right <= innerWidth && dialog.bottom <= innerHeight,
      paper: document.querySelector("#print-paper").value,
      orientation: document.querySelector('input[name="print-orientation"]:checked').value,
      orientationEnabled: [...document.querySelectorAll('input[name="print-orientation"]')].every((input) => !input.disabled),
      margins: document.querySelector("#print-margins").value,
      pages: document.querySelector("#print-page-range").value,
      title: document.querySelector("#print-title-checkbox").checked,
      pageNumbers: document.querySelector("#print-page-numbers").checked,
      previewBlob: embed.src.startsWith("blob:") && !embed.hidden,
      status: document.querySelector("#print-preview-status").textContent,
      submitEnabled: !document.querySelector("#print-submit").disabled,
      printerValue: document.querySelector("#print-printer").value,
      printerLabel: document.querySelector("#print-printer option:checked")?.textContent,
      directCopy: document.querySelector(".print-dialog-footer")?.textContent,
      description: document.querySelector("#print-description").textContent,
    };
  })()`);
  for (const [name, passed] of Object.entries({
    visible: audit.visible,
    twoPane: audit.twoPane,
    contained: audit.contained,
    paper: audit.paper === "A4",
    orientation: audit.orientation === "landscape",
    orientationEnabled: audit.orientationEnabled,
    margins: audit.margins === "normal",
    range: audit.pages === "2",
    title: audit.title,
    pageNumbers: audit.pageNumbers,
    previewBlob: audit.previewBlob,
    pageStatus: audit.status.startsWith("1 page"),
    submitEnabled: audit.submitEnabled,
    windowsDefault: audit.printerValue === "" && audit.printerLabel === "Default Windows printer",
    directOnly: /no second system dialog/i.test(audit.directCopy),
    explanation: /left|right/i.test(audit.description),
  })) {
    if (!passed) throw new Error(`Print dialog UI check failed: ${name}\n${JSON.stringify(audit, null, 2)}`);
  }

  if (process.env.SIMPLE_PRINT_SMOKE_SCREENSHOT) {
    await new Promise((resolve) => setTimeout(resolve, 1_500));
    const screenshot = await cdp.send("Page.captureScreenshot", { format: "png", fromSurface: true });
    await writeFile(path.resolve(process.env.SIMPLE_PRINT_SMOKE_SCREENSHOT), Buffer.from(screenshot.data, "base64"));
  }

  await cdp.evaluate(`document.querySelector("#print-close").focus()`);
  await cdp.send("Input.dispatchKeyEvent", { type: "keyDown", key: "Tab", code: "Tab", windowsVirtualKeyCode: 9, modifiers: 8 });
  await cdp.send("Input.dispatchKeyEvent", { type: "keyUp", key: "Tab", code: "Tab", windowsVirtualKeyCode: 9, modifiers: 8 });
  const trapped = await cdp.evaluate(`document.activeElement === document.querySelector("#print-submit")`);
  if (!trapped) throw new Error("Shift+Tab did not wrap focus to the final print-dialog control.");

  await cdp.send("Input.dispatchKeyEvent", { type: "keyDown", key: "Escape", code: "Escape", windowsVirtualKeyCode: 27 });
  await cdp.send("Input.dispatchKeyEvent", { type: "keyUp", key: "Escape", code: "Escape", windowsVirtualKeyCode: 27 });
  await waitFor(cdp, `document.querySelector("#print-modal").hidden`, "Escape dismissal");
  console.log(JSON.stringify({ ok: true, audit, documentPaperAudit, disabledDuringDebounce, focusTrap: trapped, escapeDismissed: true }, null, 2));
} finally {
  cdp?.close();
  if (!child.killed) child.kill();
  if (child.exitCode === null) {
    await Promise.race([
      new Promise((resolve) => child.once("exit", resolve)),
      new Promise((resolve) => setTimeout(resolve, 3_000)),
    ]);
  }
  for (let attempt = 0; attempt < 6; attempt += 1) {
    try {
      await rm(profile, { recursive: true, force: true, maxRetries: 2, retryDelay: 100 });
      break;
    } catch (error) {
      if (attempt === 5) console.warn(`Could not remove temporary UI profile: ${error.message}`);
      else await new Promise((resolve) => setTimeout(resolve, 250));
    }
  }
}

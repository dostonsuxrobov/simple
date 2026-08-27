import { WordCanvas, type EditorHandle } from "@forevka/wordcanvas";
import { DocumentBuilder } from "@forevka/wordcanvas/builder";
import "./styles.css";

const icon = (paths: string, viewBox = "0 0 24 24") => `<svg aria-hidden="true" viewBox="${viewBox}" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${paths}</svg>`;
const icons = {
  file: icon('<path d="M6 2.75h8l4 4v14.5H6z"/><path d="M14 2.75v4h4"/><path d="M9 12h6M9 15.5h6"/>'),
  plus: icon('<path d="M12 5v14M5 12h14"/>'),
  open: icon('<path d="M3.5 7.5h6l2 2h9l-2.2 9H5.7z"/><path d="M5 7.5V5h6l2 2h5"/>'),
  save: icon('<path d="M5 3.5h12l2 2v15H5z"/><path d="M8 3.5v6h8v-6M8 20.5v-7h8v7"/>'),
  saveAs: icon('<path d="M4.5 3.5h12l2 2v8.5"/><path d="M7.5 3.5v6h8v-6M7.5 20.5v-7h6"/><path d="m15 18 4-4 2 2-4 4-3 .75z"/>'),
  pdf: icon('<path d="M6 2.75h8l4 4v14.5H6z"/><path d="M14 2.75v4h4"/><path d="M8.5 16.5h7M8.5 13.5h4"/>'),
  print: icon('<path d="M7 8V3.5h10V8M7 17H4.5v-7h15v7H17"/><path d="M7 14h10v6.5H7z"/>'),
  focus: icon('<path d="M8 3H3v5M16 3h5v5M21 16v5h-5M3 16v5h5"/>'),
  minimize: icon('<path d="M5 12h14"/>'),
  maximize: icon('<rect x="5" y="5" width="14" height="14"/>'),
  restore: icon('<path d="M8 8V4h12v12h-4M4 8h12v12H4z"/>'),
  close: icon('<path d="m6 6 12 12M18 6 6 18"/>'),
  chevron: icon('<path d="m8 10 4 4 4-4"/>'),
};

document.querySelector<HTMLDivElement>("#app")!.innerHTML = `
  <div class="app-shell">
    <header class="titlebar" role="banner">
      <div class="drag-region title-identity">
        <span class="app-mark">${icons.file}</span>
        <span class="app-name">Simple Docs</span>
        <span class="title-divider" aria-hidden="true"></span>
        <span class="document-title" id="document-title">Untitled document</span>
        <span class="save-state" id="save-state">Ready</span>
      </div>
      <nav class="file-actions" aria-label="Document actions">
        <button class="action-button" id="new-button" title="New window (Ctrl+N)">${icons.plus}<span>New</span></button>
        <button class="action-button" id="open-button" title="Open DOCX (Ctrl+O)">${icons.open}<span>Open</span></button>
        <button class="action-button primary-action" id="save-button" title="Save (Ctrl+S)">${icons.save}<span>Save</span></button>
        <button class="action-button icon-only" id="more-button" title="More document actions" aria-haspopup="menu" aria-expanded="false">${icons.chevron}</button>
        <div class="action-menu" id="action-menu" role="menu" hidden>
          <button role="menuitem" id="save-as-button">${icons.saveAs}<span><strong>Save as</strong><small>Ctrl+Shift+S</small></span></button>
          <button role="menuitem" id="export-pdf-button">${icons.pdf}<span><strong>Export PDF</strong><small>Preserves page layout</small></span></button>
          <button role="menuitem" id="print-button">${icons.print}<span><strong>Print</strong><small>Ctrl+P</small></span></button>
        </div>
      </nav>
      <div class="window-actions">
        <button id="focus-button" title="Focus mode (F11)" aria-pressed="false">${icons.focus}</button>
        <button id="minimize-button" title="Minimize">${icons.minimize}</button>
        <button id="maximize-button" title="Maximize">${icons.maximize}</button>
        <button id="close-button" class="close-button" title="Close">${icons.close}</button>
      </div>
    </header>
    <main class="workspace">
      <div id="editor" class="editor-host" aria-label="Document editor"></div>
      <section class="welcome" id="welcome" aria-labelledby="welcome-title">
        <div class="welcome-card">
          <div class="welcome-brand"><span class="welcome-mark">${icons.file}</span><div><h1 id="welcome-title">Simple Docs</h1><p>Write clearly. Keep your documents yours.</p></div></div>
          <div class="start-actions">
            <button class="start-card primary-start" id="blank-document">
              <span class="start-icon">${icons.plus}</span>
              <span><strong>Blank document</strong><small>Start writing immediately</small></span>
            </button>
            <button class="start-card" id="open-document">
              <span class="start-icon">${icons.open}</span>
              <span><strong>Open DOCX</strong><small>Continue an existing document</small></span>
            </button>
          </div>
          <div class="welcome-section recovery-section" id="recovery-section" hidden>
            <div class="section-heading"><h2>Recovered work</h2><span>Saved locally after an interruption</span></div>
            <div class="file-list" id="recovery-list"></div>
          </div>
          <div class="welcome-section">
            <div class="section-heading"><h2>Recent</h2><span>Files stay on this computer</span></div>
            <div class="file-list" id="recent-list"><p class="empty-list">No recent documents yet.</p></div>
          </div>
          <p class="drop-hint">You can also drop a .docx file anywhere in this window.</p>
        </div>
      </section>
      <div class="loading-overlay" id="loading-overlay">
        <div class="loader-mark">${icons.file}</div>
        <div class="loader-copy"><strong id="loader-title">Preparing Simple Docs</strong><span id="loader-detail">Loading the page layout engine…</span></div>
        <div class="progress-track"><span id="progress-bar"></span></div>
      </div>
      <div class="drop-overlay" id="drop-overlay" hidden><div>${icons.open}<strong>Open in Simple Docs</strong><span>Drop the DOCX file here</span></div></div>
    </main>
    <div class="toast" id="toast" role="status" aria-live="polite" hidden></div>
    <div class="modal-backdrop" id="close-modal" hidden>
      <div class="modal" role="dialog" aria-modal="true" aria-labelledby="close-title">
        <div class="modal-icon">${icons.file}</div>
        <h2 id="close-title">Save your changes?</h2>
        <p>Your latest edits to <strong id="close-document-name">this document</strong> have not been saved.</p>
        <div class="modal-actions"><button id="cancel-close">Cancel</button><button id="discard-close">Discard</button><button class="modal-primary" id="save-close">Save</button></div>
      </div>
    </div>
  </div>`;

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
const welcome = $("welcome");
const editorHost = $("editor");
const loadingOverlay = $("loading-overlay");
const loaderTitle = $("loader-title");
const loaderDetail = $("loader-detail");
const progressBar = $("progress-bar");
const titleLabel = $("document-title");
const saveState = $("save-state");
const toast = $("toast");
const actionMenu = $("action-menu");
const moreButton = $("more-button");
const closeModal = $("close-modal");

let handle: EditorHandle | null = null;
let currentPath: string | null = null;
let documentName = "Untitled document";
let documentOpen = false;
let dirty = false;
let revision = 0;
let saving = false;
let recoveryTimer: number | null = null;
let recoveryMaxTimer: number | null = null;
let recoveryInFlight = false;
let toastTimer: number | null = null;
let pendingExternalPath: string | null = null;
const sessionId = crypto.randomUUID();

const RECOVERY_IDLE_DELAY_MS = 12_000;
const RECOVERY_MAX_DELAY_MS = 60_000;
const RECOVERY_RETRY_DELAY_MS = 2_000;

const editor = new WordCanvas({
  container: editorHost,
  mode: "edit",
  user: { id: "local-user", firstName: "Local", lastName: "Author" },
  view: {
    ruler: true,
    verticalRuler: false,
    outline: true,
    reviewPane: false,
    toolbar: true,
    statusBar: true,
    exportDocx: false,
    exportPdf: false,
    zoom: 1,
  },
  theme: {
    canvasBackground: "#f2f2f2",
    grid: "#d7d7d7",
    externalLink: "#111111",
    accent: "#111111",
    caret: "#111111",
    searchHighlight: "#f59e0b",
    pageGapPx: 18,
    ruler: { bg: "#f7f7f7", content: "#ffffff", line: "#a3a3a3", label: "#525252" },
  },
  overrideDefaultStyles: {
    fontFamily: "Calibri",
    fontSizePx: 14.667,
    color: "#111111",
    lineHeight: 1.15,
  },
  behavior: { zoomStep: 1.1, zoomMin: 0.35, zoomMax: 3 },
  customizeRibbon(api) {
    api.removeTab("file");
  },
  onLoadProgress(progress) {
    progressBar.style.width = `${Math.max(4, progress.percent * 100)}%`;
    if (progress.phase === "fonts") loaderDetail.textContent = "Loading document fonts…";
    if (progress.phase === "ready") loaderDetail.textContent = "Ready";
  },
});

function fileNameWithoutExtension(name: string) {
  return name.replace(/\.docx$/i, "") || "Untitled document";
}

function humanTime(timestamp: number) {
  const elapsed = Date.now() - timestamp;
  if (elapsed < 60_000) return "Just now";
  if (elapsed < 3_600_000) return `${Math.floor(elapsed / 60_000)} min ago`;
  if (elapsed < 86_400_000) return `${Math.floor(elapsed / 3_600_000)} hr ago`;
  return new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric", year: timestamp < Date.now() - 31_536_000_000 ? "numeric" : undefined }).format(timestamp);
}

function escapeHtml(value: string) {
  return value.replace(/[&<>"]/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[character]!);
}

function setTitle() {
  titleLabel.textContent = documentName;
  saveState.textContent = saving ? "Saving…" : dirty ? "Unsaved" : documentOpen ? "Saved locally" : "Ready";
  saveState.classList.toggle("is-dirty", dirty);
  saveState.classList.toggle("is-saving", saving);
  const title = `${dirty ? "• " : ""}${documentName} — Simple Docs`;
  document.title = title;
  window.simpleDocs.setTitle(title);
}

function notify(message: string, tone: "normal" | "error" = "normal") {
  if (toastTimer !== null) window.clearTimeout(toastTimer);
  toast.textContent = message;
  toast.dataset.tone = tone;
  toast.hidden = false;
  toastTimer = window.setTimeout(() => { toast.hidden = true; }, 3200);
}

function showLoading(title: string, detail: string) {
  loaderTitle.textContent = title;
  loaderDetail.textContent = detail;
  progressBar.style.width = "28%";
  loadingOverlay.hidden = false;
}

function hideLoading() {
  progressBar.style.width = "100%";
  window.setTimeout(() => { loadingOverlay.hidden = true; }, 120);
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

function markDirty() {
  if (!documentOpen) return;
  revision += 1;
  if (!dirty) {
    dirty = true;
    setTitle();
  }
  scheduleRecovery();
}

function clearRecoveryTimers() {
  if (recoveryTimer !== null) window.clearTimeout(recoveryTimer);
  if (recoveryMaxTimer !== null) window.clearTimeout(recoveryMaxTimer);
  recoveryTimer = null;
  recoveryMaxTimer = null;
}

function scheduleRecovery(delay = RECOVERY_IDLE_DELAY_MS) {
  if (!dirty || !handle) return;
  if (recoveryTimer !== null) window.clearTimeout(recoveryTimer);
  recoveryTimer = window.setTimeout(() => {
    recoveryTimer = null;
    void checkpointRecovery();
  }, delay);
  if (recoveryMaxTimer === null) {
    recoveryMaxTimer = window.setTimeout(() => {
      recoveryMaxTimer = null;
      if (recoveryTimer !== null) {
        window.clearTimeout(recoveryTimer);
        recoveryTimer = null;
      }
      void checkpointRecovery();
    }, RECOVERY_MAX_DELAY_MS);
  }
}

async function checkpointRecovery() {
  clearRecoveryTimers();
  if (!dirty || !handle) return;
  if (recoveryInFlight) return;
  if (saving) {
    scheduleRecovery(RECOVERY_RETRY_DELAY_MS);
    return;
  }

  recoveryInFlight = true;
  const checkpointRevision = revision;
  let snapshotSaved = false;
  try {
    const blob = await handle.exportDocx();
    await window.simpleDocs.saveRecovery({
      id: sessionId,
      data: new Uint8Array(await blob.arrayBuffer()),
      title: documentName,
      sourcePath: currentPath,
    });
    snapshotSaved = true;
    if (dirty && revision === checkpointRevision) {
      saveState.textContent = "Recovered locally";
      window.setTimeout(setTitle, 1400);
    }
  } catch (error) {
    console.warn("Recovery snapshot failed", error);
  } finally {
    recoveryInFlight = false;
    if (!dirty) {
      clearRecoveryTimers();
      if (snapshotSaved) await window.simpleDocs.clearRecovery(sessionId).catch(() => {});
    } else if (!snapshotSaved || revision !== checkpointRevision) {
      scheduleRecovery(snapshotSaved ? RECOVERY_IDLE_DELAY_MS : RECOVERY_RETRY_DELAY_MS);
    }
  }
}

async function loadPayload(payload: DocumentPayload, recovered = false) {
  if (!handle) return;
  showLoading(recovered ? "Recovering document" : `Opening ${payload.name}`, "Reading document structure…");
  try {
    await handle.openDocx(toArrayBuffer(payload.data));
    currentPath = payload.path;
    documentName = fileNameWithoutExtension(payload.name);
    documentOpen = true;
    dirty = recovered;
    revision += 1;
    welcome.hidden = true;
    editorHost.classList.add("is-active");
    setTitle();
    if (recovered) scheduleRecovery();
    notify(recovered ? "Recovered work is ready." : "Document opened.");
  } catch (error) {
    console.error(error);
    notify(error instanceof Error ? error.message : "This DOCX could not be opened.", "error");
  } finally {
    hideLoading();
  }
}

function beginBlankDocument() {
  documentOpen = true;
  currentPath = null;
  documentName = "Untitled document";
  dirty = false;
  welcome.hidden = true;
  editorHost.classList.add("is-active");
  setTitle();
  window.setTimeout(() => editorHost.focus(), 0);
}

async function chooseOpen() {
  if (documentOpen) {
    await window.simpleDocs.openInNewWindow();
    return;
  }
  const payload = await window.simpleDocs.openFile();
  if (payload) await loadPayload(payload);
}

async function saveDocument(forceDialog = false): Promise<boolean> {
  if (!handle || !documentOpen || saving) return false;
  saving = true;
  setTitle();
  const savingRevision = revision;
  try {
    const blob = await handle.exportDocx();
    const result = await window.simpleDocs.saveDocx({
      data: new Uint8Array(await blob.arrayBuffer()),
      path: currentPath,
      name: `${documentName || "Untitled document"}.docx`,
      forceDialog,
    });
    if (!result) return false;
    currentPath = result.path;
    documentName = fileNameWithoutExtension(result.name);
    if (revision === savingRevision) {
      dirty = false;
      clearRecoveryTimers();
      await window.simpleDocs.clearRecovery(sessionId);
    }
    const savedAllChanges = revision === savingRevision && !dirty;
    if (savedAllChanges) notify("Saved.");
    else {
      dirty = true;
      scheduleRecovery();
      notify("Saved. Newer changes remain unsaved.");
    }
    void refreshWelcomeLists();
    return savedAllChanges;
  } catch (error) {
    console.error(error);
    notify(error instanceof Error ? error.message : "The document could not be saved.", "error");
    return false;
  } finally {
    saving = false;
    if (dirty) scheduleRecovery(RECOVERY_RETRY_DELAY_MS);
    setTitle();
  }
}

async function exportPdf() {
  if (!handle || !documentOpen || saving) return;
  saving = true;
  setTitle();
  try {
    const blob = await handle.exportPdf();
    const result = await window.simpleDocs.savePdf({ data: new Uint8Array(await blob.arrayBuffer()), name: `${documentName}.pdf` });
    if (result) notify("PDF exported.");
  } catch (error) {
    console.error(error);
    notify("PDF export failed.", "error");
  } finally {
    saving = false;
    if (dirty) scheduleRecovery(RECOVERY_RETRY_DELAY_MS);
    setTitle();
  }
}

async function printDocument() {
  if (!handle || !documentOpen || saving) return;
  saving = true;
  setTitle();
  try {
    const blob = await handle.exportPdf();
    await window.simpleDocs.printPdf({ data: new Uint8Array(await blob.arrayBuffer()), name: documentName });
  } catch (error) {
    console.error(error);
    notify("Print preview could not be opened.", "error");
  } finally {
    saving = false;
    if (dirty) scheduleRecovery(RECOVERY_RETRY_DELAY_MS);
    setTitle();
  }
}

async function refreshWelcomeLists() {
  const [recents, recoveries] = await Promise.all([window.simpleDocs.getRecents(), window.simpleDocs.getRecoveries()]);
  const recentList = $("recent-list");
  recentList.innerHTML = recents.length ? recents.map((item) => `
    <button class="file-row" data-recent-path="${escapeHtml(item.path)}">
      <span class="file-row-icon">${icons.file}</span>
      <span class="file-row-copy"><strong>${escapeHtml(fileNameWithoutExtension(item.name))}</strong><small title="${escapeHtml(item.path)}">${escapeHtml(item.path)}</small></span>
      <time>${humanTime(item.openedAt)}</time>
    </button>`).join("") : '<p class="empty-list">No recent documents yet.</p>';
  recentList.querySelectorAll<HTMLButtonElement>("[data-recent-path]").forEach((button) => {
    button.addEventListener("click", async () => {
      const filePath = button.dataset.recentPath!;
      try {
        if (documentOpen) await window.simpleDocs.openInNewWindow(filePath);
        else await loadPayload(await window.simpleDocs.openPath(filePath));
      } catch {
        await window.simpleDocs.removeRecent(filePath);
        notify("That file is no longer available.", "error");
        void refreshWelcomeLists();
      }
    });
  });

  const recoverySection = $("recovery-section");
  const recoveryList = $("recovery-list");
  recoverySection.hidden = recoveries.length === 0;
  recoveryList.innerHTML = recoveries.map((item) => `
    <button class="file-row recovery-row" data-recovery-id="${escapeHtml(item.id)}">
      <span class="file-row-icon">${icons.file}</span>
      <span class="file-row-copy"><strong>${escapeHtml(item.title)}</strong><small>Unsaved recovery copy</small></span>
      <time>${humanTime(item.updatedAt)}</time>
    </button>`).join("");
  recoveryList.querySelectorAll<HTMLButtonElement>("[data-recovery-id]").forEach((button) => {
    button.addEventListener("click", async () => {
      try {
        const payload = await window.simpleDocs.loadRecovery(button.dataset.recoveryId!);
        await loadPayload(payload, true);
      } catch {
        notify("That recovery copy could not be opened.", "error");
      }
    });
  });
}

function toggleMenu(force?: boolean) {
  const shouldOpen = force ?? actionMenu.hidden;
  actionMenu.hidden = !shouldOpen;
  moreButton.setAttribute("aria-expanded", String(shouldOpen));
}

function requestClose() {
  if (!dirty || !documentOpen) {
    window.simpleDocs.confirmClose();
    return;
  }
  $("close-document-name").textContent = documentName;
  closeModal.hidden = false;
  $("save-close").focus();
}

$("blank-document").addEventListener("click", beginBlankDocument);
$("open-document").addEventListener("click", () => void chooseOpen());
$("new-button").addEventListener("click", () => void window.simpleDocs.newWindow());
$("open-button").addEventListener("click", () => void chooseOpen());
$("save-button").addEventListener("click", () => void saveDocument(false));
$("save-as-button").addEventListener("click", () => { toggleMenu(false); void saveDocument(true); });
$("export-pdf-button").addEventListener("click", () => { toggleMenu(false); void exportPdf(); });
$("print-button").addEventListener("click", () => { toggleMenu(false); void printDocument(); });
moreButton.addEventListener("click", (event) => { event.stopPropagation(); toggleMenu(); });
document.addEventListener("pointerdown", (event) => { if (!actionMenu.contains(event.target as Node) && event.target !== moreButton) toggleMenu(false); });
$("minimize-button").addEventListener("click", () => window.simpleDocs.minimize());
$("maximize-button").addEventListener("click", () => window.simpleDocs.toggleMaximize());
$("close-button").addEventListener("click", requestClose);
$("focus-button").addEventListener("click", () => window.simpleDocs.toggleFullscreen());
$("cancel-close").addEventListener("click", () => { closeModal.hidden = true; });
$("discard-close").addEventListener("click", async () => { await window.simpleDocs.clearRecovery(sessionId); window.simpleDocs.confirmClose(); });
$("save-close").addEventListener("click", async () => { if (await saveDocument(false)) window.simpleDocs.confirmClose(); });

document.addEventListener("keydown", (event) => {
  const modifier = event.ctrlKey || event.metaKey;
  if (modifier && event.key.toLowerCase() === "s") {
    event.preventDefault();
    void saveDocument(event.shiftKey);
    return;
  }
  if (modifier && event.key.toLowerCase() === "o") {
    event.preventDefault();
    void chooseOpen();
    return;
  }
  if (modifier && event.key.toLowerCase() === "n") {
    event.preventDefault();
    void window.simpleDocs.newWindow();
    return;
  }
  if (modifier && event.key.toLowerCase() === "p") {
    event.preventDefault();
    void printDocument();
    return;
  }
  if (event.key === "F11") {
    event.preventDefault();
    window.simpleDocs.toggleFullscreen();
    return;
  }
  if (event.key === "Escape" && document.body.classList.contains("focus-mode")) {
    event.preventDefault();
    window.simpleDocs.toggleFullscreen();
    return;
  }
  if (!documentOpen || closeModal.hidden === false) return;
  if (isNavigationOnlyControl(event.target)) return;
  const key = event.key.toLowerCase();
  const directMutation = !modifier && (event.key.length === 1 || ["backspace", "delete", "enter", "tab"].includes(key));
  const shortcutMutation = modifier && ["x", "v", "z", "y", "b", "i", "u"].includes(key);
  if (directMutation || shortcutMutation) markDirty();
}, true);

document.addEventListener("paste", (event) => { if (!isNavigationOnlyControl(event.target)) markDirty(); }, true);
document.addEventListener("cut", (event) => { if (!isNavigationOnlyControl(event.target)) markDirty(); }, true);
editorHost.addEventListener("beforeinput", (event) => { if (!isNavigationOnlyControl(event.target)) markDirty(); }, true);
editorHost.addEventListener("compositionend", (event) => { if (!isNavigationOnlyControl(event.target)) markDirty(); }, true);
editorHost.addEventListener("drop", (event) => {
  const opensDocument = [...event.dataTransfer?.files ?? []].some((file) => /\.docx$/i.test(file.name));
  if (!opensDocument) window.setTimeout(markDirty, 0);
}, true);

function isNavigationOnlyControl(target: EventTarget | null) {
  if (!(target instanceof Element)) return false;
  if (target.closest(".cw-outline, .cw-statusbar, .cw-ruler-row, .cw-vruler, .cw-review, .cw-mode-select")) return true;
  const control = target.closest<HTMLElement>("button, [role=button], select, input");
  if (!control) return false;
  const label = (control.getAttribute("aria-label") || control.getAttribute("title") || control.getAttribute("placeholder") || control.textContent || "").trim().toLowerCase();
  return ["home", "insert", "layout", "table", "view", "review", "editing", "suggesting", "viewing", "find", "replace", "replace all", "select all"].includes(label)
    || label.startsWith("zoom ");
}

editorHost.addEventListener("change", (event) => {
  if (!isNavigationOnlyControl(event.target)) markDirty();
}, true);
editorHost.addEventListener("click", (event) => {
  if (!isNavigationOnlyControl(event.target) && (event.target as Element).closest("button, [role=button], select, input[type=color], input[type=checkbox]")) {
    window.setTimeout(markDirty, 0);
  }
}, true);

let dragDepth = 0;
window.addEventListener("dragenter", (event) => {
  if ([...event.dataTransfer?.items ?? []].some((item) => item.kind === "file")) {
    dragDepth += 1;
    $("drop-overlay").hidden = false;
  }
});
window.addEventListener("dragleave", () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) $("drop-overlay").hidden = true;
});
window.addEventListener("dragover", (event) => event.preventDefault());
window.addEventListener("drop", async (event) => {
  event.preventDefault();
  dragDepth = 0;
  $("drop-overlay").hidden = true;
  const file = [...event.dataTransfer?.files ?? []].find((candidate) => /\.docx$/i.test(candidate.name));
  if (!file) {
    notify("Simple Docs opens .docx files.", "error");
    return;
  }
  const filePath = window.simpleDocs.pathForFile(file);
  if (documentOpen && filePath) await window.simpleDocs.openInNewWindow(filePath);
  else if (filePath) await loadPayload(await window.simpleDocs.openPath(filePath));
  else await loadPayload({ data: new Uint8Array(await file.arrayBuffer()), name: file.name, path: null, size: file.size });
});

window.simpleDocs.onMaximized((maximized) => { $("maximize-button").innerHTML = maximized ? icons.restore : icons.maximize; });
window.simpleDocs.onFullscreen((fullscreen) => {
  document.body.classList.toggle("focus-mode", fullscreen);
  $("focus-button").setAttribute("aria-pressed", String(fullscreen));
});
window.simpleDocs.onCloseRequested(requestClose);
window.simpleDocs.onOpenExternal(async (filePath) => {
  if (!handle) {
    pendingExternalPath = filePath;
    return;
  }
  try {
    await loadPayload(await window.simpleDocs.openPath(filePath));
  } catch (error) {
    console.error(error);
    notify("This DOCX could not be opened.", "error");
    hideLoading();
  }
});

async function initialize() {
  handle = await editor.whenReady();
  handle.setDocument(
    DocumentBuilder.create()
      .paragraph("", { fontFamily: "Calibri", fontSizePx: 14.667, color: "#111111" })
      .spacing({ lineHeight: 1.15, after: 10.667 })
      .build(),
  );
  editor.on("reviewChanged", markDirty);
  loadingOverlay.hidden = true;
  await refreshWelcomeLists();
  setTitle();
  if (pendingExternalPath) {
    const filePath = pendingExternalPath;
    pendingExternalPath = null;
    await loadPayload(await window.simpleDocs.openPath(filePath));
  }
}

void initialize().catch((error) => {
  console.error("Simple Docs failed to initialize", error);
  loaderTitle.textContent = "Simple Docs could not start";
  loaderDetail.textContent = error instanceof Error ? error.message : "The page layout engine failed to load.";
});

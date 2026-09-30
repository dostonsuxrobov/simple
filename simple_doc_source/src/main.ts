import { WordCanvas, type EditorHandle, type RibbonActionContext } from "@forevka/wordcanvas";
import { DocumentBuilder, pt } from "@forevka/wordcanvas/builder";
import { serializeDocument, prepareDocumentForExport, type StructuredDocumentExportFormat } from "./document-export.js";
import { activeModal, canCommitPrintPreview, documentExportDisabled, isModalBlockedShortcut, nextPrintPreviewGeneration } from "./ui-guards.js";
import "./styles.css";

const icon = (paths: string, viewBox = "0 0 24 24") => `<svg aria-hidden="true" viewBox="${viewBox}" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${paths}</svg>`;
const icons = {
  file: icon('<path d="M6 2.75h8l4 4v14.5H6z"/><path d="M14 2.75v4h4"/><path d="M9 12h6M9 15.5h6"/>'),
  plus: icon('<path d="M12 5v14M5 12h14"/>'),
  open: icon('<path d="M3.5 7.5h6l2 2h9l-2.2 9H5.7z"/><path d="M5 7.5V5h6l2 2h5"/>'),
  save: icon('<path d="M5 3.5h12l2 2v15H5z"/><path d="M8 3.5v6h8v-6M8 20.5v-7h8v7"/>'),
  saveAs: icon('<path d="M4.5 3.5h12l2 2v8.5"/><path d="M7.5 3.5v6h8v-6M7.5 20.5v-7h6"/><path d="m15 18 4-4 2 2-4 4-3 .75z"/>'),
  export: icon('<path d="M12 3v12M7.5 10.5 12 15l4.5-4.5"/><path d="M5 14.5v5h14v-5"/>'),
  pdf: icon('<path d="M6 2.75h8l4 4v14.5H6z"/><path d="M14 2.75v4h4"/><path d="M8.5 16.5h7M8.5 13.5h4"/>'),
  web: icon('<circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3c2.4 2.5 3.5 5.5 3.5 9S14.4 18.5 12 21M12 3C9.6 5.5 8.5 8.5 8.5 12S9.6 18.5 12 21"/>'),
  markdown: icon('<path d="M3 6.5h18v11H3z"/><path d="M6 14v-4l2.5 2.5L11 10v4M14 11.5l2 2 2-2M16 9v4.5"/>'),
  text: icon('<path d="M5 5h14M12 5v14M8.5 19h7"/>'),
  print: icon('<path d="M7 8V3.5h10V8M7 17H4.5v-7h15v7H17"/><path d="M7 14h10v6.5H7z"/>'),
  focus: icon('<path d="M8 3H3v5M16 3h5v5M21 16v5h-5M3 16v5h5"/>'),
  minimize: icon('<path d="M5 12h14"/>'),
  maximize: icon('<rect x="5" y="5" width="14" height="14"/>'),
  restore: icon('<path d="M8 8V4h12v12h-4M4 8h12v12H4z"/>'),
  close: icon('<path d="m6 6 12 12M18 6 6 18"/>'),
  chevron: icon('<path d="m8 10 4 4 4-4"/>'),
  envelope: icon('<rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3.5 6.5 8.5 6.5 8.5-6.5"/>'),
};

const ENVELOPE_SIZES = [
  { id: "no10", label: "#10 (9.5 × 4.125 in)", pageWidthPx: 912, pageHeightPx: 396 },
  { id: "monarch", label: "Monarch (7.5 × 3.875 in)", pageWidthPx: 720, pageHeightPx: 372 },
  { id: "dl", label: "DL (220 × 110 mm)", pageWidthPx: 832, pageHeightPx: 416 },
  { id: "c5", label: "C5 (229 × 162 mm)", pageWidthPx: 866, pageHeightPx: 612 },
];

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
        <button class="action-button" id="open-button" title="Open Word document (Ctrl+O)">${icons.open}<span>Open</span></button>
        <button class="action-button primary-action" id="save-button" title="Save (Ctrl+S)">${icons.save}<span>Save</span></button>
        <div class="export-control">
          <button class="action-button export-action" id="export-as-button" title="Open or create a document to export" aria-haspopup="menu" aria-expanded="false" aria-disabled="true" disabled>${icons.export}<span>Export As</span>${icons.chevron}</button>
          <div class="action-menu export-menu" id="export-menu" role="menu" hidden>
            <div class="menu-heading">Convert document to</div>
            <button role="menuitem" data-export-format="docx">${icons.file}<span><strong>Word document</strong><small>Editable .docx copy</small></span></button>
            <button role="menuitem" data-export-format="pdf">${icons.pdf}<span><strong>PDF document</strong><small>Fixed, printable page layout</small></span></button>
            <button role="menuitem" data-export-format="html">${icons.web}<span><strong>Web page</strong><small>Text, tables and embedded images</small></span></button>
            <button role="menuitem" data-export-format="md">${icons.markdown}<span><strong>Markdown</strong><small>Headings, lists and tables</small></span></button>
            <button role="menuitem" data-export-format="txt">${icons.text}<span><strong>Plain text</strong><small>Readable text without formatting</small></span></button>
          </div>
        </div>
        <button class="action-button icon-only" id="more-button" title="More document actions" aria-haspopup="menu" aria-expanded="false">${icons.chevron}</button>
        <div class="action-menu" id="action-menu" role="menu" hidden>
          <button role="menuitem" id="save-as-button">${icons.saveAs}<span><strong>Save as</strong><small>Ctrl+Shift+S</small></span></button>
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
    <div class="document-layout-bar" id="document-layout-bar" hidden><div role="group" aria-label="Document view"><button id="original-layout-button" aria-pressed="true">Page view</button><button id="edit-layout-button" aria-pressed="false">Edit document</button></div><select id="source-view-scale" aria-label="Page view zoom"><option value="Fit">Fit page</option><option value="FitH">Fit width</option><option value="100">100%</option><option value="150">150%</option></select><span id="document-layout-note">A local Office engine prepares the document for viewing and printing.</span></div>
    <div class="document-compatibility" id="document-compatibility" role="status" hidden></div>
    <main class="workspace">
      <div id="editor" class="editor-host" aria-label="Document editor"></div>
      <div id="original-layout-view" class="original-layout-view" hidden><p id="original-layout-loading">Preparing the document page view…</p><iframe id="original-layout-pdf" title="Document page view" hidden></iframe></div>
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
              <span><strong>Open Word document</strong><small>Open .docx or import legacy .doc</small></span>
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
          <p class="drop-hint">You can also drop a .docx or legacy .doc file anywhere in this window.</p>
        </div>
      </section>
      <div class="loading-overlay" id="loading-overlay">
        <div class="loader-mark">${icons.file}</div>
        <div class="loader-copy"><strong id="loader-title">Preparing Simple Docs</strong><span id="loader-detail">Loading the page layout engine…</span></div>
        <div class="progress-track"><span id="progress-bar"></span></div>
      </div>
      <div class="drop-overlay" id="drop-overlay" hidden><div>${icons.open}<strong>Open in Simple Docs</strong><span>Drop the Word file here</span></div></div>
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
    <div class="modal-backdrop" id="envelope-modal" hidden>
      <div class="modal" role="dialog" aria-modal="true" aria-labelledby="envelope-title">
        <div class="modal-icon">${icons.envelope}</div>
        <h2 id="envelope-title">Create an envelope</h2>
        <p>The envelope replaces the document in this window. Both addresses stay editable afterwards.</p>
        <div class="envelope-fields">
          <label>Envelope size<select id="envelope-size">${ENVELOPE_SIZES.map((size) => `<option value="${size.id}">${size.label}</option>`).join("")}</select></label>
          <label>Return address<textarea id="envelope-return" rows="3" spellcheck="false"></textarea></label>
          <label>Delivery address<textarea id="envelope-delivery" rows="4" spellcheck="false"></textarea></label>
        </div>
        <div class="modal-actions"><button id="envelope-cancel">Cancel</button><button class="modal-primary" id="envelope-create">Create envelope</button></div>
      </div>
    </div>
    <div class="modal-backdrop print-backdrop" id="print-modal" hidden>
      <section class="print-dialog" role="dialog" aria-modal="true" aria-labelledby="print-title" aria-describedby="print-description">
        <header class="print-dialog-header">
          <div>
            <span class="print-eyebrow">Print setup</span>
            <h2 id="print-title">Preview and print</h2>
            <p id="print-description">Choose the page layout on the left. The finished print appears on the right.</p>
          </div>
          <button class="print-close" id="print-close" type="button" aria-label="Close print setup">${icons.close}</button>
        </header>
        <div class="print-dialog-body">
          <form class="print-options" id="print-options" novalidate>
            <fieldset class="print-option-group print-grid-group">
              <legend>Printer</legend>
              <label>Printer<select id="print-printer" data-print-job disabled><option value="">Looking for printers…</option></select></label>
              <p class="print-option-note" id="print-printer-note" aria-live="polite">Reading the printers available in Windows…</p>
              <label>Copies<input id="print-copies" data-print-job type="number" min="1" max="999" step="1" value="1"></label>
              <label>Color<select id="print-color" data-print-job><option value="color">Color</option><option value="monochrome">Black and white</option></select></label>
              <label>Two-sided<select id="print-duplex" data-print-job><option value="simplex">One-sided</option><option value="longEdge">Flip on long edge</option><option value="shortEdge">Flip on short edge</option></select></label>
              <label class="print-toggle"><input id="print-collate" data-print-job type="checkbox" checked><span>Collate multiple copies</span></label>
            </fieldset>
            <fieldset class="print-option-group">
              <legend>Pages</legend>
              <label class="print-choice"><input type="radio" name="print-pages" value="all" checked><span><strong>All pages</strong><small id="print-all-pages">Entire document</small></span></label>
              <label class="print-choice"><input type="radio" name="print-pages" value="custom"><span><strong>Custom range</strong><small>For example: 1-3, 5</small></span></label>
              <label class="print-inline-field" for="print-page-range"><span>Pages</span><input id="print-page-range" type="text" inputmode="numeric" placeholder="1-3, 5" autocomplete="off" disabled aria-describedby="print-range-error"></label>
              <p class="print-field-error" id="print-range-error" role="alert" hidden></p>
            </fieldset>

            <fieldset class="print-option-group print-grid-group">
              <legend>Paper</legend>
              <label>Size<select id="print-paper"><option value="Document">Document size · recommended</option><option value="Letter">Letter · 8.5 × 11 in</option><option value="A4">A4 · 210 × 297 mm</option><option value="Legal">Legal · 8.5 × 14 in</option><option value="A5">A5 · 148 × 210 mm</option><option value="Tabloid">Tabloid · 11 × 17 in</option></select></label>
              <span class="print-control-label" id="orientation-label">Orientation</span>
              <div class="print-segmented" role="radiogroup" aria-labelledby="orientation-label">
                <label><input type="radio" name="print-orientation" value="portrait" checked><span><i class="paper-icon portrait" aria-hidden="true"></i>Portrait</span></label>
                <label><input type="radio" name="print-orientation" value="landscape"><span><i class="paper-icon landscape" aria-hidden="true"></i>Landscape</span></label>
              </div>
              <p class="print-option-note" id="print-document-paper-note">Document size preserves each selected page’s original size and orientation.</p>
              <label>Margins<select id="print-margins"><option value="normal">Normal · 0.75 in</option><option value="narrow">Narrow · 0.35 in</option><option value="wide">Wide · 1 in</option><option value="none">None</option><option value="custom">Custom…</option></select></label>
              <div class="print-custom-margins" id="print-custom-margins" hidden>
                <label>Top<input id="print-margin-top" type="number" min="0" max="3" step="0.05" value="0.75"><span>in</span></label>
                <label>Right<input id="print-margin-right" type="number" min="0" max="3" step="0.05" value="0.75"><span>in</span></label>
                <label>Bottom<input id="print-margin-bottom" type="number" min="0" max="3" step="0.05" value="0.75"><span>in</span></label>
                <label>Left<input id="print-margin-left" type="number" min="0" max="3" step="0.05" value="0.75"><span>in</span></label>
              </div>
            </fieldset>

            <fieldset class="print-option-group print-grid-group">
              <legend>Placement</legend>
              <label>Scale<select id="print-scaling"><option value="fit">Fit to printable area</option><option value="actual">Actual size · 100%</option><option value="custom">Custom scale…</option></select></label>
              <label class="print-percent-field" id="print-percent-field" hidden>Scale<input id="print-scale-percent" type="number" min="25" max="200" step="5" value="100"><span>%</span></label>
              <label class="print-toggle"><input id="print-center-content" type="checkbox" checked><span>Center on paper</span></label>
            </fieldset>

            <fieldset class="print-option-group">
              <legend>Page marks</legend>
              <label class="print-toggle"><input id="print-title-checkbox" type="checkbox"><span>Document title at top</span></label>
              <label class="print-toggle"><input id="print-page-numbers" type="checkbox"><span>Page numbers at bottom</span></label>
              <p class="print-option-note">Document colors, images, and existing headers or footers are preserved.</p>
            </fieldset>
          </form>

          <section class="print-preview-pane" aria-labelledby="print-preview-title">
            <div class="print-preview-heading">
              <div><span class="print-eyebrow">Final output</span><h3 id="print-preview-title">Paper preview</h3></div>
              <span class="print-preview-status" id="print-preview-status" aria-live="polite">Preparing…</span>
            </div>
            <div class="print-preview-stage" id="print-preview-stage" aria-busy="true">
              <iframe class="print-preview-pdf" id="print-preview-pdf" title="Final paginated print preview" hidden></iframe>
              <div class="print-preview-placeholder" id="print-preview-placeholder">
                <span class="print-preview-spinner" aria-hidden="true"></span>
                <strong>Building your print</strong>
                <small>Pagination and paper fit will appear here.</small>
              </div>
            </div>
            <p class="print-preview-error" id="print-preview-error" role="alert" hidden></p>
          </section>
        </div>
        <footer class="print-dialog-footer">
          <p><strong>What you see is what Simple sends.</strong> Print goes directly to the selected printer—no second system dialog.</p>
          <div class="print-footer-actions"><button id="print-cancel" type="button">Cancel</button><button class="modal-primary print-submit" id="print-submit" type="submit" form="print-options" disabled>${icons.print}<span>Print</span></button></div>
        </footer>
      </section>
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
const exportMenu = $("export-menu");
const exportAsButton = $<HTMLButtonElement>("export-as-button");
const closeModal = $("close-modal");
const envelopeModal = $("envelope-modal");
const printModal = $("print-modal");

let handle: EditorHandle | null = null;
let currentPath: string | null = null;
let currentFormat: 'doc' | 'docx' = 'docx';
let currentSourceHash: string | undefined;
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
let envelopeContext: RibbonActionContext | null = null;
let printBasePdf: Uint8Array | null = null;
let printComposition: PrintComposition | null = null;
let printPreviewUrl: string | null = null;
let printPreviewTimer: number | null = null;
let printPreviewGeneration = 0;
let printPrinterGeneration = 0;
let printPrinters: PrinterSummary[] | null = null;
let printSubmitting = false;
let printReturnFocus: HTMLElement | null = null;
type DocumentSnapshot = { revision: number; model: ReturnType<EditorHandle['getDocument']>; review: string };
let originalDocument: DocumentSnapshot & { data: Uint8Array; sourceData?: Uint8Array } | null = null;
let cachedDocumentPdf: DocumentSnapshot & { data: Uint8Array } | null = null;
let sourceLayout: DocumentSnapshot & { pdf: Promise<Uint8Array>; url?: string } | null = null;
const sessionId = crypto.randomUUID();

const RECOVERY_IDLE_DELAY_MS = 12_000;
const RECOVERY_MAX_DELAY_MS = 60_000;
const RECOVERY_RETRY_DELAY_MS = 2_000;

const installedFonts = await window.simpleDocs.getDocumentFonts().catch(() => []);
const editor = new WordCanvas({
  container: editorHost,
  fonts: { fonts: installedFonts, disableBuiltin: installedFonts.map((font) => font.family) },
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
    api.addGroup("layout", { id: "simple.envelopes", label: "Envelopes" });
    api.addButton("simple.envelopes", {
      id: "simple.envelopes.create",
      label: "Envelope…",
      tooltip: "Create an envelope",
      onClick: (ctx) => openEnvelopeDialog(ctx),
    });
  },
  onLoadProgress(progress) {
    progressBar.style.width = `${Math.max(4, progress.percent * 100)}%`;
    if (progress.phase === "fonts") loaderDetail.textContent = "Loading document fonts…";
    if (progress.phase === "ready") loaderDetail.textContent = "Ready";
  },
});

function fileNameWithoutExtension(name: string) {
  return name.replace(/\.docx?$/i, "") || "Untitled document";
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
  const exportDisabled = documentExportDisabled(documentOpen, saving);
  exportAsButton.disabled = exportDisabled;
  exportAsButton.setAttribute("aria-disabled", String(exportDisabled));
  exportAsButton.title = !documentOpen
    ? "Open or create a document to export"
    : saving
      ? "Wait for the current document operation to finish"
      : "Export As (Ctrl+Shift+E)";
  exportMenu.querySelectorAll<HTMLButtonElement>("[data-export-format]").forEach((button) => { button.disabled = exportDisabled; });
  if (exportDisabled) toggleExportMenu(false);
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
  cachedDocumentPdf = null;
  if (sourceLayout) $("document-layout-note").textContent = "Edits are in Edit document. Printing and PDF export include your current edits.";
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
    const importedLegacyDocument = payload.convertedFrom === "doc";
    const textOnlyImport = importedLegacyDocument && payload.conversionMethod !== "layout";
    const renamedModernDocument = payload.convertedFrom === "docx-renamed";
    const requiresSaveAs = textOnlyImport || renamedModernDocument || payload.requiresSaveAs === true;
    currentPath = requiresSaveAs ? null : payload.path;
    currentFormat = payload.format === 'doc' ? 'doc' : 'docx';
    currentSourceHash = payload.sourceHash;
    documentName = fileNameWithoutExtension(payload.name);
    documentOpen = true;
    dirty = recovered || requiresSaveAs;
    revision += 1;
    originalDocument = { ...documentSnapshot(), data: new Uint8Array(payload.data), sourceData: payload.sourceData };
    cachedDocumentPdf = null;
    welcome.hidden = true;
    editorHost.classList.add("is-active");
    const compatibility = $("document-compatibility");
    compatibility.hidden = !importedLegacyDocument;
    compatibility.classList.toggle("is-warning", textOnlyImport);
    compatibility.textContent = textOnlyImport
      ? "Text-only import · This .doc file’s formatting, tables and images could not be preserved. Open a DOCX copy from Word/Google Docs, or install LibreOffice for formatted import. Your original file is unchanged."
      : importedLegacyDocument
        ? "Legacy Word document · Save keeps the .doc format. Save as or Export As creates a DOCX copy. Complex layouts may change after editing."
        : "";
    resetOriginalLayout();
    if (payload.originalLayout) {
      const layout = { ...documentSnapshot(), pdf: window.simpleDocs.getOriginalLayoutPdf(payload.originalLayout), url: undefined as string | undefined };
      sourceLayout = layout;
      $("document-layout-bar").hidden = false;
      setOriginalLayoutView(true);
      void layout.pdf.then((data) => {
        if (sourceLayout !== layout) return;
        layout.url = URL.createObjectURL(new Blob([toArrayBuffer(data)], { type: "application/pdf" }));
        const preview = $<HTMLIFrameElement>("original-layout-pdf");
        preview.src = `${layout.url}#view=Fit&toolbar=0`;
        preview.hidden = false;
        $("original-layout-loading").hidden = true;
      }).catch((error) => {
        if (sourceLayout !== layout) return;
        $("original-layout-loading").textContent = `The page view could not be prepared. ${error instanceof Error ? error.message : ""} Choose Edit document to continue.`;
      });
    }
    setTitle();
    if (dirty) scheduleRecovery();
    notify(recovered
      ? "Recovered work is ready."
      : importedLegacyDocument
        ? textOnlyImport ? "Text-only copy opened. See the import notice above the document." : "Legacy .doc opened with formatted conversion."
        : renamedModernDocument
          ? "This .doc contains a modern DOCX document. Save creates a .docx copy."
        : "Document opened.");
  } catch (error) {
    console.error(error);
    notify(error instanceof Error ? error.message : "This Word document could not be opened.", "error");
  } finally {
    hideLoading();
  }
}

function beginBlankDocument() {
  resetOriginalLayout();
  $("document-compatibility").hidden = true;
  originalDocument = null;
  cachedDocumentPdf = null;
  documentOpen = true;
  currentPath = null;
  currentFormat = 'docx';
  currentSourceHash = undefined;
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
  const savingModel = handle.getDocument();
  try {
    const data = await currentDocxBytes();
    const result = await window.simpleDocs.saveDocx({
      data,
      path: currentPath,
      name: `${documentName || "Untitled document"}.docx`,
      forceDialog,
      format: currentFormat,
      sourceData: matchesDocumentSnapshot(originalDocument) ? originalDocument?.sourceData : undefined,
      expectedHash: currentSourceHash,
      protectOriginal: Boolean(sourceLayout && !matchesDocumentSnapshot(sourceLayout)),
    });
    if (!result) return false;
    currentPath = result.path;
    currentFormat = result.format === 'doc' ? 'doc' : 'docx';
    currentSourceHash = result.sourceHash;
    documentName = fileNameWithoutExtension(result.name);
    if (revision === savingRevision && handle.getDocument() === savingModel) {
      dirty = false;
      originalDocument = { ...documentSnapshot(), data, sourceData: result.sourceData };
      clearRecoveryTimers();
      await window.simpleDocs.clearRecovery(sessionId);
    }
    const savedAllChanges = revision === savingRevision && handle.getDocument() === savingModel && !dirty;
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

type DocumentExportFormat = "docx" | "pdf" | StructuredDocumentExportFormat;

function documentSnapshot(): DocumentSnapshot {
  if (!handle) throw new Error("Open a document first.");
  return { revision, model: handle.getDocument(), review: JSON.stringify(handle.getReview()) };
}

function matchesDocumentSnapshot(snapshot: DocumentSnapshot | null): boolean {
  return Boolean(snapshot && handle && snapshot.model === handle.getDocument() && snapshot.review === JSON.stringify(handle.getReview()));
}

async function currentDocxBytes(): Promise<Uint8Array> {
  if (originalDocument && matchesDocumentSnapshot(originalDocument)) return originalDocument.data;
  if (!handle) throw new Error("Open a document first.");
  return new Uint8Array(await (await handle.exportDocx()).arrayBuffer());
}

async function currentPdfBytes(): Promise<Uint8Array> {
  if (sourceLayout && matchesDocumentSnapshot(sourceLayout)) return sourceLayout.pdf;
  if (cachedDocumentPdf && matchesDocumentSnapshot(cachedDocumentPdf)) return cachedDocumentPdf.data;
  if (!handle) throw new Error("Open a document first.");
  const exportRevision = revision;
  const exportModel = handle.getDocument();
  const data = new Uint8Array(await (await handle.exportPdf()).arrayBuffer());
  if (revision === exportRevision && handle.getDocument() === exportModel) cachedDocumentPdf = { ...documentSnapshot(), data };
  return data;
}

function resetOriginalLayout() {
  if (sourceLayout?.url) URL.revokeObjectURL(sourceLayout.url);
  sourceLayout = null;
  $("document-layout-bar").hidden = true;
  $("original-layout-view").hidden = true;
  $("original-layout-loading").hidden = false;
  $("original-layout-loading").textContent = "Preparing the document page view…";
  const preview = $<HTMLIFrameElement>("original-layout-pdf");
  preview.hidden = true;
  preview.removeAttribute("src");
  $<HTMLSelectElement>("source-view-scale").value = 'Fit';
  editorHost.hidden = false;
}

function setOriginalLayoutView(original: boolean) {
  $("original-layout-view").hidden = !original;
  editorHost.hidden = original;
  $("source-view-scale").hidden = !original;
  $("original-layout-button").setAttribute("aria-pressed", String(original));
  $("edit-layout-button").setAttribute("aria-pressed", String(!original));
  $("document-layout-note").textContent = original
    ? matchesDocumentSnapshot(sourceLayout) ? "Page view uses local Office rendering. Complex formatting can differ from Word or Google Docs." : "Source file for reference. Printing and PDF export include your current edits."
    : "Editing view · Floating objects and complex tables can reflow. Use Save as to keep a separate original copy.";
}

$("original-layout-button").addEventListener("click", () => setOriginalLayoutView(true));
$("edit-layout-button").addEventListener("click", () => setOriginalLayoutView(false));
$("source-view-scale").addEventListener("change", () => {
  if (!sourceLayout?.url) return;
  const value = $<HTMLSelectElement>("source-view-scale").value;
  $<HTMLIFrameElement>("original-layout-pdf").src = `${sourceLayout.url}#${value.startsWith('Fit') ? `view=${value}` : `zoom=${value}`}&toolbar=0`;
});

const EXPORT_LABELS: Record<DocumentExportFormat, string> = {
  docx: "Word document",
  pdf: "PDF",
  html: "web page",
  md: "Markdown",
  txt: "plain text",
};

async function exportDocument(format: DocumentExportFormat) {
  if (!handle || !documentOpen || saving) return;
  saving = true;
  setTitle();
  try {
    let data: Uint8Array;
    if (format === "docx") {
      data = await currentDocxBytes();
    } else if (format === "pdf") {
      data = await currentPdfBytes();
    } else {
      const document = format === "txt" ? handle.getDocument() : await prepareDocumentForExport(handle.getDocument());
      data = new TextEncoder().encode(serializeDocument(document, format, documentName));
    }
    const result = await window.simpleDocs.saveExport({ data, name: documentName, format });
    if (result) notify(`${EXPORT_LABELS[format]} exported.`);
  } catch (error) {
    console.error(error);
    notify(error instanceof Error ? error.message : `${EXPORT_LABELS[format]} export failed.`, "error");
  } finally {
    saving = false;
    if (dirty) scheduleRecovery(RECOVERY_RETRY_DELAY_MS);
    setTitle();
  }
}

function printRadioValue(name: string, fallback: string) {
  return document.querySelector<HTMLInputElement>(`input[name="${name}"]:checked`)?.value ?? fallback;
}

function printNumberValue(id: string, fallback: number) {
  const value = Number($<HTMLInputElement>(id).value);
  return Number.isFinite(value) ? value : fallback;
}

function selectedPrintPrinter() {
  const name = $<HTMLSelectElement>("print-printer").value;
  if (!printPrinters?.length) return null;
  if (!name) return { name: "", displayName: "Default Windows printer", supportsDuplex: true, supportsColor: true };
  return printPrinters.find((printer) => printer.name === name) ?? null;
}

function updatePrintSubmitState() {
  $<HTMLButtonElement>("print-submit").disabled = printSubmitting || printComposition === null || selectedPrintPrinter() === null;
}

function setPrintJobControls() {
  const printer = selectedPrintPrinter();
  const color = $<HTMLSelectElement>("print-color");
  const duplex = $<HTMLSelectElement>("print-duplex");
  const copies = Math.min(999, Math.max(1, Math.trunc(printNumberValue("print-copies", 1))));
  $<HTMLInputElement>("print-copies").value = String(copies);
  color.disabled = Boolean(printer && !printer.supportsColor);
  if (color.disabled) color.value = "monochrome";
  duplex.disabled = Boolean(printer && !printer.supportsDuplex);
  if (duplex.disabled) duplex.value = "simplex";
  $<HTMLInputElement>("print-collate").disabled = copies < 2;
  updatePrintSubmitState();
}

async function refreshPrintPrinters() {
  const generation = ++printPrinterGeneration;
  const select = $<HTMLSelectElement>("print-printer");
  const note = $("print-printer-note");
  select.disabled = true;
  select.replaceChildren(new Option("Looking for printers…", ""));
  note.textContent = "Reading the printers available in Windows…";
  printPrinters = null;
  updatePrintSubmitState();
  try {
    const printers = await window.simpleDocs.listPrinters();
    if (generation !== printPrinterGeneration || printModal.hidden) return;
    printPrinters = printers;
    select.replaceChildren();
    const hasPrinters = printers.length > 0;
    select.disabled = !hasPrinters;
    if (hasPrinters) {
      select.add(new Option("Default Windows printer", ""));
      for (const printer of printers) select.add(new Option(printer.displayName, printer.name));
      select.value = "";
      note.textContent = "Print sends directly to this printer without opening another dialog.";
    } else {
      select.add(new Option("No printer available", ""));
      note.textContent = "No printer is available. Add or enable a printer in Windows, then try again.";
    }
  } catch (error) {
    if (generation !== printPrinterGeneration || printModal.hidden) return;
    printPrinters = [];
    select.replaceChildren(new Option("Printers unavailable", ""));
    note.textContent = error instanceof Error
      ? `Simple could not read the Windows printers: ${error.message}`
      : "Simple could not read the Windows printers.";
  }
  setPrintJobControls();
}

function readPrintSettings(): PrintSettings {
  return {
    paper: $<HTMLSelectElement>("print-paper").value as PrintSettings["paper"],
    orientation: printRadioValue("print-orientation", "portrait") as PrintSettings["orientation"],
    margins: {
      preset: $<HTMLSelectElement>("print-margins").value as PrintSettings["margins"]["preset"],
      custom: {
        top: printNumberValue("print-margin-top", 0.75),
        right: printNumberValue("print-margin-right", 0.75),
        bottom: printNumberValue("print-margin-bottom", 0.75),
        left: printNumberValue("print-margin-left", 0.75),
      },
    },
    scaling: $<HTMLSelectElement>("print-scaling").value as PrintSettings["scaling"],
    scalePercent: printNumberValue("print-scale-percent", 100),
    pages: printRadioValue("print-pages", "all") as PrintSettings["pages"],
    pageRange: $<HTMLInputElement>("print-page-range").value,
    printTitle: $<HTMLInputElement>("print-title-checkbox").checked,
    printPageNumbers: $<HTMLInputElement>("print-page-numbers").checked,
    centerContent: $<HTMLInputElement>("print-center-content").checked,
  };
}

function setPrintControls() {
  const pages = printRadioValue("print-pages", "all");
  const customRange = $<HTMLInputElement>("print-page-range");
  customRange.disabled = pages !== "custom";
  const documentPaper = $<HTMLSelectElement>("print-paper").value === "Document";
  document.querySelectorAll<HTMLInputElement>('input[name="print-orientation"]').forEach((input) => { input.disabled = documentPaper; });
  $("orientation-label").textContent = documentPaper ? "Orientation · follows document" : "Orientation";
  $("print-document-paper-note").hidden = !documentPaper;
  $<HTMLElement>("print-custom-margins").hidden = $<HTMLSelectElement>("print-margins").value !== "custom";
  $<HTMLElement>("print-percent-field").hidden = $<HTMLSelectElement>("print-scaling").value !== "custom";
}

function resetPrintControls() {
  const section = handle?.getDocument().section;
  const orientation = section && section.pageWidthPx > section.pageHeightPx ? "landscape" : "portrait";
  document.querySelector<HTMLInputElement>(`input[name="print-orientation"][value="${orientation}"]`)!.checked = true;
  document.querySelector<HTMLInputElement>('input[name="print-pages"][value="all"]')!.checked = true;
  $<HTMLSelectElement>("print-paper").value = "Document";
  $<HTMLSelectElement>("print-margins").value = "none";
  $<HTMLSelectElement>("print-scaling").value = "fit";
  $<HTMLInputElement>("print-page-range").value = "";
  $<HTMLInputElement>("print-scale-percent").value = "100";
  $<HTMLInputElement>("print-title-checkbox").checked = false;
  $<HTMLInputElement>("print-page-numbers").checked = false;
  $<HTMLInputElement>("print-center-content").checked = true;
  $<HTMLInputElement>("print-copies").value = "1";
  $<HTMLSelectElement>("print-color").value = "color";
  $<HTMLSelectElement>("print-duplex").value = "simplex";
  $<HTMLInputElement>("print-collate").checked = true;
  setPrintControls();
}

function releasePrintPreview() {
  if (printPreviewUrl) URL.revokeObjectURL(printPreviewUrl);
  printPreviewUrl = null;
  const preview = $<HTMLIFrameElement>("print-preview-pdf");
  preview.removeAttribute("src");
  preview.hidden = true;
}

function setPrintPreviewLoading(message = "Preparing…") {
  printComposition = null;
  $("print-preview-stage").setAttribute("aria-busy", "true");
  $("print-preview-status").textContent = message;
  $("print-preview-error").hidden = true;
  $("print-preview-placeholder").hidden = false;
  updatePrintSubmitState();
}

function schedulePrintPreview(immediate = false) {
  setPrintControls();
  printPreviewGeneration = nextPrintPreviewGeneration(printPreviewGeneration);
  const generation = printPreviewGeneration;
  if (printPreviewTimer !== null) window.clearTimeout(printPreviewTimer);
  printPreviewTimer = null;
  if (!printBasePdf || printModal.hidden) return;
  setPrintPreviewLoading("Updating…");
  printPreviewTimer = window.setTimeout(() => {
    printPreviewTimer = null;
    void updatePrintPreview(generation);
  }, immediate ? 0 : 220);
}

async function updatePrintPreview(generation: number) {
  const basePdf = printBasePdf;
  if (!basePdf || !canCommitPrintPreview(generation, printPreviewGeneration, printModal.hidden)) return;
  setPrintPreviewLoading("Rendering…");
  const rangeError = $("print-range-error");
  rangeError.hidden = true;
  try {
    const composition = await window.simpleDocs.composePrintPdf({ data: basePdf, name: documentName, settings: readPrintSettings() });
    if (!canCommitPrintPreview(generation, printPreviewGeneration, printModal.hidden)) return;
    printComposition = composition;
    if (composition.settings.scaling === "custom") $<HTMLInputElement>("print-scale-percent").value = String(composition.settings.scalePercent);
    if (composition.settings.margins.preset === "custom") {
      $<HTMLInputElement>("print-margin-top").value = String(composition.settings.margins.top);
      $<HTMLInputElement>("print-margin-right").value = String(composition.settings.margins.right);
      $<HTMLInputElement>("print-margin-bottom").value = String(composition.settings.margins.bottom);
      $<HTMLInputElement>("print-margin-left").value = String(composition.settings.margins.left);
    }
    releasePrintPreview();
    printPreviewUrl = URL.createObjectURL(new Blob([composition.data], { type: "application/pdf" }));
    const preview = $<HTMLIFrameElement>("print-preview-pdf");
    preview.src = `${printPreviewUrl}#toolbar=0&navpanes=0&view=FitH`;
    preview.hidden = false;
    $("print-preview-placeholder").hidden = true;
    $("print-preview-stage").setAttribute("aria-busy", "false");
    const clipped = composition.placements.some((placement) => placement.clipped);
    $("print-preview-status").textContent = `${composition.outputPageCount} ${composition.outputPageCount === 1 ? "page" : "pages"}${composition.mixedPaperSizes ? " · mixed paper" : ""}${clipped ? " · clipped" : ""}`;
    $("print-all-pages").textContent = `${composition.sourcePageCount} ${composition.sourcePageCount === 1 ? "page" : "pages"}`;
    const previewError = $("print-preview-error");
    const warnings = [];
    if (clipped) warnings.push("Some content extends outside the printable area. Choose Fit to printable area or reduce the custom scale.");
    if (composition.mixedPaperSizes) warnings.push("This document mixes page sizes or orientations. The preview preserves every page, but Windows accepts one paper setting per print job, so the printer may fit pages to its default paper.");
    previewError.hidden = warnings.length === 0;
    previewError.textContent = warnings.join(" ");
    updatePrintSubmitState();
  } catch (error) {
    if (!canCommitPrintPreview(generation, printPreviewGeneration, printModal.hidden)) return;
    printComposition = null;
    releasePrintPreview();
    $("print-preview-placeholder").hidden = false;
    $("print-preview-stage").setAttribute("aria-busy", "false");
    $("print-preview-status").textContent = "Needs attention";
    const message = error instanceof Error ? error.message : "The print preview could not be built.";
    const previewError = $("print-preview-error");
    previewError.textContent = message;
    previewError.hidden = false;
    if (readPrintSettings().pages === "custom") {
      rangeError.textContent = message;
      rangeError.hidden = false;
    }
  }
}

function closePrintDialog(force = false) {
  if (printSubmitting && !force) return;
  printPreviewGeneration = nextPrintPreviewGeneration(printPreviewGeneration);
  printPrinterGeneration += 1;
  if (printPreviewTimer !== null) window.clearTimeout(printPreviewTimer);
  printPreviewTimer = null;
  printBasePdf = null;
  printComposition = null;
  releasePrintPreview();
  printModal.hidden = true;
  const returnFocus = printReturnFocus;
  printReturnFocus = null;
  window.setTimeout(() => returnFocus?.focus(), 0);
}

async function submitPrint() {
  const composition = printComposition;
  if (!composition || printSubmitting) return;
  printSubmitting = true;
  const submit = $<HTMLButtonElement>("print-submit");
  submit.disabled = true;
  submit.querySelector("span")!.textContent = "Printing…";
  try {
    const printer = selectedPrintPrinter();
    if (!printer) throw new Error("Choose an available printer before printing.");
    const copies = Math.min(999, Math.max(1, Math.trunc(printNumberValue("print-copies", 1))));
    const result = await window.simpleDocs.printPdf({
      data: composition.data,
      name: documentName,
      deviceName: printer.name,
      copies,
      color: $<HTMLSelectElement>("print-color").value === "color",
      collate: copies > 1 && $<HTMLInputElement>("print-collate").checked,
      duplexMode: $<HTMLSelectElement>("print-duplex").value as "simplex" | "shortEdge" | "longEdge",
      paper: composition.settings.paper,
      orientation: composition.settings.orientation,
      paperWidth: composition.paperWidth,
      paperHeight: composition.paperHeight,
      mixedPaperSizes: composition.mixedPaperSizes,
    });
    if (result.success) {
      closePrintDialog(true);
      notify(`Sent to ${result.printerLabel || printer.displayName}.`);
    } else {
      notify(result.failureReason || "The printer did not accept the job.", "error");
    }
  } catch (error) {
    console.error(error);
    notify(error instanceof Error ? error.message : "The document could not be sent to the printer.", "error");
  } finally {
    printSubmitting = false;
    submit.querySelector("span")!.textContent = "Print";
    if (!printModal.hidden) updatePrintSubmitState();
  }
}

async function printDocument() {
  if (!handle || !documentOpen || saving || !printModal.hidden) return;
  toggleMenu(false);
  toggleExportMenu(false);
  printReturnFocus = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  resetPrintControls();
  printModal.hidden = false;
  setPrintPreviewLoading();
  $<HTMLSelectElement>("print-printer").focus();
  void refreshPrintPrinters();
  saving = true;
  setTitle();
  try {
    const data = await currentPdfBytes();
    if (printModal.hidden) return;
    printBasePdf = data;
    schedulePrintPreview(true);
  } catch (error) {
    console.error(error);
    closePrintDialog(true);
    notify(error instanceof Error ? error.message : "Print setup could not be opened.", "error");
  } finally {
    saving = false;
    if (dirty) scheduleRecovery(RECOVERY_RETRY_DELAY_MS);
    setTitle();
  }
}

function addressLines(value: string) {
  return value.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
}

function documentIsEmpty(ctx: RibbonActionContext) {
  return ctx.getDocument().blocks.every((block) => block.kind === "paragraph" && block.runs.every((run) => !run.text.trim()));
}

function openEnvelopeDialog(ctx: RibbonActionContext) {
  envelopeContext = ctx;
  envelopeModal.hidden = false;
  $<HTMLTextAreaElement>("envelope-return").focus();
}

function applyEnvelope(ctx: RibbonActionContext, size: (typeof ENVELOPE_SIZES)[number], returnAddress: string, deliveryAddress: string) {
  const margins = { top: 48, right: 72, bottom: 48, left: 72 };
  const returnLines = addressLines(returnAddress);
  const deliveryLines = addressLines(deliveryAddress);
  const builder = DocumentBuilder.create({ pageSize: { pageWidthPx: size.pageWidthPx, pageHeightPx: size.pageHeightPx }, margins });
  for (const line of returnLines.length ? returnLines : [""]) {
    builder.paragraph(line, { fontFamily: "Calibri", fontSizePx: pt(10), color: "#111111" }).spacing({ lineHeight: 1.15, before: 0, after: 0 });
  }
  const returnBlockPx = Math.max(returnLines.length, 1) * Math.round(pt(10) * 1.4);
  const deliverySpaceBeforePx = Math.max(24, Math.round(size.pageHeightPx / 2 - margins.top - returnBlockPx));
  const deliveryIndentPx = Math.round(size.pageWidthPx * 0.42);
  (deliveryLines.length ? deliveryLines : [""]).forEach((line, index) => {
    builder.paragraph(line, { fontFamily: "Calibri", fontSizePx: pt(12), color: "#111111" })
      .indent({ left: deliveryIndentPx })
      .spacing({ lineHeight: 1.15, before: index === 0 ? deliverySpaceBeforePx : 0, after: 0 });
  });
  ctx.setDocument(builder.build());
}

function createEnvelope() {
  const ctx = envelopeContext;
  if (!ctx) return;
  if (saving) {
    notify("A save is in progress. Try again when it finishes.", "error");
    return;
  }
  const sizeId = $<HTMLSelectElement>("envelope-size").value;
  const size = ENVELOPE_SIZES.find((candidate) => candidate.id === sizeId) ?? ENVELOPE_SIZES[0];
  if (documentOpen && (dirty || !documentIsEmpty(ctx)) && !window.confirm(`Replace "${documentName}" with the envelope? Unsaved changes will be lost.`)) return;
  applyEnvelope(ctx, size, $<HTMLTextAreaElement>("envelope-return").value, $<HTMLTextAreaElement>("envelope-delivery").value);
  resetOriginalLayout();
  envelopeModal.hidden = true;
  currentPath = null;
  currentFormat = 'docx';
  currentSourceHash = undefined;
  documentName = "Envelope";
  documentOpen = true;
  welcome.hidden = true;
  editorHost.classList.add("is-active");
  markDirty();
  setTitle();
  notify("Envelope created.");
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
  if (shouldOpen) toggleExportMenu(false);
  actionMenu.hidden = !shouldOpen;
  moreButton.setAttribute("aria-expanded", String(shouldOpen));
}

function toggleExportMenu(force?: boolean) {
  const shouldOpen = force ?? exportMenu.hidden;
  if (shouldOpen && exportAsButton.disabled) return;
  if (shouldOpen) {
    actionMenu.hidden = true;
    moreButton.setAttribute("aria-expanded", "false");
  }
  exportMenu.hidden = !shouldOpen;
  exportAsButton.setAttribute("aria-expanded", String(shouldOpen));
  if (shouldOpen) exportMenu.querySelector<HTMLButtonElement>("button")?.focus();
}

function requestClose() {
  if (!printModal.hidden) {
    if (printSubmitting) return;
    closePrintDialog(true);
  }
  if (!dirty || !documentOpen) {
    window.simpleDocs.confirmClose();
    return;
  }
  $("close-document-name").textContent = documentName;
  closeModal.hidden = false;
  $("save-close").focus();
}

function trapModalFocus(event: KeyboardEvent, modal: HTMLElement) {
  if (event.key !== "Tab") return false;
  const focusable = [...modal.querySelectorAll<HTMLElement>('button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])')]
    .filter((element) => !element.hidden && element.getClientRects().length > 0);
  if (!focusable.length) return false;
  const first = focusable[0];
  const last = focusable[focusable.length - 1];
  if (event.shiftKey && document.activeElement === first) {
    event.preventDefault();
    last.focus();
    return true;
  }
  if (!event.shiftKey && document.activeElement === last) {
    event.preventDefault();
    first.focus();
    return true;
  }
  if (!modal.contains(document.activeElement)) {
    event.preventDefault();
    first.focus();
    return true;
  }
  return false;
}

$("blank-document").addEventListener("click", beginBlankDocument);
$("open-document").addEventListener("click", () => void chooseOpen());
$("new-button").addEventListener("click", () => void window.simpleDocs.newWindow());
$("open-button").addEventListener("click", () => void chooseOpen());
$("save-button").addEventListener("click", () => void saveDocument(false));
$("save-as-button").addEventListener("click", () => { toggleMenu(false); void saveDocument(true); });
$("print-button").addEventListener("click", () => { toggleMenu(false); void printDocument(); });
moreButton.addEventListener("click", (event) => { event.stopPropagation(); toggleMenu(); });
exportAsButton.addEventListener("click", (event) => { event.stopPropagation(); toggleExportMenu(); });
exportMenu.querySelectorAll<HTMLButtonElement>("[data-export-format]").forEach((button) => {
  button.addEventListener("click", () => {
    const format = button.dataset.exportFormat as DocumentExportFormat;
    toggleExportMenu(false);
    void exportDocument(format);
  });
});
document.addEventListener("pointerdown", (event) => {
  const target = event.target as Node;
  if (!actionMenu.contains(target) && target !== moreButton) toggleMenu(false);
  if (!exportMenu.contains(target) && !exportAsButton.contains(target)) toggleExportMenu(false);
});
$("minimize-button").addEventListener("click", () => window.simpleDocs.minimize());
$("maximize-button").addEventListener("click", () => window.simpleDocs.toggleMaximize());
$("close-button").addEventListener("click", requestClose);
$("focus-button").addEventListener("click", () => window.simpleDocs.toggleFullscreen());
$("cancel-close").addEventListener("click", () => { closeModal.hidden = true; $("close-button").focus(); });
$("discard-close").addEventListener("click", async () => { await window.simpleDocs.clearRecovery(sessionId); window.simpleDocs.confirmClose(); });
$("save-close").addEventListener("click", async () => { if (await saveDocument(false)) window.simpleDocs.confirmClose(); });
$("envelope-cancel").addEventListener("click", () => { envelopeModal.hidden = true; envelopeContext = null; editorHost.focus(); });
$("envelope-create").addEventListener("click", createEnvelope);
$<HTMLFormElement>("print-options").addEventListener("submit", (event) => { event.preventDefault(); void submitPrint(); });
$<HTMLFormElement>("print-options").addEventListener("input", (event) => {
  const target = event.target;
  if (target instanceof HTMLElement && target.hasAttribute("data-print-job")) {
    setPrintJobControls();
    return;
  }
  schedulePrintPreview();
});
$("print-close").addEventListener("click", () => closePrintDialog());
$("print-cancel").addEventListener("click", () => closePrintDialog());
$<HTMLSelectElement>("print-margins").addEventListener("change", () => {
  if ($<HTMLSelectElement>("print-margins").value === "custom") $<HTMLInputElement>("print-margin-top").focus();
});
$<HTMLSelectElement>("print-scaling").addEventListener("change", () => {
  if ($<HTMLSelectElement>("print-scaling").value === "custom") $<HTMLInputElement>("print-scale-percent").focus();
});
$<HTMLInputElement>("print-page-range").addEventListener("input", () => { $("print-range-error").hidden = true; });

document.addEventListener("keydown", (event) => {
  const modifier = event.ctrlKey || event.metaKey;
  const modal = activeModal(document);
  if (modal) {
    if (trapModalFocus(event, modal as HTMLElement)) return;
    if (event.key === "Escape") {
      if (modal === closeModal) {
        event.preventDefault();
        closeModal.hidden = true;
        $("close-button").focus();
      } else if (modal === envelopeModal) {
        event.preventDefault();
        envelopeModal.hidden = true;
        envelopeContext = null;
        editorHost.focus();
      } else if (modal === printModal) {
        event.preventDefault();
        if (!printSubmitting) closePrintDialog();
      }
      return;
    }
    if (isModalBlockedShortcut(event)) {
      event.preventDefault();
      event.stopImmediatePropagation();
    }
    return;
  }
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
  if (modifier && event.shiftKey && event.key.toLowerCase() === "e") {
    event.preventDefault();
    toggleExportMenu();
    return;
  }
  if (event.key === "Escape" && (!actionMenu.hidden || !exportMenu.hidden)) {
    event.preventDefault();
    toggleMenu(false);
    toggleExportMenu(false);
    exportAsButton.focus();
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
  if (!documentOpen || closeModal.hidden === false || envelopeModal.hidden === false) return;
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
  const opensDocument = [...event.dataTransfer?.files ?? []].some((file) => /\.docx?$/i.test(file.name));
  if (!opensDocument) window.setTimeout(markDirty, 0);
}, true);

function isNavigationOnlyControl(target: EventTarget | null) {
  if (!(target instanceof Element)) return false;
  if (target.closest(".cw-outline, .cw-statusbar, .cw-ruler-row, .cw-vruler, .cw-review, .cw-mode-select, .modal-backdrop")) return true;
  const control = target.closest<HTMLElement>("button, [role=button], select, input");
  if (!control) return false;
  const label = (control.getAttribute("aria-label") || control.getAttribute("title") || control.getAttribute("placeholder") || control.textContent || "").trim().toLowerCase();
  return ["home", "insert", "layout", "table", "view", "review", "editing", "suggesting", "viewing", "find", "replace", "replace all", "select all", "envelope…", "create an envelope"].includes(label)
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

// WordCanvas renders its semantic right-click menu in a body-level portal. Those
// commands do not consistently emit beforeinput/cut/paste events through the
// editor host, so compare the document on menu activation and only mark a real
// mutation. Copy, Select all, disabled items, and submenu navigation remain clean.
document.addEventListener("click", (event) => {
  if (!(event.target instanceof Element)) return;
  const menuItem = event.target.closest<HTMLElement>(".cw-menu-item");
  if (!menuItem || !documentOpen || !handle) return;

  const activeHandle = handle;
  const startingRevision = revision;
  let before: string;
  try {
    before = JSON.stringify(activeHandle.getDocument());
  } catch {
    return;
  }

  let mutationObserved = false;
  const detectMutation = () => {
    if (mutationObserved || handle !== activeHandle || !documentOpen) return;
    try {
      if (JSON.stringify(activeHandle.getDocument()) === before) return;
      mutationObserved = true;
      if (revision === startingRevision) markDirty();
    } catch {
      // A document replacement or teardown can race the deferred comparison.
    }
  };

  window.setTimeout(detectMutation, 0);
  window.setTimeout(detectMutation, 180);
}, true);

window.addEventListener("beforeunload", releasePrintPreview);
window.addEventListener("beforeunload", () => { if (sourceLayout?.url) URL.revokeObjectURL(sourceLayout.url); });
window.simpleDocs.onDocumentShortcut((action) => {
  if (activeModal(document) || saving) return;
  if (action === 'print') void printDocument();
  else if (action === 'save') void saveDocument();
  else if (action === 'save-as') void saveDocument(true);
  else if (action === 'export') toggleExportMenu();
});

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
  const file = [...event.dataTransfer?.files ?? []].find((candidate) => /\.docx?$/i.test(candidate.name));
  if (!file) {
    notify("Simple Docs opens .docx and legacy .doc files.", "error");
    return;
  }
  const filePath = window.simpleDocs.pathForFile(file);
  if (documentOpen && filePath) await window.simpleDocs.openInNewWindow(filePath);
  else if (filePath) await loadPayload(await window.simpleDocs.openPath(filePath));
  else await loadPayload(await window.simpleDocs.openBytes({ data: new Uint8Array(await file.arrayBuffer()), name: file.name }));
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
    notify("This Word document could not be opened.", "error");
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

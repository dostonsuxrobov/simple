import { WordCanvas, type CustomFontDef, type DocSelection, type EditMode, type EditorHandle, type RibbonActionContext } from "@forevka/wordcanvas";
import { DocumentBuilder, pt } from "@forevka/wordcanvas/builder";
import type { BandContainer, CharStyle, ParaStyle } from "@forevka/wordcanvas/query";
import { installAutoFormat, type AutoFormatInstallation } from "./autoformat";
import { serializeDocumentWithWarnings, prepareDocumentForExport, type StructuredDocumentExportFormat } from "./document-export.js";
import { createEngineBridge, RIBBON_ITEM_IDS, type EngineBridge } from "./engine-bridge";
import {
  activeModal,
  canCommitPrintPreview,
  cleanErrorMessage,
  describeFidelityItems,
  describeReviewItems,
  documentContentKey,
  documentExportDisabled,
  isMissingFileError,
  isModalBlockedShortcut,
  nextPrintPreviewGeneration,
  notificationDuration,
  openErrorMessage,
  recoveredDocumentName,
  reviewContentKey,
  reviewSummary,
  scanDocxFidelity,
  type FidelityItem,
  type ReviewSummary,
} from "./ui-guards.js";
import { importDocument, type ImportFormat } from "./importers/index.ts";
import { imageNaturalSize, sniffImageType } from "./importers/images.ts";
import {
  buildPasteTransaction,
  fragmentFromClipboard,
  fragmentText,
  imageOnlyHtml,
  jpegOrientation,
  looksLikeAddress,
  mergeFragment,
  pastedLink,
  type ClipboardContent,
  type PasteContent,
  type PasteDocument,
  type PasteMode,
} from "./importers/paste.ts";
import { authorDisplayName, authorIdentity, cleanAuthorName, readStoredAuthorName, resolveAuthorName, storeAuthorName, windowsUserFromAppUrl } from "./ui/author.ts";
import { createCommandSearch, describeRibbonTitle, type SearchCommand } from "./ui/command-search.ts";
import { findParagraph, firstParagraphId, headingStyleDefinition, isCollapsed, paragraphText, resolveStyleChar, runStyleAt, selectionRanges, type DocumentLike, type StylesheetLike } from "./ui/editing.ts";
import { createEditorCommands, EDITOR_SEARCH_COMMANDS, type EditorCommands } from "./ui/editor-commands.ts";
import { cancelFormDialog, openFormDialog } from "./ui/form-dialog.ts";
import { formatCombo, shortcutLetter, withShortcutHint } from "./ui/keys.ts";
import { findShortcut, ribbonItemShortcut, shortcutFor, type ShortcutCommand } from "./ui/shortcuts.ts";
import { createFontPicker, type FontPicker } from "./ui/font-picker.ts";
import { documentFontFamilies, familyKey, normalizeFamily, readRecentFonts, rememberDocumentFonts, rememberRecentFont, startupFontRequests, type FontCatalogView, type FontEntry } from "./ui/font-list.ts";
import {
  bandCaretTarget,
  bandCharStyle,
  currentPageNumberOptions,
  differentFirstPageOps,
  documentHasPageNumbers,
  ensureBandOps,
  hasBand,
  hasDifferentFirstPage,
  insertPageNumberOps,
  removeBandOps,
  removePageNumberOps,
  type BandKind,
  type HeaderFooterDocument,
} from "./ui/header-footer.ts";
import { openPageNumberDialog } from "./ui/page-number-dialog.ts";
import { openParagraphDialog } from "./ui/paragraph-dialog.ts";
import { paragraphPatch, spaceToggle, unitForLocale } from "./ui/paragraph-format.ts";
import { closePopupMenu, openPopupMenu } from "./ui/popup-menu.ts";
import { cancelSettingsDialog } from "./ui/settings-dialog.ts";
import { createSpelling, type Spelling } from "./ui/spelling.ts";
import { getSimpleIO } from "./simple-io/io-client";
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
  undo: icon('<path d="M9 14 4.5 9.5 9 5"/><path d="M4.5 9.5h10a5 5 0 0 1 0 10H11"/>'),
  redo: icon('<path d="m15 14 4.5-4.5L15 5"/><path d="M19.5 9.5h-10a5 5 0 0 0 0 10H13"/>'),
  search: icon('<circle cx="10.5" cy="10.5" r="6"/><path d="m15 15 5 5"/>'),
  link: icon('<path d="M10 14a4.5 4.5 0 0 0 6.4.1l3-3a4.5 4.5 0 0 0-6.4-6.4l-1.6 1.6"/><path d="M14 10a4.5 4.5 0 0 0-6.4-.1l-3 3a4.5 4.5 0 0 0 6.4 6.4l1.6-1.6"/>'),
  comment: icon('<path d="M4.5 5h15v10.5h-8l-4.5 3.75V15.5H4.5z"/><path d="M8.5 9.25h7M8.5 12h4.5"/>'),
  bookmark: icon('<path d="M7 3.5h10v17l-5-3.75-5 3.75z"/>'),
  list: icon('<path d="M9 7h11M9 12h11M9 17h11"/><path d="M4.5 7h.01M4.5 12h.01M4.5 17h.01" stroke-width="2.6"/>'),
  person: icon('<circle cx="12" cy="8.5" r="3.75"/><path d="M4.75 20c.9-3.6 3.7-5.5 7.25-5.5s6.35 1.9 7.25 5.5"/>'),
  clipboard: icon('<path d="M8.5 4.5h-2v16h11v-16h-2"/><path d="M9 3h6v3H9z"/><path d="M9.5 11h5M9.5 14.5h5"/>'),
  check: icon('<path d="m5 12.5 4.5 4.5L19 7.5"/>'),
  spelling: icon('<path d="M4 15 7.5 5h1L12 15M5.4 11.5h5.2"/><path d="m13.5 13.5 2.5 2.5 4.5-6"/><path d="M3.5 19.5c1.2-1 2.3-1 3.5 0s2.3 1 3.5 0 2.3-1 3.5 0 2.3 1 3.5 0 2.3-1 3.5 0"/>'),
  autoformat: icon('<path d="M6 4.5h7.5M6 4.5v8"/><path d="M15.5 4.5c-1.2.7-1.8 1.7-1.8 3h1.8v2.2h-2.4"/><path d="M19.8 4.5c-1.2.7-1.8 1.7-1.8 3h1.8v2.2h-2.4"/><path d="m4 18 2.5 2.5L11 16"/><path d="M13.5 18.5h7"/>'),
  header: icon('<path d="M5 3.5h14v17H5z"/><path d="M5 8.5h14" stroke-dasharray="2 2"/><path d="M8 6h8"/><path d="M8 12h8M8 15h6"/>'),
  footer: icon('<path d="M5 3.5h14v17H5z"/><path d="M5 15.5h14" stroke-dasharray="2 2"/><path d="M8 18h8"/><path d="M8 7h8M8 10h6"/>'),
  pageNumber: icon('<path d="M5 3.5h14v17H5z"/><path d="M8 7h8M8 10h8"/><path d="M11 18.5v-4l-1.2.8M10 18.5h2.2"/>'),
  paragraph: icon('<path d="M4 5h16M8 9h12M8 13h12M4 17h16"/><path d="m4 9.5 2 1.5-2 1.5"/>'),
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
      <div class="command-search" id="command-search" role="search">
        <span class="command-search-icon" aria-hidden="true">${icons.search}</span>
        <input id="command-search-input" type="text" placeholder="Search commands" aria-label="Search commands" title="Search commands (Alt+Q)" role="combobox" aria-expanded="false" aria-controls="command-search-list" aria-autocomplete="list" autocomplete="off" spellcheck="false">
        <kbd class="command-search-key" aria-hidden="true">Alt+Q</kbd>
        <div class="command-search-list" id="command-search-list" role="listbox" aria-label="Commands" hidden></div>
      </div>
      <nav class="file-actions" aria-label="Document actions">
        <div class="history-actions" role="group" aria-label="Undo and redo">
          <button class="action-button icon-only history-button" id="undo-button" title="Undo (Ctrl+Z)" aria-label="Undo" disabled>${icons.undo}</button>
          <button class="action-button icon-only history-button" id="redo-button" title="Redo (Ctrl+Y)" aria-label="Redo" disabled>${icons.redo}</button>
        </div>
        <span class="action-divider" aria-hidden="true"></span>
        <button class="action-button" id="new-button" title="New window (Ctrl+N)">${icons.plus}<span>New</span></button>
        <button class="action-button" id="open-button" title="Open a document (Ctrl+O)">${icons.open}<span>Open</span></button>
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
          <button role="menuitem" id="save-as-button">${icons.saveAs}<span><strong>Save as</strong><small>F12 or Ctrl+Shift+S</small></span></button>
          <button role="menuitem" id="print-button">${icons.print}<span><strong>Print</strong><small>Ctrl+P</small></span></button>
          <button role="menuitem" id="author-button" title="The name shown on your comments and suggestions">${icons.person}<span><strong>Author name</strong><small id="author-name-label">Author</small></span></button>
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
    <div class="document-compatibility" id="document-compatibility" role="status" hidden><span id="document-notice-text"></span><button class="notice-close" id="document-notice-close" type="button" title="Dismiss" aria-label="Dismiss this notice">${icons.close}</button></div>
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
              <span><strong>Open document</strong><small>Word, OpenDocument, RTF, web page, Markdown or text</small></span>
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
          <p class="drop-hint">You can also drop a document anywhere in this window.</p>
        </div>
      </section>
      <div class="loading-overlay" id="loading-overlay">
        <div class="loader-mark">${icons.file}</div>
        <div class="loader-copy"><strong id="loader-title">Preparing Simple Docs</strong><span id="loader-detail">Loading the page layout engine…</span></div>
        <div class="progress-track"><span id="progress-bar"></span></div>
      </div>
      <div class="drop-overlay" id="drop-overlay" hidden><div>${icons.open}<strong id="drop-overlay-title">Open in Simple Docs</strong><span id="drop-overlay-detail">Drop the document here</span></div></div>
    </main>
    <div class="paste-options" id="paste-options" hidden>
      <button class="paste-options-button" id="paste-options-button" type="button" title="Paste options (Ctrl)" aria-label="Paste options" aria-haspopup="menu" aria-expanded="false" aria-controls="paste-options-menu">${icons.clipboard}${icons.chevron}</button>
      <div class="paste-options-menu" id="paste-options-menu" role="menu" aria-label="Paste options" hidden>
        <button type="button" role="menuitemradio" aria-checked="true" data-paste-mode="keep">${icons.check}<span>Keep source formatting</span></button>
        <button type="button" role="menuitemradio" aria-checked="false" data-paste-mode="merge">${icons.check}<span>Merge formatting</span></button>
        <button type="button" role="menuitemradio" aria-checked="false" data-paste-mode="text">${icons.check}<span>Keep text only</span><kbd>Ctrl+Shift+V</kbd></button>
      </div>
    </div>
    <div class="toast" id="toast" role="status" aria-live="polite" hidden></div>
    <div class="modal-backdrop" id="close-modal" hidden>
      <div class="modal" role="dialog" aria-modal="true" aria-labelledby="close-title">
        <div class="modal-icon">${icons.file}</div>
        <h2 id="close-title">Save your changes?</h2>
        <p>Your latest edits to <strong id="close-document-name">this document</strong> have not been saved.</p>
        <p class="modal-note" id="close-review-note" hidden></p>
        <div class="modal-actions"><button id="cancel-close">Cancel</button><button id="discard-close">Discard</button><button class="modal-primary" id="save-close">Save</button></div>
      </div>
    </div>
    <div class="modal-backdrop" id="choice-modal" hidden>
      <div class="modal" role="dialog" aria-modal="true" aria-labelledby="choice-title" aria-describedby="choice-message">
        <div class="modal-icon">${icons.file}</div>
        <h2 id="choice-title"></h2>
        <p id="choice-message"></p>
        <p class="modal-note" id="choice-note" hidden></p>
        <div class="modal-actions" id="choice-actions"></div>
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
const choiceModal = $("choice-modal");

let handle: EditorHandle | null = null;
/** The typed Simple layer over the engine (document-change events, undo, review seeding…). */
let bridge: EngineBridge | null = null;
/** Word/Docs editing commands (shortcuts, command search, in-app engine prompts). */
let commands: EditorCommands | null = null;
/** Spelling as you type (Review > Spelling). */
let spelling: Spelling | null = null;
/** AutoFormat as you type (Review > AutoFormat), on the engine's insert-text hook. */
let autoFormat: AutoFormatInstallation | null = null;
const AUTOFORMAT_KEY = "simple-docs:autoformat";
let autoFormatEnabled = (() => {
  try {
    return window.localStorage.getItem(AUTOFORMAT_KEY) !== "off";
  } catch {
    return true;
  }
})();
/** The Home tab's font box (every installed font, type-to-search, recent fonts). */
let fontPicker: FontPicker | null = null;
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
type WordDocument = ReturnType<EditorHandle['getDocument']>;
type DocumentSnapshot = { revision: number; model: WordDocument; review: string };
let originalDocument: DocumentSnapshot & { data: Uint8Array; sourceData?: Uint8Array } | null = null;
let cachedDocumentPdf: DocumentSnapshot & { data: Uint8Array } | null = null;
let sourceLayout: DocumentSnapshot & { pdf: Promise<Uint8Array>; url?: string } | null = null;
const sessionId = crypto.randomUUID();

// Unsaved-work tracking. Dirty state comes from the model: every committed
// change (SIMPLE_HOOKS "simple:docchange") and every review change is compared
// with what was last saved, so no edit path can bypass it and view-only
// actions never count. `model: null` means content that was never saved.
let savedState: { model: WordDocument | null; review: string } | null = null;
let lastSeenModel: WordDocument | null = null;
let lastSeenReview = "";
/** A model found equal in content to the saved one (for example after undoing every edit). */
let contentCleanModel: WordDocument | null = null;
let contentCheckTimer: number | null = null;
const contentKeys = new WeakMap<object, string>();
let loadingDocument = false;
let documentGeneration = 0;
/** Review items the user chose to leave out of the saved file. */
let reviewNotSaved: ReviewSummary | null = null;
/** Where the open document came from: a writable file, a converted original Simple never overwrites, or recovered work. */
let documentOrigin: "new" | "file" | "converted" | "recovered" = "new";
/** The original file on disk (converted originals and recovered copies keep it for dialogs and messages). */
let documentSourcePath: string | null = null;
let fidelityItems: FidelityItem[] = [];
let fidelityScan: Promise<FidelityItem[]> = Promise.resolve([]);
let restoredRecoveryId: string | null = null;
let ownRecoveryExists = false;
let opening = false;
let saveFlowActive = false;
let printReviewNote = "";
const documentNotice = { conversion: "", fidelity: "", recovery: "", tone: "info" as "info" | "warning" };

type SaveResult = SavedFile & { warnings?: string[]; originalPath?: string };
type SiblingPlan = { sourcePath: string; path: string; name: string; officeEngine: boolean };
// Main-process handlers added for converted originals; declared here so this
// file does not depend on when src/types.d.ts gains them.
type SaveBesideSourceBridge = {
  planSaveBesideSource?: () => Promise<SiblingPlan | null>;
  saveBesideSource?: (input: { data: Uint8Array; sourcePath?: string | null }) => Promise<SaveResult>;
};
// The Windows user name for comment authors, once main exposes it (feature-detected).
type UserNameBridge = { getUserName?: () => Promise<string | null | undefined> | string | null | undefined };
const docsBridge = window.simpleDocs as SimpleDocsBridge & SaveBesideSourceBridge & UserNameBridge;
const RECOVERY_REVIEW_PREFIX = "simple-docs:recovery-review:";

const RECOVERY_IDLE_DELAY_MS = 12_000;
const RECOVERY_MAX_DELAY_MS = 60_000;
const RECOVERY_RETRY_DELAY_MS = 2_000;

function localSettings(): Storage | null {
  try {
    return window.localStorage;
  } catch {
    return null;
  }
}

// The author of new comments and suggestions. WordCanvas reads this object each
// time it attributes one, so a new name applies from the next comment on.
const storedAuthorName = readStoredAuthorName(localSettings());
let authorChosen = Boolean(storedAuthorName);
const author = authorIdentity(resolveAuthorName(storedAuthorName, null).name);

// Fonts the editor loads now: the curated set, plus the installed fonts used recently
// and in recently opened documents (main enumerates them locally; see src/ui/font-list.ts).
type FontBridge = {
  getDocumentFonts(options?: { families?: string[] }): Promise<CustomFontDef[]>;
  getFontFamilies?: () => Promise<Array<{ family: string; styles: number }>>;
};
const fontBridge = window.simpleDocs as unknown as FontBridge;
const installedFonts: CustomFontDef[] = await fontBridge.getDocumentFonts({ families: startupFontRequests(localSettings()) }).catch(() => []);
const editor = new WordCanvas({
  container: editorHost,
  fonts: { fonts: installedFonts, disableBuiltin: installedFonts.map((font) => font.family) },
  mode: "edit",
  user: author,
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
    // Insert > Header & footer (Word's Header, Footer and Page Number; Docs' Insert menu).
    api.addGroup("insert", { id: "simple.header-footer", label: "Header & footer", before: "insert.symbols" });
    const editableNow = () => Boolean(handle && handle.getMode() !== "view");
    api.addButton("simple.header-footer", { id: "simple.header-footer.header", icon: `${icons.header}<span class="big-cap">Header</span>`, tooltip: "Header — edit or remove the text at the top of every page", onClick: () => openBandMenu("header"), enabled: editableNow });
    api.addButton("simple.header-footer", { id: "simple.header-footer.footer", icon: `${icons.footer}<span class="big-cap">Footer</span>`, tooltip: "Footer — edit or remove the text at the bottom of every page", onClick: () => openBandMenu("footer"), enabled: editableNow });
    api.addButton("simple.header-footer", { id: "simple.header-footer.page-number", icon: `${icons.pageNumber}<span class="big-cap">Page number</span>`, tooltip: "Page numbers — position, format, starting number and first page", onClick: () => void openPageNumbers(), enabled: editableNow });
    // Home > Paragraph: numeric indents and spacing (Word's Paragraph dialog).
    api.addButton("home.paragraph", { id: "simple.paragraph.settings", icon: icons.paragraph, tooltip: "Indents and spacing — paragraph indents, space before and after, line spacing", onClick: () => void openParagraphSettings(), enabled: editableNow });
    // Review > Proofing: the two as-you-type switches, remembered on this computer.
    api.addTab({ id: "review", label: "Review", before: "view" });
    api.addGroup("review", { id: "simple.proofing", label: "Proofing" });
    api.addButton("simple.proofing", {
      id: "simple.proofing.spelling",
      icon: `${icons.spelling}<span class="big-cap">Spelling</span>`,
      tooltip: "Check spelling as you type",
      onClick: () => toggleSpelling(),
      // Called on every caret and formatting change: also keeps the font box current.
      active: (format) => {
        onFormatChange(format);
        return Boolean(spelling?.enabled && spelling.available);
      },
    });
    api.addButton("simple.proofing", {
      id: "simple.proofing.autoformat",
      icon: `${icons.autoformat}<span class="big-cap">AutoFormat</span>`,
      tooltip: "AutoFormat as you type: smart quotes, dashes, symbols, lists and links",
      onClick: () => toggleAutoFormat(),
      active: () => autoFormatEnabled && autoFormat?.active !== false,
    });
  },
  onLoadProgress(progress) {
    progressBar.style.width = `${Math.max(4, progress.percent * 100)}%`;
    if (progress.phase === "fonts") loaderDetail.textContent = "Loading document fonts…";
    if (progress.phase === "ready") loaderDetail.textContent = "Ready";
  },
});

function fileNameWithoutExtension(name: string) {
  return name.replace(/\.(?:docx|docm|dotx|dotm|doc|rtf|odt|html?|md|markdown|txt)$/i, "") || "Untitled document";
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

function savedStateLabel() {
  if (reviewNotSaved) return reviewNotSaved.suggestions && reviewNotSaved.comments ? "Saved without review items" : reviewNotSaved.comments ? "Saved without comments" : "Saved without suggestions";
  return currentPath ? "Saved locally" : "Not saved yet";
}

function setTitle() {
  titleLabel.textContent = documentName;
  saveState.textContent = saving ? "Saving…" : dirty ? "Unsaved" : documentOpen ? savedStateLabel() : "Ready";
  saveState.title = !saving && !dirty && documentOpen && reviewNotSaved
    ? `The saved file doesn't include ${describeReviewItems(reviewNotSaved)}. Simple Docs can't save comments and suggestions in this version yet.`
    : "";
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
  toastTimer = window.setTimeout(() => { toast.hidden = true; }, notificationDuration(message, tone));
}

/** The document's editing view is on screen (not the welcome screen or the Page view). */
function editorShown() {
  return Boolean(handle && documentOpen && welcome.hidden && !editorHost.hidden);
}

/**
 * Puts keyboard focus in the document, so typing right after Blank document,
 * Open or a closed dialog lands in it. A document that has no caret yet gets
 * one at its start (WordCanvas loads documents without a selection).
 */
function focusEditor() {
  if (!bridge || !handle || !editorShown() || activeModal(document)) return false;
  if (!bridge.getSelection()) {
    const first = firstParagraphId(handle.getDocument());
    if (first) bridge.setSelection({ anchor: { blockId: first, offset: 0 }, focus: { blockId: first, offset: 0 } });
  }
  return bridge.focus();
}

/**
 * Returns focus after a dialog: to a title-bar control that is still on screen
 * when the person came from one, otherwise to the document (also when they
 * came from the document, from nowhere, or from a control that has gone).
 */
function restoreFocus(element: HTMLElement | null) {
  if (activeModal(document)) return;
  const control = element && element !== document.body && element.isConnected && !editorHost.contains(element) && element.getClientRects().length > 0;
  if (control) element.focus();
  else focusEditor();
}

function setAuthorName(name: string, chosen: boolean) {
  const identity = authorIdentity(name);
  author.firstName = identity.firstName;
  author.lastName = identity.lastName;
  if (chosen) {
    authorChosen = true;
    storeAuthorName(localSettings(), authorDisplayName(author));
  }
  $("author-name-label").textContent = authorDisplayName(author);
}

async function editAuthorName() {
  toggleMenu(false);
  const returnFocus = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  const result = await openFormDialog({
    title: "Your name",
    icon: icons.person,
    message: "Simple Docs shows this name on the comments and suggestions you add. It stays on this computer.",
    fields: [{ id: "name", label: "Name", value: authorDisplayName(author), maxLength: 80, placeholder: "Your name" }],
    submitLabel: "Save",
    validate: (values) => cleanAuthorName(values.name) ? null : { field: "name", message: "Type the name to show on your comments." },
    onClosed: () => window.setTimeout(() => restoreFocus(returnFocus), 0),
  });
  if (!result) return;
  setAuthorName(result.values.name, true);
  notify(`New comments and suggestions show “${authorDisplayName(author)}”.`);
}

/**
 * Uses the Windows user name until the person chooses a name: from main when
 * it reports one (getUserName, feature-detected), else from the user profile
 * the app runs in.
 */
async function adoptWindowsUserName() {
  if (authorChosen) return;
  let windowsUser: string | null | undefined = null;
  if (typeof docsBridge.getUserName === "function") {
    try {
      windowsUser = await docsBridge.getUserName();
    } catch (error) {
      console.warn("The Windows user name could not be read", error);
    }
  }
  windowsUser ||= windowsUserFromAppUrl(window.location.href);
  if (windowsUser && !authorChosen) setAuthorName(resolveAuthorName(null, windowsUser).name, false);
}

let suggestingNameShown = false;
/** The first switch to Suggesting says whose name the suggestions carry (once, until a name is chosen). */
function noteSuggestingAuthor(mode: EditMode) {
  if (mode !== "suggest" || authorChosen || suggestingNameShown) return;
  suggestingNameShown = true;
  notify(`Your suggestions show “${authorDisplayName(author)}”. To use another name, search for “Author name” (Alt+Q).`);
}

/** Undo/Redo in the title bar follow the editor history (disabled with nothing to undo, in Viewing and in Page view). */
function refreshHistoryButtons(state?: { canUndo: boolean; canRedo: boolean }) {
  const available = editorShown() && Boolean(bridge);
  const canUndo = available && (state ? state.canUndo : bridge!.canUndo());
  const canRedo = available && (state ? state.canRedo : bridge!.canRedo());
  $<HTMLButtonElement>("undo-button").disabled = !canUndo;
  $<HTMLButtonElement>("redo-button").disabled = !canRedo;
}

/** Closes WordCanvas's find bar, menus and popovers (they live on <body>, outside the editor host). */
function closeEditorOverlays() {
  const find = [...document.querySelectorAll<HTMLElement>(".cw-float-panel")].find((panel) => panel.style.display !== "none" && panel.querySelector("input"));
  [...(find?.querySelectorAll<HTMLButtonElement>("button") ?? [])].find((button) => button.title.startsWith("Close"))?.click();
  // A press outside closes menus and popovers through the engine's own handlers,
  // which also remove their window listeners.
  if (document.querySelector(".cw-pop, .cw-menu")) document.body.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
  document.querySelectorAll(".cw-pop, .cw-menu").forEach((element) => element.remove());
  if (document.activeElement instanceof HTMLElement && (editorHost.contains(document.activeElement) || document.activeElement.closest(".cw-float-panel, .cw-float-drawer"))) document.activeElement.blur();
}

function capitalize(text: string) {
  return text ? text[0].toUpperCase() + text.slice(1) : text;
}

function sentence(text: string) {
  const trimmed = text.trim();
  return !trimmed || /[.!?…]$/.test(trimmed) ? trimmed : `${trimmed}.`;
}

function refreshDocumentNotice() {
  const notice = $("document-compatibility");
  const text = [documentNotice.recovery, documentNotice.conversion, documentNotice.fidelity].filter(Boolean).join(" ");
  $("document-notice-text").textContent = text;
  notice.classList.toggle("is-warning", documentNotice.tone === "warning");
  notice.hidden = !text;
}

function clearDocumentNotice() {
  documentNotice.conversion = "";
  documentNotice.fidelity = "";
  documentNotice.recovery = "";
  documentNotice.tone = "info";
  refreshDocumentNotice();
}

type Choice = { id: string; label: string; primary?: boolean };
let choiceResolve: ((id: string) => void) | null = null;
let choiceCancelId = "cancel";

/** A small in-app question (never a native dialog); resolves the chosen id, or the cancel id on Escape. */
function askChoice(options: { title: string; message: string; note?: string; choices: Choice[]; cancelId?: string }): Promise<string> {
  choiceResolve?.(choiceCancelId);
  const returnFocus = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  $("choice-title").textContent = options.title;
  $("choice-message").textContent = options.message;
  const note = $("choice-note");
  note.textContent = options.note ?? "";
  note.hidden = !options.note;
  const actions = $("choice-actions");
  actions.replaceChildren();
  return new Promise((resolve) => {
    const finish = (id: string) => {
      if (choiceResolve !== finish) return;
      choiceResolve = null;
      choiceModal.hidden = true;
      resolve(id);
      window.setTimeout(() => restoreFocus(returnFocus), 0);
    };
    choiceResolve = finish;
    choiceCancelId = options.cancelId ?? "cancel";
    let primary: HTMLButtonElement | null = null;
    for (const choice of options.choices) {
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = choice.label;
      button.dataset.choice = choice.id;
      if (choice.primary) {
        button.className = "modal-primary";
        primary = button;
      }
      button.addEventListener("click", () => finish(choice.id));
      actions.append(button);
    }
    choiceModal.hidden = false;
    (primary ?? actions.querySelector<HTMLButtonElement>("button"))?.focus();
  });
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

function contentKeyOf(model: WordDocument) {
  let key = contentKeys.get(model);
  if (key === undefined) {
    key = documentContentKey(model);
    contentKeys.set(model, key);
  }
  return key;
}

/** Same document content (identity first; undo creates equal content in new objects). */
function sameContent(left: WordDocument | null | undefined, right: WordDocument | null | undefined) {
  if (!left || !right) return false;
  if (left === right) return true;
  try {
    return contentKeyOf(left) === contentKeyOf(right);
  } catch (error) {
    console.warn("Document comparison failed", error);
    return false;
  }
}

function currentReviewKey() {
  return handle ? reviewContentKey(handle.getReview()) : "";
}

function hasUnsavedChanges(model = handle?.getDocument(), review = currentReviewKey()) {
  if (!handle || !documentOpen || !savedState) return false;
  if (savedState.model === null) return true;
  if (review !== savedState.review) return true;
  return model !== savedState.model && model !== contentCleanModel;
}

/**
 * Re-derives the unsaved state from the model. Called for every committed
 * document change and review change (and before Save and Close), so typing,
 * ribbon and mini-toolbar commands, Find & Replace, context-menu commands,
 * ruler, table and picture drags, keyboard shortcuts and dialog Apply buttons
 * all count, while selection, zoom, panes and other view changes never do.
 */
function syncDirty(origin = "unknown") {
  if (!handle || !documentOpen || loadingDocument || !savedState) return;
  const model = handle.getDocument();
  const review = currentReviewKey();
  const changed = model !== lastSeenModel || review !== lastSeenReview;
  if (changed) {
    lastSeenModel = model;
    lastSeenReview = review;
    revision += 1;
    cachedDocumentPdf = null;
    if (sourceLayout) $("document-layout-note").textContent = matchesDocumentSnapshot(sourceLayout)
      ? "Page view uses local Office rendering. Complex formatting can differ from Word or Google Docs."
      : "Edits are in Edit document. Printing and PDF export include your current edits.";
  }
  const unsaved = hasUnsavedChanges(model, review);
  // Undo and redo can return to the saved content in a new model object.
  if (unsaved && savedState.model && (origin === "undo" || origin === "redo" || origin === "unknown")) scheduleContentCheck();
  if (unsaved !== dirty) {
    dirty = unsaved;
    if (!dirty) {
      clearRecoveryTimers();
      void dropOwnRecovery();
    }
    setTitle();
  }
  if (dirty && changed) scheduleRecovery();
}

function scheduleContentCheck() {
  if (contentCheckTimer !== null) window.clearTimeout(contentCheckTimer);
  contentCheckTimer = window.setTimeout(() => {
    contentCheckTimer = null;
    if (!handle || !documentOpen || !savedState?.model) return;
    const model = handle.getDocument();
    if (model === savedState.model || model === contentCleanModel) return;
    if (sameContent(model, savedState.model)) {
      contentCleanModel = model;
      syncDirty("content");
    }
  }, 250);
}

/** Starts tracking a freshly loaded or created model; `saved` is false for content that has never been saved. */
function adoptDocumentModel(saved: boolean) {
  if (!handle) return null;
  const model = handle.getDocument();
  const review = currentReviewKey();
  savedState = { model: saved ? model : null, review };
  lastSeenModel = model;
  lastSeenReview = review;
  contentCleanModel = null;
  reviewNotSaved = null;
  if (contentCheckTimer !== null) window.clearTimeout(contentCheckTimer);
  contentCheckTimer = null;
  revision += 1;
  cachedDocumentPdf = null;
  dirty = !saved;
  documentGeneration += 1;
  return model;
}

/** Forgets the open-time content scan: after a save the file holds exactly what the editor keeps. */
function resetFidelity() {
  fidelityItems = [];
  fidelityScan = Promise.resolve([]);
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
  const review = reviewSummary(handle.getReview());
  let snapshotSaved = false;
  try {
    const blob = await handle.exportDocx();
    await window.simpleDocs.saveRecovery({
      id: sessionId,
      data: new Uint8Array(await blob.arrayBuffer()),
      title: documentName,
      sourcePath: currentPath ?? documentSourcePath,
    });
    snapshotSaved = true;
    ownRecoveryExists = true;
    // The snapshot cannot hold review items; remember that they existed so a
    // restore can say so instead of dropping them silently.
    rememberRecoveryReview(sessionId, review);
    // Restored work now lives in this session's snapshot; retire the old entry.
    if (restoredRecoveryId) {
      const adopted = restoredRecoveryId;
      restoredRecoveryId = null;
      await forgetRecovery(adopted);
    }
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
      if (snapshotSaved) await dropOwnRecovery();
    } else if (!snapshotSaved || revision !== checkpointRevision) {
      scheduleRecovery(snapshotSaved ? RECOVERY_IDLE_DELAY_MS : RECOVERY_RETRY_DELAY_MS);
    }
  }
}

function rememberRecoveryReview(id: string, summary: ReviewSummary) {
  try {
    if (summary.total) window.localStorage.setItem(`${RECOVERY_REVIEW_PREFIX}${id}`, JSON.stringify({ suggestions: summary.suggestions, comments: summary.comments }));
    else window.localStorage.removeItem(`${RECOVERY_REVIEW_PREFIX}${id}`);
  } catch {
    // Storage is a courtesy note only; recovery itself does not depend on it.
  }
}

function recoveredReviewItems(id: string): ReviewSummary | null {
  try {
    const value = JSON.parse(window.localStorage.getItem(`${RECOVERY_REVIEW_PREFIX}${id}`) ?? "null");
    const count = (raw: unknown) => Math.max(0, Math.floor(Number(raw) || 0));
    const suggestions = count(value?.suggestions);
    const comments = count(value?.comments);
    return suggestions + comments ? { suggestions, comments, total: suggestions + comments } : null;
  } catch {
    return null;
  }
}

async function forgetRecovery(id: string) {
  try {
    await window.simpleDocs.clearRecovery(id);
    rememberRecoveryReview(id, reviewSummary(null));
  } catch (error) {
    console.warn("Recovery entry could not be cleared", error);
  }
}

/** Removes this window's snapshot once nothing unsaved is left (saved, undone or discarded). */
async function dropOwnRecovery() {
  if (!ownRecoveryExists || recoveryInFlight) return;
  ownRecoveryExists = false;
  await forgetRecovery(sessionId);
}

async function forgetAllRecoveries() {
  clearRecoveryTimers();
  // Let a snapshot that is being written finish first, so it cannot reappear.
  for (let waited = 0; recoveryInFlight && waited < 15_000; waited += 50) await new Promise((resolve) => window.setTimeout(resolve, 50));
  clearRecoveryTimers();
  ownRecoveryExists = false;
  await forgetRecovery(sessionId);
  if (restoredRecoveryId) {
    const restored = restoredRecoveryId;
    restoredRecoveryId = null;
    await forgetRecovery(restored);
  }
}

/**
 * What main returns for a file (electron/main.cjs loadDocumentBytes); declared here so
 * this file does not depend on when src/types.d.ts gains the import fields.
 */
type OpenPayload = Omit<DocumentPayload, "convertedFrom"> & {
  convertedFrom?: string;
  /** RTF, OpenDocument, web page, Markdown or text: `data` is the file itself, for the native importers. */
  importFormat?: ImportFormat;
  /** A Word template (.dotx/.dotm): opens as a new untitled document. */
  template?: boolean;
  /** The package carries macros, which are never run or kept. */
  macros?: boolean;
};

// Originals Simple never writes back in place: edits save as a .docx beside them.
const CONVERTED_ORIGIN_KINDS = new Set(["rtf", "html", "mht", "word2003", "docx-renamed", "docm", "odt", "md", "txt"]);
const IMPORT_LABELS: Record<ImportFormat, string> = {
  txt: "Plain text",
  md: "Markdown",
  html: "Web page",
  rtf: "Rich Text",
  odt: "OpenDocument text",
};

function conversionNotice(payload: OpenPayload, textOnlyImport: boolean, importWarnings: readonly string[] = []) {
  const warnings = [...(payload.conversionWarnings ?? []), ...importWarnings].map((warning) => sentence(String(warning))).filter(Boolean);
  if (payload.importFormat) return [`${IMPORT_LABELS[payload.importFormat]} · Saving creates a .docx file; “${payload.name}” stays unchanged.`, ...warnings].join(" ");
  if (payload.template) return `New document from the template “${payload.name}”.${payload.macros ? " Its macros aren’t run or kept." : ""} Saving creates a new .docx file; the template stays unchanged.`;
  if (payload.convertedFrom === "docm") return `Macro-enabled document · Macros aren’t run or kept. Saving creates a .docx file; “${payload.name}” stays unchanged.`;
  if (warnings.length) return warnings.join(" ");
  if (textOnlyImport) return "Text-only import · This .doc file’s formatting, tables and images could not be read on this computer, so only its text was recovered. Saving creates a separate .docx file; the original stays unchanged.";
  if (payload.convertedFrom === "doc") return "Legacy Word document · Save keeps the .doc format. Save as or Export As creates a DOCX copy. Complex layouts may change after editing.";
  if (payload.macros) return "This document contains macros. Simple Docs doesn’t run them, and saving leaves them out.";
  return "";
}

function fidelityNotice(items: FidelityItem[]) {
  if (!items.length) return "";
  const list = describeFidelityItems(items);
  return currentPath
    ? `Simple Docs can't keep some content in this file: ${list}. Saving offers a copy, so the original stays intact.`
    : `Simple Docs can't keep some content in this file: ${list}. It won't be in the saved copy.`;
}

/** Opens a payload in this window. Throws, with nothing changed, when the editor cannot read it. */
async function loadPayload(payload: OpenPayload, options: { recoveryId?: string } = {}) {
  if (!handle) throw new Error("Simple Docs is still starting. Try again in a moment.");
  const recovered = Boolean(options.recoveryId);
  const importFormat = recovered ? undefined : payload.importFormat;
  const template = !recovered && payload.template === true;
  showLoading(recovered ? "Recovering document" : `Opening ${payload.name}`, "Reading document structure…");
  hidePasteOptions();
  const previousModel = handle.getDocument();
  let importWarnings: string[] = [];
  loadingDocument = true;
  try {
    if (importFormat) {
      // RTF, OpenDocument, web pages, Markdown and text: read here, without LibreOffice
      // or the network, into a new document (the original file is never written).
      const imported = await importDocument(payload.data, { format: importFormat, fileName: payload.name });
      importWarnings = imported.warnings;
      handle.setDocument(imported.document);
    } else {
      await handle.openDocx(toArrayBuffer(payload.data));
    }
  } catch (error) {
    throw new Error(`Simple Docs couldn't read “${payload.name}”. ${sentence(cleanErrorMessage(error, "Its contents are damaged or use a layout the editor doesn't support."))}`);
  } finally {
    loadingDocument = false;
  }
  // An unpatched engine reports import failures in an alert and keeps the old
  // document; never bind the file to a model that is not its content.
  if (handle.getDocument() === previousModel) throw new Error(`Simple Docs couldn't read “${payload.name}”. Its contents are damaged or use a layout the editor doesn't support.`);

  const importedLegacyDocument = payload.convertedFrom === "doc";
  const textOnlyImport = importedLegacyDocument && payload.conversionMethod !== "layout";
  const convertedOriginal = textOnlyImport || CONVERTED_ORIGIN_KINDS.has(String(payload.convertedFrom ?? ""));
  const requiresSaveAs = recovered || convertedOriginal || template || payload.requiresSaveAs === true;
  currentPath = requiresSaveAs ? null : payload.path;
  currentFormat = !recovered && payload.format === "doc" ? "doc" : "docx";
  currentSourceHash = recovered ? undefined : payload.sourceHash;
  // A template is only a starting point: nothing is saved next to it or over it.
  documentSourcePath = template ? null : payload.sourcePath ?? payload.path ?? null;
  documentOrigin = recovered ? "recovered" : convertedOriginal ? "converted" : "file";
  documentName = recovered ? recoveredDocumentName(payload.name) : template ? "Untitled document" : fileNameWithoutExtension(payload.name);
  documentOpen = true;
  restoredRecoveryId = options.recoveryId ?? null;
  // Recovered work has never been saved anywhere, so it starts unsaved.
  const model = adoptDocumentModel(!recovered)!;
  const generation = documentGeneration;
  // Imported files, templates and macro-enabled documents are never written back as
  // they came: their first save exports a fresh .docx from the editor.
  const freshExport = Boolean(importFormat) || template || payload.convertedFrom === "docm";
  originalDocument = freshExport ? null : { revision, model, review: savedState!.review, data: new Uint8Array(payload.data), sourceData: payload.sourceData };
  welcome.hidden = true;
  editorHost.classList.add("is-active");

  resetFidelity();
  documentNotice.conversion = recovered ? "" : conversionNotice(payload, textOnlyImport, importWarnings);
  documentNotice.fidelity = "";
  documentNotice.recovery = "";
  documentNotice.tone = textOnlyImport || importWarnings.length ? "warning" : "info";
  if (recovered) {
    const lostReview = recoveredReviewItems(options.recoveryId!);
    documentNotice.recovery = `Recovered work from an interrupted session. It opens as a new copy, so ${documentSourcePath ? "the original file stays as it is" : "nothing is overwritten"}; save it to keep it.${lostReview ? ` ${capitalize(describeReviewItems(lostReview))} from that session couldn't be recovered.` : ""}`;
    documentNotice.tone = lostReview ? "warning" : "info";
  } else if (!importFormat) {
    // (An imported file has no package to scan: its import reported what it left out.)
    // Content the editor drops (comments, tracked changes, charts…) is listed
    // once, and the first save over the original offers a copy instead.
    fidelityScan = scanDocxFidelity(payload.data).then((report) => report.items).catch((error) => {
      console.warn("Document content scan failed", error);
      return [];
    });
    void fidelityScan.then((items) => {
      if (generation !== documentGeneration || !items.length) return;
      fidelityItems = items;
      documentNotice.fidelity = fidelityNotice(items);
      documentNotice.tone = "warning";
      refreshDocumentNotice();
    });
  }
  refreshDocumentNotice();

  resetOriginalLayout();
  if (payload.originalLayout) {
    const layout = { ...documentSnapshot(), pdf: window.simpleDocs.getOriginalLayoutPdf(payload.originalLayout), url: undefined as string | undefined };
    sourceLayout = layout;
    $("document-layout-bar").hidden = false;
    setOriginalLayoutView(true);
    void layout.pdf.then((data) => {
      if (sourceLayout !== layout) return;
      layout.url = URL.createObjectURL(new Blob([toArrayBuffer(data)], { type: "application/pdf" }));
      // Honors a zoom chosen while the page view was still being prepared.
      showOriginalLayoutPdf(layout.url);
      $("original-layout-loading").hidden = true;
    }).catch((error) => {
      if (sourceLayout !== layout) return;
      $("original-layout-loading").textContent = `The page view could not be prepared. ${error instanceof Error ? error.message : ""} Choose Edit document to continue.`;
    });
  }
  setTitle();
  // Recovered work is written to this session's own snapshot right away, then
  // the restored entry is retired.
  if (dirty) scheduleRecovery(recovered ? RECOVERY_RETRY_DELAY_MS : RECOVERY_IDLE_DELAY_MS);
  notify(recovered
    ? "Recovered work is ready. Save it to keep it."
    : template
      ? "New document from the template. Save it to keep it."
      : importedLegacyDocument
        ? textOnlyImport ? "Text-only copy opened. See the import notice above the document." : "Legacy .doc opened with formatted conversion."
        : payload.convertedFrom === "docx-renamed"
          ? "This file contains a Word document (.docx). Saving creates a .docx copy; the original stays unchanged."
          : convertedOriginal
            ? "Document opened. Saving creates a .docx file; the original stays unchanged."
            : "Document opened.");
  refreshHistoryButtons();
  annotateRibbonShortcuts();
  // Typing right after Open goes into the document (the Page view has no caret).
  focusEditor();
}

/**
 * Runs one open. Every failure is shown in plain words and leaves the window
 * as it was: no file path is bound and no 'Document opened' is reported.
 */
async function openWith(load: () => Promise<OpenPayload | null | undefined>, options: { recoveryId?: string; onError?: (error: unknown) => boolean } = {}) {
  if (opening) {
    notify("Another document is still opening. Try again when it is ready.", "error");
    return false;
  }
  opening = true;
  // Main may be converting a legacy file for a while; show progress if so.
  const loader = window.setTimeout(() => showLoading("Opening document", "Reading the file…"), 400);
  try {
    const payload = await load();
    window.clearTimeout(loader);
    if (!payload) return false;
    await loadPayload(payload, options);
    return true;
  } catch (error) {
    console.error(error);
    if (!options.onError?.(error)) notify(openErrorMessage(error), "error");
    return false;
  } finally {
    window.clearTimeout(loader);
    opening = false;
    if (!loadingOverlay.hidden) hideLoading();
  }
}

async function openInNewWindow(filePath?: string) {
  try {
    await window.simpleDocs.openInNewWindow(filePath);
    return true;
  } catch (error) {
    console.error(error);
    notify(openErrorMessage(error), "error");
    return false;
  }
}

function beginBlankDocument() {
  hidePasteOptions();
  resetOriginalLayout();
  clearDocumentNotice();
  resetFidelity();
  originalDocument = null;
  cachedDocumentPdf = null;
  documentOpen = true;
  currentPath = null;
  currentFormat = 'docx';
  currentSourceHash = undefined;
  documentSourcePath = null;
  documentOrigin = "new";
  restoredRecoveryId = null;
  documentName = "Untitled document";
  adoptDocumentModel(true);
  welcome.hidden = true;
  editorHost.classList.add("is-active");
  setTitle();
  refreshHistoryButtons();
  // "Start writing immediately": the caret is in the document, ready for typing.
  focusEditor();
}

async function chooseOpen() {
  if (documentOpen) {
    await openInNewWindow();
    return;
  }
  await openWith(() => window.simpleDocs.openFile());
}

/** The review guard: the file cannot hold comments or suggestions yet, so say so before saving. */
async function confirmReviewLoss(action: "save" | "export", summary: ReviewSummary, exportLabel = ""): Promise<"accept" | "without" | "cancel"> {
  const items = describeReviewItems(summary);
  const them = summary.total === 1 ? "it" : "them";
  const choices: Choice[] = [{ id: "cancel", label: "Cancel" }];
  const structuredExport = action === "export" && exportLabel !== "";
  const withoutLabel = action === "export" ? "Export without suggestions" : summary.suggestions && summary.comments ? "Save without them" : summary.comments ? "Save without comments" : "Save without suggestions";
  if (!structuredExport) choices.push({ id: "without", label: withoutLabel, primary: !summary.suggestions });
  if (summary.suggestions) choices.push({ id: "accept", label: action === "save" ? "Accept suggestions and save" : "Accept suggestions and export", primary: true });
  const comments = summary.comments ? `; ${summary.comments === 1 ? "the comment stays" : "comments stay"} here only until you close this window` : "";
  const message = action === "save"
    ? `This document has ${items}. The saved file won't include ${them}${comments}.`
    : structuredExport
      ? `This document has ${items}. A ${exportLabel} export can't show pending suggestions, so accept them first or cancel.`
      : `This document has ${items}. The exported file won't include ${them}.`;
  const choice = await askChoice({
    title: action === "save" ? "Comments and suggestions can't be saved in this version yet" : "Suggestions can't be exported yet",
    message,
    note: summary.suggestions ? "Accepting keeps the suggested text in the document; you can still undo it with Ctrl+Z." : undefined,
    choices,
  });
  return choice === "accept" || choice === "without" ? choice : "cancel";
}

/** Converted originals (.doc without a local engine, RTF/HTML saved as .doc): edits go to a .docx beside them. */
async function planBesideSource(): Promise<SiblingPlan | null> {
  if (!docsBridge.planSaveBesideSource || !docsBridge.saveBesideSource) return null;
  const unchangedDoc = currentFormat === "doc" && Boolean(originalDocument?.sourceData) && sameContent(handle?.getDocument(), originalDocument?.model);
  if (documentOrigin !== "converted" && !(currentFormat === "doc" && currentPath && !unchangedDoc)) return null;
  try {
    const plan = await docsBridge.planSaveBesideSource();
    if (!plan) return null;
    // A formatted .doc saves in place while a local Office engine exists.
    return documentOrigin === "converted" || !plan.officeEngine ? plan : null;
  } catch (error) {
    console.warn("Could not plan a save beside the original", error);
    return null;
  }
}

function fileName(filePath: string | null | undefined) {
  return String(filePath ?? "").split(/[\\/]/).pop() || "";
}

async function saveDocument(forceDialog = false): Promise<boolean> {
  if (!handle || !documentOpen) return false;
  if (saving || saveFlowActive) {
    notify(saving ? "Wait for the current save, export or print preparation to finish, then save again." : "Finish the open save question first.", "error");
    return false;
  }
  saveFlowActive = true;
  try {
    syncDirty();
    // 1. Review items cannot be stored in the file yet.
    let droppedReview: ReviewSummary | null = null;
    const review = reviewSummary(handle.getReview());
    if (review.total) {
      const choice = await confirmReviewLoss("save", review);
      if (choice === "cancel") return false;
      if (choice === "accept") {
        handle.acceptAllSuggestions();
        syncDirty("command");
      }
      const left = reviewSummary(handle.getReview());
      droppedReview = left.total ? left : null;
    }

    // 2. Where the edits go.
    let mode: "in-place" | "dialog" | "beside" = forceDialog || !currentPath ? "dialog" : "in-place";
    let dialogName = documentName || "Untitled document";
    let besideSource: string | null = null;
    let savingCopy = false;
    let replacingOriginal = false;
    let protectOriginal = Boolean(sourceLayout && !matchesDocumentSnapshot(sourceLayout));
    const unchanged = Boolean(originalDocument && sameContent(handle.getDocument(), originalDocument.model));
    if (!forceDialog) {
      const plan = await planBesideSource();
      if (plan) {
        const original = fileName(plan.sourcePath);
        const choice = await askChoice({
          title: `Save as “${plan.name}”?`,
          message: `Simple Docs can't write edits into “${original}”, so they go to a new file next to it. The original stays unchanged.`,
          choices: [{ id: "cancel", label: "Cancel" }, { id: "choose", label: "Choose location…" }, { id: "save", label: "Save", primary: true }],
        });
        if (choice === "cancel") return false;
        if (choice === "choose") mode = "dialog";
        else {
          mode = "beside";
          besideSource = plan.sourcePath;
        }
      }
    }
    if (mode === "in-place" && !unchanged) {
      // Until the first save over the original, content the editor drops
      // makes "Save a copy" the default.
      const items = await fidelityScan;
      if (items.length) {
        const original = fileName(currentPath);
        const choice = await askChoice({
          title: "Keep the original file?",
          message: `“${original}” has content Simple Docs can't keep: ${describeFidelityItems(items)}. Saving over it would remove that content.`,
          note: `Save a copy keeps the original intact. Replace original keeps a backup of it next to the file.`,
          choices: [{ id: "cancel", label: "Cancel" }, { id: "replace", label: "Replace original" }, { id: "copy", label: "Save a copy", primary: true }],
        });
        if (choice === "cancel") return false;
        if (choice === "copy") {
          mode = "dialog";
          savingCopy = true;
          dialogName = `${documentName} (edited)`;
        } else {
          replacingOriginal = true;
          protectOriginal = true;
        }
      }
    }

    // 3. Write.
    saving = true;
    setTitle();
    const savingRevision = revision;
    const savingModel = handle.getDocument();
    const savingReview = currentReviewKey();
    const previousPath = currentPath;
    try {
      const data = await currentDocxBytes();
      let result: SaveResult | null;
      if (mode === "beside") {
        result = await docsBridge.saveBesideSource!({ data, sourcePath: besideSource });
      } else {
        result = await window.simpleDocs.saveDocx({
          data,
          path: currentPath,
          name: `${dialogName}.docx`,
          forceDialog: mode === "dialog" && currentPath !== null,
          format: mode === "in-place" ? currentFormat : "docx",
          sourceData: mode === "in-place" && unchanged ? originalDocument?.sourceData : undefined,
          expectedHash: currentSourceHash,
          protectOriginal: mode === "in-place" && protectOriginal,
        }) as SaveResult | null;
      }
      if (!result) return false;
      currentPath = result.path;
      currentFormat = result.format === 'doc' ? 'doc' : 'docx';
      currentSourceHash = result.sourceHash;
      documentName = fileNameWithoutExtension(result.name);
      documentSourcePath = result.path;
      documentOrigin = "file";
      // A fresh export holds exactly what the editor keeps, so the content
      // warning is over; the original bytes (nothing edited) still carry it.
      if (!unchanged) {
        resetFidelity();
        documentNotice.fidelity = "";
      }
      documentNotice.recovery = "";
      if (mode === "beside") {
        documentNotice.conversion = `Now editing “${result.name}”, saved next to the original “${fileName(result.originalPath ?? besideSource)}”, which is unchanged.`;
        documentNotice.tone = "info";
      } else if (documentNotice.conversion && currentFormat === "docx") {
        documentNotice.conversion = "";
      }
      refreshDocumentNotice();
      savedState = { model: savingModel, review: savingReview };
      contentCleanModel = null;
      reviewNotSaved = droppedReview;
      originalDocument = { revision: savingRevision, model: savingModel, review: savingReview, data, sourceData: result.sourceData };
      dirty = hasUnsavedChanges();
      if (!dirty) await forgetAllRecoveries();
      const messages: string[] = [];
      if (mode === "beside") messages.push(`Saved as “${result.name}” next to the original, which stays unchanged.`);
      else if (savingCopy) messages.push(`Saved a copy as “${result.name}”. The original “${fileName(previousPath)}” is unchanged.`);
      else if (replacingOriginal) messages.push(`Saved. A backup of the original is next to it as “${fileNameWithoutExtension(fileName(previousPath))}.before-simple-edit.docx”.`);
      else messages.push("Saved.");
      if (droppedReview) messages.push(`It doesn't include ${describeReviewItems(droppedReview)}.`);
      for (const warning of result.warnings ?? []) messages.push(sentence(String(warning)));
      if (dirty) {
        scheduleRecovery();
        messages.push("Newer changes remain unsaved.");
      }
      notify(messages.join(" "));
      void refreshWelcomeLists();
      return !dirty;
    } catch (error) {
      console.error(error);
      notify(cleanErrorMessage(error, "The document could not be saved."), "error");
      return false;
    } finally {
      saving = false;
      if (dirty) scheduleRecovery(RECOVERY_RETRY_DELAY_MS);
      setTitle();
    }
  } finally {
    saveFlowActive = false;
  }
}

type DocumentExportFormat = "docx" | "pdf" | StructuredDocumentExportFormat;

function documentSnapshot(): DocumentSnapshot {
  if (!handle) throw new Error("Open a document first.");
  return { revision, model: handle.getDocument(), review: reviewContentKey(handle.getReview()) };
}

function matchesDocumentSnapshot(snapshot: DocumentSnapshot | null): boolean {
  return Boolean(snapshot && handle && snapshot.model === handle.getDocument() && snapshot.review === reviewContentKey(handle.getReview()));
}

/** The original bytes while the content is unchanged (including after undoing every edit), else a fresh export. */
async function currentDocxBytes(): Promise<Uint8Array> {
  if (!handle) throw new Error("Open a document first.");
  if (originalDocument && sameContent(handle.getDocument(), originalDocument.model)) return originalDocument.data;
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
  document.body.classList.remove("page-view");
}

function setOriginalLayoutView(original: boolean) {
  // The editor is hidden in Page view, so its find bar, floating bars and panes
  // close too: nothing may edit the hidden document behind the page.
  hidePasteOptions();
  if (original) closeEditorOverlays();
  document.body.classList.toggle("page-view", original);
  $("original-layout-view").hidden = !original;
  editorHost.hidden = original;
  $("source-view-scale").hidden = !original;
  $("original-layout-button").setAttribute("aria-pressed", String(original));
  $("edit-layout-button").setAttribute("aria-pressed", String(!original));
  $("document-layout-note").textContent = original
    ? matchesDocumentSnapshot(sourceLayout) ? "Page view uses local Office rendering. Complex formatting can differ from Word or Google Docs." : "Source file for reference. Printing and PDF export include your current edits."
    : "Editing view · Floating objects and complex tables can reflow. Use Save as to keep a separate original copy.";
  refreshHistoryButtons();
  if (!original) focusEditor();
}

/** The PDF viewer fragment for the Page view zoom choice. */
function originalLayoutFragment(value: string) {
  return `${value.startsWith("Fit") ? `view=${value}` : `zoom=${value}`}&toolbar=0`;
}

/**
 * Shows the Page view PDF at the chosen zoom. Chromium's PDF viewer ignores a
 * fragment-only change of a loaded document, so the frame is replaced to load
 * it again (the new frame keeps the id).
 */
function showOriginalLayoutPdf(url: string) {
  const frame = $<HTMLIFrameElement>("original-layout-pdf");
  const next = frame.cloneNode(false) as HTMLIFrameElement;
  next.src = `${url}#${originalLayoutFragment($<HTMLSelectElement>("source-view-scale").value)}`;
  next.hidden = false;
  frame.replaceWith(next);
}

$("original-layout-button").addEventListener("click", () => setOriginalLayoutView(true));
$("edit-layout-button").addEventListener("click", () => setOriginalLayoutView(false));
$("source-view-scale").addEventListener("change", () => {
  // Before the page view is ready the choice is kept and applied when it arrives.
  if (sourceLayout?.url) showOriginalLayoutPdf(sourceLayout.url);
});

const EXPORT_LABELS: Record<DocumentExportFormat, string> = {
  docx: "Word document",
  pdf: "PDF",
  html: "web page",
  md: "Markdown",
  txt: "plain text",
};

async function exportDocument(format: DocumentExportFormat) {
  if (!handle || !documentOpen || saving || saveFlowActive) return;
  // Pending suggestions change the exported text; the export never drops them silently.
  const review = reviewSummary(handle.getReview());
  if (review.suggestions) {
    saveFlowActive = true;
    try {
      const structured = format !== "docx" && format !== "pdf";
      const choice = await confirmReviewLoss("export", review, structured ? EXPORT_LABELS[format] : "");
      if (choice === "cancel") return;
      if (choice === "accept") {
        handle.acceptAllSuggestions();
        syncDirty("command");
      }
    } finally {
      saveFlowActive = false;
    }
  }
  saving = true;
  setTitle();
  try {
    let data: Uint8Array;
    const warnings: string[] = [];
    if (format === "docx") {
      data = await currentDocxBytes();
    } else if (format === "pdf") {
      data = await currentPdfBytes();
    } else {
      const document = format === "txt" ? handle.getDocument() : await prepareDocumentForExport(handle.getDocument());
      const serialized = serializeDocumentWithWarnings(document, format, documentName);
      data = new TextEncoder().encode(serialized.text);
      warnings.push(...serialized.warnings);
    }
    const left = reviewSummary(handle.getReview());
    const result = await window.simpleDocs.saveExport({ data, name: documentName, format }) as SaveResult | null;
    if (result) {
      warnings.push(...(result.warnings ?? []));
      const notes = [left.total ? `It doesn't include ${describeReviewItems(left)}.` : "", ...warnings.map((warning) => sentence(String(warning)))].filter(Boolean);
      notify([`${EXPORT_LABELS[format]} exported.`, ...notes].join(" "));
    }
  } catch (error) {
    console.error(error);
    notify(cleanErrorMessage(error, `${EXPORT_LABELS[format]} export failed.`), "error");
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
    if (printReviewNote) warnings.push(printReviewNote);
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
  window.setTimeout(() => restoreFocus(returnFocus), 0);
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
  if (!handle || !documentOpen || saving || saveFlowActive || !printModal.hidden) return;
  toggleMenu(false);
  toggleExportMenu(false);
  const review = reviewSummary(handle.getReview());
  printReviewNote = review.total
    ? `${capitalize(describeReviewItems(review))} ${review.total === 1 ? "isn't" : "aren't"} printed; the preview shows the document without ${review.total === 1 ? "it" : "them"}.`
    : "";
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

async function createEnvelope() {
  const ctx = envelopeContext;
  if (!ctx) return;
  if (saving) {
    notify("A save is in progress. Try again when it finishes.", "error");
    return;
  }
  const sizeId = $<HTMLSelectElement>("envelope-size").value;
  const size = ENVELOPE_SIZES.find((candidate) => candidate.id === sizeId) ?? ENVELOPE_SIZES[0];
  if (documentOpen && (dirty || !documentIsEmpty(ctx))) {
    syncDirty();
    const choice = await askChoice({
      title: "Replace the document with the envelope?",
      message: dirty
        ? `The envelope replaces “${documentName}” in this window, and its unsaved changes will be lost.`
        : `The envelope replaces “${documentName}” in this window.`,
      choices: [{ id: "cancel", label: "Cancel" }, { id: "replace", label: "Replace", primary: true }],
    });
    if (choice !== "replace" || envelopeContext !== ctx || envelopeModal.hidden || saving) return;
  }
  applyEnvelope(ctx, size, $<HTMLTextAreaElement>("envelope-return").value, $<HTMLTextAreaElement>("envelope-delivery").value);
  resetOriginalLayout();
  clearDocumentNotice();
  resetFidelity();
  envelopeModal.hidden = true;
  originalDocument = null;
  currentPath = null;
  currentFormat = 'docx';
  currentSourceHash = undefined;
  documentSourcePath = null;
  documentOrigin = "new";
  restoredRecoveryId = null;
  documentName = "Envelope";
  documentOpen = true;
  welcome.hidden = true;
  editorHost.classList.add("is-active");
  // The envelope is new content that has never been saved.
  adoptDocumentModel(false);
  scheduleRecovery();
  setTitle();
  refreshHistoryButtons();
  focusEditor();
  notify("Envelope created.");
}

function cancelEnvelope() {
  envelopeModal.hidden = true;
  envelopeContext = null;
  focusEditor();
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
      // Only a file that is really gone leaves the list; a locked, damaged or
      // password-protected file stays, and the real reason is shown.
      const forgetIfMissing = (error: unknown) => {
        if (!isMissingFileError(error)) return false;
        notify("That file is no longer available. It may have been moved, renamed or deleted.", "error");
        void window.simpleDocs.removeRecent(filePath).catch(() => {}).finally(() => void refreshWelcomeLists().catch(() => {}));
        return true;
      };
      if (documentOpen) {
        try {
          await window.simpleDocs.openInNewWindow(filePath);
        } catch (error) {
          if (!forgetIfMissing(error)) notify(openErrorMessage(error), "error");
        }
        return;
      }
      await openWith(() => window.simpleDocs.openPath(filePath), { onError: forgetIfMissing });
    });
  });

  const recoverySection = $("recovery-section");
  const recoveryList = $("recovery-list");
  recoverySection.hidden = recoveries.length === 0;
  recoveryList.innerHTML = recoveries.map((item) => `
    <button class="file-row recovery-row" data-recovery-id="${escapeHtml(item.id)}">
      <span class="file-row-icon">${icons.file}</span>
      <span class="file-row-copy"><strong>${escapeHtml(item.title)}</strong><small${item.sourcePath ? ` title="${escapeHtml(item.sourcePath)}"` : ""}>${item.sourcePath ? `Unsaved changes to ${escapeHtml(fileName(item.sourcePath))}` : "Unsaved recovery copy"}</small></span>
      <time>${humanTime(item.updatedAt)}</time>
    </button>`).join("");
  recoveryList.querySelectorAll<HTMLButtonElement>("[data-recovery-id]").forEach((button) => {
    button.addEventListener("click", () => {
      const recoveryId = button.dataset.recoveryId!;
      void openWith(() => window.simpleDocs.loadRecovery(recoveryId), {
        recoveryId,
        onError: (error) => {
          notify(`That recovery copy could not be opened. ${sentence(cleanErrorMessage(error, ""))}`.trim(), "error");
          return true;
        },
      });
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
  // Closing answers an open save question (and any other open question) with Cancel.
  choiceResolve?.(choiceCancelId);
  cancelFormDialog();
  cancelSettingsDialog();
  closePopupMenu();
  syncDirty();
  if (!dirty || !documentOpen) {
    window.simpleDocs.confirmClose();
    return;
  }
  $("close-document-name").textContent = documentName;
  const review = handle ? reviewSummary(handle.getReview()) : null;
  const note = $("close-review-note");
  note.textContent = review?.total
    ? `${capitalize(describeReviewItems(review))} can't be saved in this version yet and will be lost when this window closes.`
    : "";
  note.hidden = !review?.total;
  closeModal.hidden = false;
  $("save-close").focus();
}

/** Cancel on the close question goes back to the document. */
function cancelCloseQuestion() {
  closeModal.hidden = true;
  if (!focusEditor()) $("close-button").focus();
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
  if (!pasteOptions.hidden && !pasteOptions.contains(target)) hidePasteOptions();
});
$("minimize-button").addEventListener("click", () => window.simpleDocs.minimize());
$("maximize-button").addEventListener("click", () => window.simpleDocs.toggleMaximize());
$("close-button").addEventListener("click", requestClose);
$("focus-button").addEventListener("click", () => window.simpleDocs.toggleFullscreen());
$("cancel-close").addEventListener("click", cancelCloseQuestion);
$("undo-button").addEventListener("click", () => { if (editorShown() && bridge?.undo()) focusEditor(); });
$("redo-button").addEventListener("click", () => { if (editorShown() && bridge?.redo()) focusEditor(); });
// Like the ribbon, the title bar's Undo and Redo leave the caret where it is.
for (const id of ["undo-button", "redo-button"]) $(id).addEventListener("mousedown", (event) => event.preventDefault());
$("author-button").addEventListener("click", () => void editAuthorName());
$("discard-close").addEventListener("click", async () => {
  closeModal.hidden = true;
  // Discarding also retires the recovery copies of this work, including a restored one.
  dirty = false;
  await forgetAllRecoveries();
  window.simpleDocs.confirmClose();
});
$("save-close").addEventListener("click", async () => {
  closeModal.hidden = true;
  if (await saveDocument(false)) window.simpleDocs.confirmClose();
  else focusEditor();
});
$("document-notice-close").addEventListener("click", () => { $("document-compatibility").hidden = true; });
toast.addEventListener("click", () => { toast.hidden = true; });
choiceModal.addEventListener("pointerdown", (event) => { if (event.target === choiceModal) event.preventDefault(); });
$("envelope-cancel").addEventListener("click", cancelEnvelope);
$("envelope-create").addEventListener("click", () => void createEnvelope());
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

/**
 * Whether an editing shortcut may act on the document from this key target:
 * the document itself (or nothing focused), never the title bar, a text box or
 * a dialog field. Find and Replace also work from the find bar.
 */
function documentShortcutTarget(target: EventTarget | null, command: ShortcutCommand) {
  if (!(target instanceof Element) || target === document.body || target === document.documentElement) return true;
  if (target.closest(".cw-float-panel")) return command === "find" || command === "replace";
  if (!editorHost.contains(target)) return false;
  if (target.matches('[contenteditable="true"][role="textbox"]')) return true;
  return !target.closest('input, textarea, select, [contenteditable="true"]');
}

/** Word and Docs shortcuts the engine lacks (see src/ui/shortcuts.ts); physical keys, so they work on any layout. */
function handleShortcut(event: KeyboardEvent) {
  const binding = findShortcut(event);
  if (!binding) return false;
  if (binding.scope === "app") {
    event.preventDefault();
    event.stopImmediatePropagation();
    if (binding.command === "save-as") void saveDocument(true);
    else if (binding.command === "close") requestClose();
    else if (binding.command === "command-search") commandSearch.focus();
    return true;
  }
  const search = binding.command === "find" || binding.command === "replace";
  if (!commands || !documentOpen || !welcome.hidden) {
    // Without a document, the engine's own Ctrl+F would open a find bar over the welcome screen.
    if (search) {
      event.preventDefault();
      event.stopImmediatePropagation();
    }
    return search;
  }
  if (editorHost.hidden) {
    // Page view shows the original pages; the editor behind them is hidden and
    // must not be searched or changed there. Find switches to Edit document.
    if (!search) return false;
    event.preventDefault();
    event.stopImmediatePropagation();
    setOriginalLayoutView(false);
    notify("Switched to Edit document to search.");
    void commands.run(binding.command);
    return true;
  }
  if (!documentShortcutTarget(event.target, binding.command)) return false;
  event.preventDefault();
  event.stopImmediatePropagation();
  void commands.run(binding.command);
  return true;
}

/**
 * Tab in a paragraph inserts a tab (Word and Docs) instead of moving keyboard
 * focus out of the document. WordCanvas handles Tab first in tables, list items
 * and selected objects; a selection across paragraphs indents them.
 */
function handleDocumentTab(outdent: boolean) {
  if (!bridge || !handle || handle.getMode() === "view") return;
  const selection = bridge.getSelection();
  if (!selection) return;
  if (selection.anchor.blockId !== selection.focus.blockId) {
    bridge.clickRibbonItem(outdent ? RIBBON_ITEM_IDS.decreaseIndent : RIBBON_ITEM_IDS.increaseIndent);
    return;
  }
  if (outdent) {
    const block = findParagraph(handle.getDocument(), selection.focus.blockId);
    if (block && isCollapsed(selection) && selection.focus.offset === 0 && Number(block.style.indentLeftPx) > 0) bridge.clickRibbonItem(RIBBON_ITEM_IDS.decreaseIndent);
    return;
  }
  handle.insertText("\t");
}

// Bubble phase on the host: WordCanvas's own Tab handling (inside the editor) runs first.
editorHost.addEventListener("keydown", (event) => {
  if (event.key !== "Tab" || event.defaultPrevented || event.ctrlKey || event.altKey || event.metaKey || event.isComposing) return;
  if (!(event.target instanceof HTMLElement) || !event.target.matches('[contenteditable="true"][role="textbox"]')) return;
  event.preventDefault();
  handleDocumentTab(event.shiftKey);
});

document.addEventListener("keydown", (event) => {
  const modifier = event.ctrlKey || event.metaKey;
  const modal = activeModal(document);
  if (modal) {
    if (trapModalFocus(event, modal as HTMLElement)) return;
    if (event.key === "Escape") {
      if (modal === closeModal) {
        event.preventDefault();
        cancelCloseQuestion();
      } else if (modal === envelopeModal) {
        event.preventDefault();
        cancelEnvelope();
      } else if (modal === printModal) {
        event.preventDefault();
        if (!printSubmitting) closePrintDialog();
      } else if (modal === choiceModal) {
        event.preventDefault();
        choiceResolve?.(choiceCancelId);
      }
      return;
    }
    if (isModalBlockedShortcut(event)) {
      event.preventDefault();
      event.stopImmediatePropagation();
    }
    return;
  }
  // Physical letters, so Ctrl+S is Ctrl+S on a Cyrillic layout too; never with
  // Alt, which is AltGr typing a character (Ctrl+Alt+S is "ś" in Polish).
  const letter = modifier && !event.altKey ? shortcutLetter(event) : null;
  // Ctrl+Shift+V: Chromium pastes the clipboard's plain text, and the paste
  // handler formats it like the text at the caret.
  if (letter === "v" && event.shiftKey) plainPasteRequestedAt = performance.now();
  if (letter === "s") {
    event.preventDefault();
    void saveDocument(event.shiftKey);
    return;
  }
  if (letter === "o" && !event.shiftKey) {
    event.preventDefault();
    void chooseOpen();
    return;
  }
  if (letter === "n" && !event.shiftKey) {
    event.preventDefault();
    void window.simpleDocs.newWindow();
    return;
  }
  if (letter === "p" && !event.shiftKey) {
    event.preventDefault();
    void printDocument();
    return;
  }
  if (letter === "e" && event.shiftKey) {
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
  handleShortcut(event);
}, true);

// Unsaved state is not guessed from keys, clicks or menus: syncDirty() runs on
// every committed model change and review change (see initialize()).

window.addEventListener("beforeunload", releasePrintPreview);
window.addEventListener("beforeunload", () => { if (sourceLayout?.url) URL.revokeObjectURL(sourceLayout.url); });
window.simpleDocs.onDocumentShortcut((action) => {
  if (activeModal(document)) return;
  if (saving && action !== 'save' && action !== 'save-as') return;
  if (action === 'print') void printDocument();
  else if (action === 'save') void saveDocument();
  else if (action === 'save-as') void saveDocument(true);
  else if (action === 'export') toggleExportMenu();
  // Main routes Ctrl+O here on non-Latin keyboard layouts.
  else if ((action as string) === 'open') void chooseOpen();
});

// ---------------------------------------------------------------------------
// Command search (Alt+Q): every ribbon control by name, Simple's own commands
// and the editing commands that have no button.

/** Ribbon items command search leaves out because an editing command covers them better. */
// The proofing switches and Paragraph settings have their own, better-worded entries (see searchCommands).
const SEARCH_SKIPPED_RIBBON_ITEMS = new Set([RIBBON_ITEM_IDS.findReplace, "home.editing.replace", RIBBON_ITEM_IDS.hyperlink, "home.styles.gallery", "simple.proofing.spelling", "simple.proofing.autoformat", "simple.paragraph.settings"]);
const SEARCH_SUGGESTIONS = ["cmd:link", "cmd:comment", "cmd:find", "cmd:replace", `ribbon:${RIBBON_ITEM_IDS.insertTable}`, `ribbon:${RIBBON_ITEM_IDS.insertImage}`, "mode:suggest", "simple:export-pdf"];

function ribbonSearchCommands(): SearchCommand[] {
  const tabs = new Map([...editorHost.querySelectorAll<HTMLElement>("[data-ribbon-tab]")].map((tab) => [tab.dataset.ribbonTab ?? "", tab.textContent?.trim() ?? ""]));
  const found: SearchCommand[] = [];
  for (const element of editorHost.querySelectorAll<HTMLElement>("[data-ribbon-item]")) {
    const id = element.dataset.ribbonItem ?? "";
    if (!id || SEARCH_SKIPPED_RIBBON_ITEMS.has(id) || /\.sep(?:-\d+)?$/.test(id)) continue;
    const control = element.matches("button, select, input") ? element : element.querySelector<HTMLElement>(":scope > button, select, input, button");
    if (!control) continue;
    const described = describeRibbonTitle(control.title || element.title || control.getAttribute("aria-label") || "");
    const { label, detail } = described;
    const combo = ribbonItemShortcut(id);
    const shortcut = described.shortcut ?? (combo ? formatCombo(combo) : undefined);
    if (label.length < 2) continue;
    const tab = tabs.get(element.closest<HTMLElement>("[data-ribbon-panel]")?.dataset.ribbonPanel ?? "") ?? "";
    const group = element.closest(".rib-group")?.querySelector(".rib-label")?.textContent?.trim() ?? "";
    const disabled = (control as HTMLButtonElement).disabled || control.getAttribute("aria-disabled") === "true";
    found.push({
      id: `ribbon:${id}`,
      label,
      group: tab,
      keywords: [group, tab, ...(detail ? [detail] : [])],
      ...(shortcut ? { shortcut } : {}),
      ...(!editorShown() ? { hint: "Switch to Edit document first." } : detail ? { hint: detail } : {}),
      enabled: !disabled && editorShown(),
      run: () => {
        commands?.showRibbonTab(element);
        // The font list is Simple's font box (the editor's own list stays hidden behind it).
        if (id === "home.font.font-family" && fontPicker) {
          fontPicker.open();
          return;
        }
        if (control instanceof HTMLSelectElement || control instanceof HTMLInputElement) {
          control.focus();
          if (control instanceof HTMLSelectElement) try { control.showPicker(); } catch { /* the select stays focused */ }
          else control.select();
          return;
        }
        bridge?.clickRibbonItem(id);
      },
    });
  }
  for (const card of editorHost.querySelectorAll<HTMLElement>(".rib-gallery .style-card")) {
    const name = card.querySelector(".name")?.textContent?.replace(/\s*ⓐ$/, "").trim() ?? "";
    if (!name || /^(?:normal|heading [1-3])$/i.test(name)) continue;
    found.push({ id: `style:${name}`, label: `${name} style`, group: "Home", keywords: ["style", "styles", "paragraph style"], enabled: editorShown(), hint: "Switch to Edit document first.", run: () => { commands?.showRibbonTab(card); card.click(); } });
  }
  return found;
}

function searchCommands(): SearchCommand[] {
  const list: SearchCommand[] = [];
  const hint = (command: ShortcutCommand) => {
    const combo = shortcutFor(command);
    return combo ? { shortcut: formatCombo(combo) } : {};
  };
  const exportDisabled = documentExportDisabled(documentOpen, saving);
  list.push(
    { id: "simple:new", label: "New window", group: "File", keywords: ["new document", "blank document"], shortcut: "Ctrl+N", enabled: true, run: () => void window.simpleDocs.newWindow() },
    { id: "simple:open", label: "Open document", group: "File", keywords: ["open file", "browse"], shortcut: "Ctrl+O", enabled: true, run: () => void chooseOpen() },
    { id: "simple:save", label: "Save", group: "File", shortcut: "Ctrl+S", enabled: documentOpen, run: () => void saveDocument(false) },
    { id: "simple:save-as", label: "Save as", group: "File", keywords: ["save copy", "rename"], shortcut: "F12", enabled: documentOpen, run: () => void saveDocument(true) },
    { id: "simple:print", label: "Print", group: "File", keywords: ["printer", "paper"], shortcut: "Ctrl+P", enabled: documentOpen && !saving, run: () => void printDocument() },
    ...(["docx", "pdf", "html", "md", "txt"] as DocumentExportFormat[]).map((format): SearchCommand => ({
      id: `simple:export-${format}`,
      label: `Export as ${EXPORT_LABELS[format]}`,
      group: "File",
      keywords: ["export", "convert", "download", format],
      enabled: !exportDisabled,
      hint: "Open or create a document first.",
      run: () => void exportDocument(format),
    })),
    { id: "simple:undo", label: "Undo", group: "Edit", shortcut: "Ctrl+Z", enabled: !$<HTMLButtonElement>("undo-button").disabled, run: () => { bridge?.undo(); } },
    { id: "simple:redo", label: "Redo", group: "Edit", shortcut: "Ctrl+Y", enabled: !$<HTMLButtonElement>("redo-button").disabled, run: () => { bridge?.redo(); } },
    { id: "simple:focus", label: "Focus mode", group: "View", keywords: ["full screen", "distraction free"], shortcut: "F11", enabled: true, run: () => window.simpleDocs.toggleFullscreen() },
    { id: "simple:author", label: "Author name for comments", group: "Review", keywords: ["user name", "my name", "initials", "suggestions"], enabled: true, run: () => void editAuthorName() },
    { id: "simple:close", label: "Close window", group: "File", keywords: ["exit", "quit"], shortcut: "Ctrl+W", enabled: true, run: () => requestClose() },
  );
  if (sourceLayout) {
    list.push(
      { id: "simple:page-view", label: "Page view", group: "View", keywords: ["original layout", "office rendering"], enabled: documentOpen, run: () => setOriginalLayoutView(true) },
      { id: "simple:edit-view", label: "Edit document", group: "View", keywords: ["editing view"], enabled: documentOpen, run: () => setOriginalLayoutView(false) },
    );
  }
  if (!documentOpen || !handle) return list;
  for (const entry of EDITOR_SEARCH_COMMANDS) {
    list.push({ id: `cmd:${entry.command}`, label: entry.label, group: entry.group, keywords: entry.keywords ?? [], ...hint(entry.command), enabled: editorShown(), hint: "Switch to Edit document first.", run: () => void commands?.run(entry.command) });
  }
  const modes: Array<[EditMode, string, string[]]> = [["edit", "Editing mode", ["edit"]], ["suggest", "Suggesting (track changes)", ["track changes", "review", "revisions"]], ["view", "Viewing mode", ["read only", "read mode"]]];
  for (const [mode, label, keywords] of modes) {
    list.push({ id: `mode:${mode}`, label, group: "Review", keywords, enabled: editorShown() && handle.getMode() !== mode, hint: handle.getMode() === mode ? "Already on." : "Switch to Edit document first.", run: () => { handle?.setMode(mode); } });
  }
  // Word features: proofing switches, paragraph spacing and headers/footers.
  const changeable = editorShown() && handle.getMode() !== "view";
  const changeHint = editorShown() ? "Switch from Viewing to Editing first." : "Switch to Edit document first.";
  const spaces = changeable ? paragraphSpaceLabels() : { before: "Add space before paragraph", after: "Add space after paragraph" };
  const firstPage = hasDifferentFirstPage(headerFooterDocument());
  list.push(
    { id: "word:spelling", label: spelling?.enabled ? "Turn off spelling as you type" : "Check spelling as you type", group: "Review", keywords: ["spelling", "spell check", "proofing", "misspelled", "dictionary"], enabled: Boolean(spelling), run: () => toggleSpelling() },
    { id: "word:autoformat", label: autoFormatEnabled ? "Turn off AutoFormat as you type" : "AutoFormat as you type", group: "Review", keywords: ["autocorrect", "smart quotes", "curly quotes", "dashes", "automatic numbered list", "autoformat"], enabled: Boolean(autoFormat?.active), hint: "This version of the editor can't format as you type.", run: () => toggleAutoFormat() },
    { id: "word:paragraph", label: "Indents and spacing", group: "Home", keywords: ["paragraph settings", "paragraph dialog", "indent", "indentation", "hanging indent", "first line indent", "spacing", "space before", "space after", "line spacing options"], enabled: changeable, hint: changeHint, run: () => void openParagraphSettings() },
    { id: "word:space-before", label: spaces.before, group: "Home", keywords: ["paragraph spacing", "space before"], enabled: changeable, hint: changeHint, run: () => { toggleParagraphSpace("before"); } },
    { id: "word:space-after", label: spaces.after, group: "Home", keywords: ["paragraph spacing", "space after"], enabled: changeable, hint: changeHint, run: () => { toggleParagraphSpace("after"); } },
    { id: "word:edit-header", label: "Edit header", group: "Insert", keywords: ["header", "top of page", "running head"], enabled: changeable, hint: changeHint, run: () => void editHeaderFooter("header") },
    { id: "word:edit-footer", label: "Edit footer", group: "Insert", keywords: ["footer", "bottom of page"], enabled: changeable, hint: changeHint, run: () => void editHeaderFooter("footer") },
    { id: "word:remove-header", label: "Remove header", group: "Insert", keywords: ["delete header"], enabled: changeable && hasBand(headerFooterDocument(), "header"), hint: changeable ? "This document has no header." : changeHint, run: () => { removeHeaderFooter("header"); } },
    { id: "word:remove-footer", label: "Remove footer", group: "Insert", keywords: ["delete footer"], enabled: changeable && hasBand(headerFooterDocument(), "footer"), hint: changeable ? "This document has no footer." : changeHint, run: () => { removeHeaderFooter("footer"); } },
    { id: "word:page-numbers", label: "Page numbers", group: "Insert", keywords: ["page number", "numbering pages", "page x of y", "start page number"], enabled: changeable, hint: changeHint, run: () => void openPageNumbers() },
    { id: "word:different-first-page", label: firstPage ? "Same header and footer on the first page" : "Different first page", group: "Insert", keywords: ["title page", "first page header", "first page footer"], enabled: changeable, hint: changeHint, run: () => { toggleDifferentFirstPage(); } },
  );
  return [...list, ...ribbonSearchCommands()];
}

const commandSearch = createCommandSearch({
  input: $<HTMLInputElement>("command-search-input"),
  list: $("command-search-list"),
  container: $("command-search"),
  getCommands: searchCommands,
  suggestedIds: SEARCH_SUGGESTIONS,
  storage: localSettings(),
  beforeRun: () => { focusEditor(); },
  onEscape: () => { focusEditor(); },
  onUnavailable: (command) => notify(command.hint ? `${command.label} isn't available: ${command.hint.replace(/^./, (c) => c.toLowerCase())}` : `${command.label} isn't available right now.`, "error"),
});

function annotateRibbonShortcuts() {
  commands?.annotateRibbon(withShortcutHint);
}

// ---------------------------------------------------------------------------
// Clipboard and drag and drop (DOC-008, DOC-012, DOC-SIE-25).
//
// Ctrl+V, the ribbon's Paste and the context menu's Paste share one path:
// pictures (a screenshot, a browser's Copy image) are inserted as pictures;
// formatted content (web pages, Word, Docs) keeps its lists, tables, headings and
// pictures through the import mapper (src/importers/paste.ts); plain text takes
// the formatting at the caret. Every paste is one undo step. Ctrl+Shift+V pastes
// text only, and a web address pasted over selected text links that text. After
// a formatted paste a small Paste options chip offers Keep source formatting,
// Merge formatting and Keep text only (pressing Ctrl alone opens it). Nothing is
// ever downloaded: pictures stored on the web are left out, with a note.

/** Files Simple Docs opens (mirrors OPEN_FORMATS in electron/docx-files.cjs). */
const OPENABLE_DOCUMENT = /\.(?:docx|docm|dotx|dotm|doc|rtf|odt|html?|md|markdown|txt)$/i;
const PICTURE_FILE = /\.(?:png|jpe?g|jfif|gif|webp|bmp|svg|ico|avif)$/i;
const MAX_PICTURE_BYTES = 40 * 1024 * 1024;
const MAX_PICTURE_SIDE_PX = 4096;
const MAX_PASTED_PICTURES = 20;

type ClipboardPayload = ClipboardContent & { pictures: File[] };
type PasteEditorState = { doc: PasteDocument; selection: DocSelection | null };
type PasteEditor = {
  dispatch(command: (state: PasteEditorState) => { ops: unknown[]; selectionAfter: DocSelection | null; origin: string } | null): void;
  paste?: () => void;
};

const isPictureFile = (file: File) => /^image\//i.test(file.type) || PICTURE_FILE.test(file.name);

/** The engine's live editor (window.__cw.editor), when it can dispatch a transaction. */
function pasteEditor(): PasteEditor | null {
  const editor = (globalThis as { __cw?: { editor?: Partial<PasteEditor> } }).__cw?.editor;
  return editor && typeof editor.dispatch === "function" ? editor as PasteEditor : null;
}

const pasteSession = Math.random().toString(36).slice(2, 8);
let pasteSerial = 0;
let pasteBlockSerial = 0;
const newPasteBlockId = () => `pasted-${pasteSession}-${(pasteBlockSerial++).toString(36)}`;

/** The engine's hidden text box, where pastes land while the document has focus. */
function documentInput(target: EventTarget | null) {
  return target instanceof HTMLElement && editorHost.contains(target) && target.matches('[contenteditable="true"][role="textbox"]');
}

/** The editing view is on screen, nothing modal is open, and the mode allows edits. */
function canPaste() {
  return Boolean(handle && bridge && editorShown() && !activeModal(document) && handle.getMode() !== "view" && pasteEditor());
}

function clipboardPayload(data: DataTransfer): ClipboardPayload {
  const read = (type: string) => {
    try {
      return data.getData(type);
    } catch {
      return "";
    }
  };
  return { html: read("text/html"), text: read("text/plain"), rtf: read("text/rtf"), pictures: [...data.files].filter(isPictureFile) };
}

function caretCharStyle(doc: PasteDocument, selection: DocSelection | null): CharStyle | null {
  const paragraph = selection ? findParagraph(doc as unknown as DocumentLike, selection.focus.blockId) : null;
  return paragraph ? runStyleAt(paragraph, selection!.focus.offset) : null;
}

/** What text the source leaves unformatted looks like: the text at the caret (body text when the caret is in a heading). */
function pasteBaseStyle(doc: PasteDocument, selection: DocSelection | null): Partial<CharStyle> {
  const paragraph = selection ? findParagraph(doc as unknown as DocumentLike, selection.focus.blockId) : null;
  const normal = resolveStyleChar(doc.stylesheet as StylesheetLike | null | undefined, doc.stylesheet?.defaultStyleId);
  const heading = /^(?:Heading\d|Title|Subtitle)$/.test(String(paragraph?.style.namedStyle ?? ""));
  const caret = !heading ? caretCharStyle(doc, selection) : null;
  return {
    fontFamily: caret?.fontFamily ?? normal.fontFamily,
    fontSizePx: caret?.fontSizePx ?? normal.fontSizePx,
    color: caret?.color ?? normal.color,
  };
}

function pasteTextWidth(doc: PasteDocument) {
  const section = doc.section;
  const width = section ? section.pageWidthPx - section.marginPx.left - section.marginPx.right : 0;
  return width > 48 ? width : 624;
}

/** Applies a paste as one engine transaction (one undo step); not applied when it cannot be placed here. */
function dispatchPaste(content: PasteContent): { applied: boolean; warnings: string[] } {
  const editor = pasteEditor();
  if (!editor || !handle) return { applied: false, warnings: [] };
  const before = handle.getDocument();
  const built: { warnings: string[] | null } = { warnings: null };
  try {
    editor.dispatch((state) => {
      const transaction = buildPasteTransaction(state.doc, state.selection, content, newPasteBlockId);
      if (!transaction) return null;
      built.warnings = transaction.warnings;
      return { ops: transaction.ops, selectionAfter: transaction.selectionAfter, origin: "paste" };
    });
  } catch (error) {
    console.error("The paste could not be applied", error);
    syncDirty("paste");
    notify("Simple Docs couldn't paste that here.", "error");
    return { applied: false, warnings: [] };
  }
  return { applied: built.warnings !== null && handle.getDocument() !== before, warnings: built.warnings ?? [] };
}

/** When Ctrl+Shift+V was pressed (performance.now()), or -Infinity. */
let plainPasteRequestedAt = Number.NEGATIVE_INFINITY;
let lastPaste: { content: ClipboardPayload; mode: PasteMode; revision: number } | null = null;
let applyingPasteOption = false;

/**
 * Pastes clipboard (or dropped) content at the selection. Returns false when this
 * path has nothing to paste or cannot place it, so the engine's own paste can run.
 */
function pasteContent(content: ClipboardPayload, mode: PasteMode, options: { offerOptions?: boolean } = {}): boolean {
  if (!handle || !bridge || !canPaste()) return false;
  const doc = handle.getDocument() as unknown as PasteDocument;
  const selection = bridge.getSelection();
  if (!selection) return false;
  const text = content.text ?? "";
  const html = content.html?.trim() ? content.html : "";
  const rtf = content.rtf?.trim() ? content.rtf : "";
  const formatted = Boolean(html || rtf);

  // A screenshot or a browser's Copy image: the picture itself.
  if (mode !== "text" && content.pictures.length && (!html || imageOnlyHtml(html)) && !rtf) {
    hidePasteOptions();
    void insertPictures(content.pictures, { replaceSelection: true });
    return true;
  }
  // A web address pasted over selected text links the text instead of replacing it.
  const link = mode === "text" ? null : pastedLink(text);
  if (link && !isCollapsed(selection)) {
    const selected = selectionRanges(doc as unknown as DocumentLike, selection).map((range) => paragraphText(range.block).slice(range.start, range.end)).join("\n");
    if (selected.trim() && !selected.includes("\n") && !looksLikeAddress(selected) && dispatchPaste({ kind: "link", url: link }).applied) {
      hidePasteOptions();
      return true;
    }
  }
  let fragment = null;
  // Text only needs the markup just when the clipboard has no plain text.
  if (formatted && (mode !== "text" || !text)) {
    fragment = fragmentFromClipboard({ html, rtf }, {
      base: pasteBaseStyle(doc, selection),
      maxImageWidthPx: pasteTextWidth(doc),
      idPrefix: `paste-${pasteSession}-${(pasteSerial++).toString(36)}-`,
      footnotes: doc.footnotes ?? {},
    });
  }
  if (fragment && mode !== "text") {
    const shaped = mode === "merge" ? mergeFragment(fragment, caretCharStyle(doc, selection) ?? pasteBaseStyle(doc, selection)) : fragment;
    const result = dispatchPaste({ kind: "fragment", fragment: shaped, mode });
    if (!result.applied) return false;
    finishPaste(content, mode, [...fragment.warnings, ...result.warnings], options.offerOptions !== false);
    return true;
  }
  const plain = text || (fragment ? fragmentText(fragment) : "");
  if (!plain) return false;
  const result = dispatchPaste({ kind: "text", text: plain, link: mode === "keep" && !formatted && isCollapsed(selection) ? link : null });
  if (!result.applied) return false;
  finishPaste(content, mode, result.warnings, formatted && options.offerOptions !== false);
  return true;
}

/** After a paste: the options chip and any notes. Never throws (the paste itself is done). */
function finishPaste(content: ClipboardPayload, mode: PasteMode, warnings: readonly string[], offerOptions: boolean) {
  try {
    if (offerOptions && bridge) {
      lastPaste = { content, mode, revision: bridge.revision() };
      showPasteOptions(mode);
    } else {
      hidePasteOptions();
    }
    const notes = [...new Set(warnings.map((warning) => sentence(String(warning))).filter(Boolean))];
    if (notes.length) notify(notes.join(" "));
  } catch (error) {
    console.warn("Paste options could not be shown", error);
  }
}

/**
 * Picture bytes the editor, PDF and Word all show: PNG and upright JPEG as they are;
 * other formats, rotated photos and pictures over 4096 px redrawn (as PNG, or JPEG for photos).
 */
async function pictureForInsert(file: Blob): Promise<{ bytes: Uint8Array; type: string } | null> {
  if (!file.size || file.size > MAX_PICTURE_BYTES) return null;
  const bytes = new Uint8Array(await file.arrayBuffer());
  const type = sniffImageType(bytes) ?? (/^image\/[\w.+-]+$/i.test(file.type) ? file.type.toLowerCase() : null);
  if (!type) return null;
  const size = type === "image/png" || type === "image/jpeg" ? imageNaturalSize(bytes, type) : null;
  const oversized = Boolean(size && Math.max(size.width, size.height) > MAX_PICTURE_SIDE_PX);
  if (!oversized && (type === "image/png" || (type === "image/jpeg" && jpegOrientation(bytes) === 1))) return { bytes, type };
  try {
    let source: CanvasImageSource;
    let width: number;
    let height: number;
    let release = () => {};
    if (type === "image/svg+xml") {
      const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
      const image = new Image();
      image.src = url;
      try {
        await image.decode();
      } finally {
        URL.revokeObjectURL(url);
      }
      // Vector art is drawn at twice its size so it stays sharp when printed.
      width = (image.naturalWidth || 300) * 2;
      height = (image.naturalHeight || 150) * 2;
      source = image;
    } else {
      // Decoding applies the EXIF orientation, so photos come out upright.
      const bitmap = await createImageBitmap(new Blob([bytes as BlobPart], { type }));
      width = bitmap.width;
      height = bitmap.height;
      source = bitmap;
      release = () => bitmap.close();
    }
    try {
      const scale = Math.min(1, MAX_PICTURE_SIDE_PX / Math.max(width, height));
      const canvas = new OffscreenCanvas(Math.max(1, Math.round(width * scale)), Math.max(1, Math.round(height * scale)));
      canvas.getContext("2d")!.drawImage(source, 0, 0, canvas.width, canvas.height);
      const photo = type === "image/jpeg";
      const blob = await canvas.convertToBlob(photo ? { type: "image/jpeg", quality: 0.92 } : { type: "image/png" });
      return { bytes: new Uint8Array(await blob.arrayBuffer()), type: photo ? "image/jpeg" : "image/png" };
    } finally {
      release();
    }
  } catch (error) {
    console.warn("The picture could not be read", error);
    return null;
  }
}

/** Inserts pictures at the caret (or the first at a drop point), each as one undo step. */
async function insertPictures(files: File[], options: { at?: { clientX: number; clientY: number }; replaceSelection?: boolean } = {}) {
  if (!bridge) return;
  // A paste replaces the selection, as typing does.
  if (options.replaceSelection && !isCollapsed(bridge.getSelection())) dispatchPaste({ kind: "text", text: "" });
  let inserted = 0;
  let unreadable = 0;
  let refused = 0;
  for (const [index, file] of files.slice(0, MAX_PASTED_PICTURES).entries()) {
    const picture = await pictureForInsert(file);
    if (!picture) {
      unreadable += 1;
      continue;
    }
    if (await bridge.insertImageBytes(picture.bytes, picture.type, index === 0 && options.at ? { at: options.at } : {})) inserted += 1;
    else refused += 1;
  }
  const notes: string[] = [];
  if (unreadable) notes.push(unreadable === 1 && files.length === 1 ? "This picture can’t be inserted: its format isn’t supported. Use PNG, JPEG, GIF, WebP, BMP or SVG." : `${unreadable} pictures couldn’t be inserted: their format isn’t supported.`);
  if (refused) notes.push(refused === 1 && files.length === 1 ? "The picture can’t be inserted here. Put the cursor in the document text or a table cell." : `${refused} pictures couldn’t be inserted here.`);
  if (files.length > MAX_PASTED_PICTURES) notes.push(`Only the first ${MAX_PASTED_PICTURES} pictures were inserted.`);
  if (notes.length) notify(notes.join(" "), inserted ? "normal" : "error");
  if (inserted) focusEditor();
}

/** Reads the clipboard for the ribbon's and the context menu's Paste. */
async function pasteFromSystemClipboard() {
  focusEditor();
  const content: ClipboardPayload = { html: "", text: "", rtf: "", pictures: [] };
  try {
    for (const item of await navigator.clipboard.read()) {
      for (const type of item.types) {
        const blob = await item.getType(type);
        if (type === "text/html" && !content.html) content.html = await blob.text();
        else if (type === "text/plain" && !content.text) content.text = await blob.text();
        else if (/^image\//i.test(type)) content.pictures.push(new File([blob], "Pasted picture", { type }));
      }
    }
  } catch {
    try {
      content.text = await navigator.clipboard.readText();
    } catch {
      notify("Simple Docs couldn’t read the clipboard. Press Ctrl+V to paste.", "error");
      return;
    }
  }
  if (!content.html && !content.text && !content.pictures.length) {
    notify("There is nothing on the clipboard to paste.");
    return;
  }
  // Content controls and other places this path does not handle keep the engine's own paste.
  let handled = false;
  try {
    handled = pasteContent(content, "keep");
  } catch (error) {
    console.error("The paste could not be prepared", error);
  }
  if (!handled) pasteEditor()?.paste?.();
}

// The paste options chip.
const pasteOptions = $("paste-options");
const pasteOptionsButton = $<HTMLButtonElement>("paste-options-button");
const pasteOptionsMenu = $("paste-options-menu");
const pasteOptionItems = () => [...pasteOptionsMenu.querySelectorAll<HTMLButtonElement>("[data-paste-mode]")];

function hidePasteOptions() {
  lastPaste = null;
  togglePasteOptionsMenu(false);
  pasteOptions.hidden = true;
}

/** Keeps the chip just after the caret, inside the editor; hidden while the caret is scrolled away. */
function placePasteOptions() {
  if (pasteOptions.hidden) return;
  const caret = editorHost.querySelector<HTMLElement>('[contenteditable="true"][role="textbox"]')?.getBoundingClientRect();
  const area = editorHost.getBoundingClientRect();
  const visible = Boolean(caret && editorShown() && caret.height > 0 && caret.bottom > area.top && caret.top < area.bottom && caret.left >= area.left - 8 && caret.left <= area.right);
  pasteOptions.classList.toggle("is-offscreen", !visible);
  if (!visible || !caret) return;
  const width = pasteOptionsButton.offsetWidth || 44;
  const height = pasteOptionsButton.offsetHeight || 28;
  pasteOptions.style.left = `${Math.round(Math.min(Math.max(area.left + 4, caret.right + 6), area.right - width - 12))}px`;
  pasteOptions.style.top = `${Math.round(Math.min(Math.max(area.top + 4, caret.bottom + 4), area.bottom - height - 12))}px`;
}

function showPasteOptions(mode: PasteMode) {
  for (const item of pasteOptionItems()) item.setAttribute("aria-checked", String(item.dataset.pasteMode === mode));
  togglePasteOptionsMenu(false);
  pasteOptions.hidden = false;
  placePasteOptions();
  // The engine moves its text box to the new caret when it repaints.
  window.requestAnimationFrame(placePasteOptions);
}

function togglePasteOptionsMenu(open = pasteOptionsMenu.hidden, focusItem = false) {
  const show = open && !pasteOptions.hidden;
  if (show) {
    // The menu opens upward or leftward near the window's edges.
    const chip = pasteOptions.getBoundingClientRect();
    pasteOptions.classList.toggle("opens-up", chip.bottom + 140 > window.innerHeight);
    pasteOptions.classList.toggle("opens-left", chip.left + 244 > window.innerWidth);
  }
  pasteOptionsMenu.hidden = !show;
  pasteOptionsButton.setAttribute("aria-expanded", String(show));
  pasteOptions.classList.toggle("is-open", show);
  if (show && focusItem) (pasteOptionItems().find((item) => item.getAttribute("aria-checked") === "true") ?? pasteOptionItems()[0])?.focus();
}

/** Re-pastes the last paste another way: the paste is undone (one step) and applied again. */
function applyPasteOption(mode: PasteMode) {
  const last = lastPaste;
  togglePasteOptionsMenu(false);
  if (!last || !bridge || bridge.revision() !== last.revision) {
    hidePasteOptions();
    focusEditor();
    return;
  }
  if (mode !== last.mode) {
    applyingPasteOption = true;
    try {
      focusEditor();
      bridge.undo();
      let pasted = false;
      try {
        pasted = pasteContent(last.content, mode);
      } catch (error) {
        console.error("The paste could not be changed", error);
      }
      if (!pasted) {
        // Nothing to paste that way: the original paste comes back.
        bridge.redo();
        hidePasteOptions();
      }
    } finally {
      applyingPasteOption = false;
    }
  }
  focusEditor();
}

// Choosing with the mouse keeps the document's caret and focus.
pasteOptions.addEventListener("mousedown", (event) => event.preventDefault());
pasteOptionsButton.addEventListener("click", () => togglePasteOptionsMenu(undefined, false));
pasteOptionsMenu.addEventListener("click", (event) => {
  const item = event.target instanceof Element ? event.target.closest<HTMLButtonElement>("[data-paste-mode]") : null;
  if (item) applyPasteOption(item.dataset.pasteMode as PasteMode);
});
pasteOptions.addEventListener("keydown", (event) => {
  const items = pasteOptionItems();
  const index = items.indexOf(document.activeElement as HTMLButtonElement);
  if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    togglePasteOptionsMenu(false);
    focusEditor();
  } else if ((event.key === "ArrowDown" || event.key === "ArrowUp") && !pasteOptionsMenu.hidden) {
    event.preventDefault();
    const step = event.key === "ArrowDown" ? 1 : -1;
    items[(index + step + items.length) % items.length]?.focus();
  } else if (event.key === "Tab") {
    togglePasteOptionsMenu(false);
  }
});
editorHost.addEventListener("scroll", placePasteOptions, true);
window.addEventListener("resize", placePasteOptions);
// Like Word: pressing Ctrl on its own opens the chip; Esc in the document dismisses it.
let ctrlPressedAlone = false;
document.addEventListener("keydown", (event) => {
  ctrlPressedAlone = event.key === "Control" && !event.repeat && !event.shiftKey && !event.altKey && !event.metaKey && !pasteOptions.hidden;
  if (event.key === "Escape" && !pasteOptions.hidden && pasteOptionsMenu.hidden && documentInput(event.target)) hidePasteOptions();
}, true);
document.addEventListener("keyup", (event) => {
  if (event.key !== "Control" || !ctrlPressedAlone) return;
  ctrlPressedAlone = false;
  if (!pasteOptions.hidden && !activeModal(document)) togglePasteOptionsMenu(true, true);
}, true);

editorHost.addEventListener("paste", (event) => {
  if (event.defaultPrevented || !documentInput(event.target) || !event.clipboardData) return;
  // Ctrl+Shift+V reaches here as Chromium's plain-text paste.
  const plain = performance.now() - plainPasteRequestedAt < 1500;
  plainPasteRequestedAt = Number.NEGATIVE_INFINITY;
  let handled = false;
  try {
    handled = pasteContent(clipboardPayload(event.clipboardData), plain ? "text" : "keep", { offerOptions: !plain });
  } catch (error) {
    // Nothing was pasted yet: the editor's own paste runs instead.
    console.error("The paste could not be prepared", error);
  }
  if (handled) {
    event.preventDefault();
    event.stopImmediatePropagation();
  }
}, true);

// The ribbon's Paste and the context menu's Paste read the clipboard the same way.
editorHost.addEventListener("click", (event) => {
  const button = event.target instanceof Element ? event.target.closest(`[data-ribbon-item="${RIBBON_ITEM_IDS.paste}"]`) : null;
  if (!button || !canPaste()) return;
  event.preventDefault();
  event.stopImmediatePropagation();
  void pasteFromSystemClipboard();
}, true);
document.addEventListener("mouseup", (event) => {
  const item = event.target instanceof Element ? event.target.closest<HTMLElement>(".cw-menu-item") : null;
  if (!item || item.classList.contains("cw-disabled") || !canPaste()) return;
  if (item.querySelector(".cw-menu-lbl")?.textContent !== "Paste" || item.querySelector(".cw-menu-acc")?.textContent !== "Ctrl+V") return;
  event.preventDefault();
  event.stopImmediatePropagation();
  // Close the menu the way a press outside it does.
  document.body.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
  document.querySelectorAll(".cw-menu").forEach((menu) => menu.remove());
  void pasteFromSystemClipboard();
}, true);

// Drag and drop: documents open, pictures and text land where they are dropped.
type DragKind = "document" | "picture" | "content";
let dragDepth = 0;
let internalDrag = false;
window.addEventListener("dragstart", () => { internalDrag = true; });
window.addEventListener("dragend", () => { internalDrag = false; });

function dragKind(transfer: DataTransfer | null): DragKind | null {
  if (!transfer || internalDrag) return null;
  const files = [...transfer.items].filter((item) => item.kind === "file");
  if (files.length) return files.every((item) => /^image\//i.test(item.type)) ? "picture" : "document";
  return [...transfer.types].some((type) => type === "text/html" || type === "text/plain") ? "content" : null;
}

window.addEventListener("dragenter", (event) => {
  const kind = dragKind(event.dataTransfer);
  if (!kind) return;
  dragDepth += 1;
  // Pictures and text dropped on an open document go where they are dropped.
  if (kind !== "document" && documentOpen) return;
  $("drop-overlay-title").textContent = kind === "document" ? "Open in Simple Docs" : "Start a new document";
  $("drop-overlay-detail").textContent = kind === "document" ? "Drop the document here" : kind === "picture" ? "Drop the picture here" : "Drop the text here";
  $("drop-overlay").hidden = false;
});
window.addEventListener("dragleave", () => {
  if (internalDrag) return;
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) $("drop-overlay").hidden = true;
});
window.addEventListener("dragover", (event) => {
  if (internalDrag) return;
  event.preventDefault();
  if (event.dataTransfer) event.dataTransfer.dropEffect = "copy";
});

async function openDroppedDocument(file: File) {
  const filePath = window.simpleDocs.pathForFile(file);
  if (documentOpen && filePath) {
    await openInNewWindow(filePath);
    return;
  }
  // A file without a folder location (from a mail or archive app) can only
  // open in this window; it never replaces unsaved work.
  syncDirty();
  if (documentOpen && dirty) {
    notify("Save your changes first: this dropped file has no folder location, so it would open in this window and replace them.", "error");
    return;
  }
  await openWith(filePath
    ? () => window.simpleDocs.openPath(filePath)
    : async () => window.simpleDocs.openBytes({ data: new Uint8Array(await file.arrayBuffer()), name: file.name }));
}

/** A document to drop into: a new one from the start screen, the editing view from Page view. */
function editableDocumentForDrop() {
  if (!handle || !bridge) return false;
  if (!documentOpen) beginBlankDocument();
  if (editorHost.hidden) setOriginalLayoutView(false);
  if (handle.getMode() === "view") {
    notify("Switch from Viewing to Editing to add to the document.");
    return false;
  }
  return editorShown();
}

window.addEventListener("drop", (event) => {
  if (internalDrag) return;
  event.preventDefault();
  dragDepth = 0;
  $("drop-overlay").hidden = true;
  const transfer = event.dataTransfer;
  if (!transfer || activeModal(document)) return;
  const files = [...transfer.files];
  const documentFile = files.find((file) => OPENABLE_DOCUMENT.test(file.name));
  if (documentFile) {
    void openDroppedDocument(documentFile);
    return;
  }
  const pictures = files.filter(isPictureFile);
  if (files.length && !pictures.length) {
    notify(`Simple Docs can’t open “${files[0].name}”. It opens Word documents, OpenDocument text, Rich Text, web pages, Markdown and plain text.`);
    return;
  }
  const point = { clientX: event.clientX, clientY: event.clientY };
  const onPage = editorHost.contains(document.elementFromPoint(point.clientX, point.clientY));
  if (pictures.length) {
    if (editableDocumentForDrop()) void insertPictures(pictures, onPage ? { at: point } : {});
    return;
  }
  const content = clipboardPayload(transfer);
  if (!content.html && !content.text) return;
  if (!editableDocumentForDrop()) return;
  const position = onPage ? bridge!.positionFromPoint(point.clientX, point.clientY) : null;
  if (position) bridge!.setSelection({ anchor: position, focus: position });
  focusEditor();
  try {
    pasteContent(content, "keep", { offerOptions: false });
  } catch (error) {
    console.error("The dropped content could not be added", error);
    notify("Simple Docs couldn’t add the dropped content.", "error");
  }
});

window.simpleDocs.onMaximized((maximized) => { $("maximize-button").innerHTML = maximized ? icons.restore : icons.maximize; });
window.simpleDocs.onFullscreen((fullscreen) => {
  document.body.classList.toggle("focus-mode", fullscreen);
  $("focus-button").setAttribute("aria-pressed", String(fullscreen));
});
window.simpleDocs.onCloseRequested(requestClose);

// The shared window guard owns every close path (title bar after the in-app question,
// Alt+F4, taskbar, quit, Windows sign-out). It asks this page about unsaved work and
// shows Save / Don't Save / Cancel natively when needed.
const simpleIO = getSimpleIO();
if (simpleIO) {
  simpleIO.onRequest("close-query", () => {
    syncDirty("close-query");
    return { dirty: dirty && documentOpen, saving, title: documentName, kind: "document", untitled: !currentPath };
  });
  simpleIO.onRequest("save-now", async () => {
    choiceResolve?.(choiceCancelId);
    return saveDocument(false);
  });
  simpleIO.onRequest("discard", async () => {
    closeModal.hidden = true;
    dirty = false;
    await forgetAllRecoveries();
    return true;
  });
  // Windows is ending the session: keep a fresh recovery copy of unsaved work.
  simpleIO.onRequest("recovery-flush", async () => {
    await checkpointRecovery();
    return true;
  });
}
window.simpleDocs.onOpenExternal((filePath) => {
  if (!handle) {
    pendingExternalPath = filePath;
    return;
  }
  if (documentOpen) void openInNewWindow(filePath);
  else void openWith(() => window.simpleDocs.openPath(filePath));
});

// ---------------------------------------------------------------------------
// Word features: spelling as you type (DOC-009), AutoFormat as you type (DOC-013),
// headers, footers and page numbers (DOC-011), paragraph indents and spacing
// (DOC-014) and the font box with every installed font (DOC-018). Model changes go
// through the engine's own operations, one undoable transaction each.

type WordState = { doc: WordDocument; selection: DocSelection | null };
type LayoutPage = {
  index: number;
  widthPx: number;
  heightPx: number;
  marginPx: { top: number; right: number; bottom: number; left: number };
  contentTopPx: number;
  contentBottomPx: number;
  headerSource?: string;
  footerSource?: string;
};
/** The engine's live editor (window.__cw.editor); feature-detected like the bridge does. */
type WordEditor = {
  dispatch(command: (state: WordState) => { ops: unknown[]; selectionAfter: DocSelection | null; origin: string } | null): void;
  setCharStyle?: (style: Partial<CharStyle>) => void;
  focus?: () => void;
  getLayoutTree?: () => { pages: LayoutPage[] };
  getLayoutInfo?: () => { pageCount: number; currentPage: number };
};

function wordEditor(): WordEditor | null {
  const live = (globalThis as { __cw?: { editor?: Partial<WordEditor> } }).__cw?.editor;
  return live && typeof live.dispatch === "function" ? live as WordEditor : null;
}

/** The document can be changed here: it is open in Edit document, and not in Viewing mode. */
function canChangeDocument(what: string) {
  if (!handle || !bridge || !editorShown()) {
    notify(documentOpen ? "Switch to Edit document first." : "Open or create a document first.");
    return false;
  }
  if (handle.getMode() === "view") {
    notify(`Switch from Viewing to Editing to ${what}.`);
    return false;
  }
  return true;
}

/** Applies engine operations built from the current model as ONE undoable edit. */
function dispatchModelOps(build: (doc: WordDocument, selection: DocSelection | null) => unknown[] | null, selectionAfter?: (selection: DocSelection | null) => DocSelection | null): boolean {
  const live = wordEditor();
  if (!live || !handle) return false;
  const before = handle.getDocument();
  try {
    live.dispatch((state) => {
      const ops = build(state.doc, state.selection);
      if (!ops?.length) return null;
      return { ops, selectionAfter: selectionAfter ? selectionAfter(state.selection) : state.selection, origin: "command" };
    });
  } catch (error) {
    console.error("The change could not be applied", error);
    notify("Simple Docs couldn't make that change.", "error");
    return false;
  }
  return handle.getDocument() !== before;
}

/** Replaces part of a paragraph, keeping the formatting of the first replaced letter (spelling corrections). */
function replaceDocumentText(blockId: string, start: number, end: number, text: string): boolean {
  const caret = { blockId, offset: start + text.length };
  return dispatchModelOps((doc) => {
    const block = findParagraph(doc as unknown as DocumentLike, blockId);
    if (!block) return null;
    const length = paragraphText(block).length;
    const from = Math.max(0, Math.min(length, start));
    const to = Math.max(from, Math.min(length, end));
    const style = runStyleAt(block, from + 1) ?? block.runs[0]?.style;
    const ops: unknown[] = [];
    if (to > from) ops.push({ type: "deleteRange", blockId, start: from, end: to });
    if (text) ops.push({ type: "insertText", at: { blockId, offset: from }, text, ...(style ? { style } : {}) });
    return ops;
  }, () => ({ anchor: caret, focus: caret }));
}

// ---- Proofing switches --------------------------------------------------------

function ribbonButton(id: string) {
  return editorHost.querySelector<HTMLButtonElement>(`[data-ribbon-item="${id}"]`);
}

function spellingTitle() {
  if (!spelling) return "Check spelling as you type";
  if (!spelling.available) return "Spelling isn't available: Windows has no spell checker for this language on this computer";
  return spelling.enabled ? `Spelling is checked as you type (${spelling.language}). Click to turn it off` : "Check spelling as you type";
}

function refreshProofingButtons() {
  for (const [id, on, title] of [
    ["simple.proofing.spelling", Boolean(spelling?.enabled && spelling.available), spellingTitle()],
    ["simple.proofing.autoformat", autoFormatEnabled && Boolean(autoFormat?.active), autoFormat?.active === false ? "AutoFormat isn't available in this version of the editor" : autoFormatEnabled ? "AutoFormat as you type is on: smart quotes, dashes, symbols, lists and links. Click to turn it off" : "AutoFormat as you type: smart quotes, dashes, symbols, lists and links"],
  ] as const) {
    const button = ribbonButton(id);
    if (!button) continue;
    button.classList.toggle("active", on);
    button.setAttribute("aria-pressed", String(on));
    button.title = title;
  }
}

function toggleSpelling() {
  if (!spelling) return;
  spelling.setEnabled(!spelling.enabled);
  refreshProofingButtons();
  notify(!spelling.enabled ? "Spelling marks are off." : spelling.available ? "Spelling is checked as you type." : "Windows has no spell checker for this language on this computer.");
  focusEditor();
}

function toggleAutoFormat() {
  autoFormatEnabled = !autoFormatEnabled;
  try {
    window.localStorage.setItem(AUTOFORMAT_KEY, autoFormatEnabled ? "on" : "off");
  } catch {
    // The choice still applies to this window.
  }
  autoFormat?.setOptions({ locale: navigator.language || "en-US", enabled: autoFormatEnabled });
  refreshProofingButtons();
  notify(autoFormatEnabled ? "AutoFormat as you type is on. Press Ctrl+Z right after a change to undo just that change." : "AutoFormat as you type is off.");
  focusEditor();
}

/** Runs on every caret and formatting change (through the Spelling button's state check). */
function onFormatChange(format: unknown) {
  const family = (format as { fontFamily?: string | null } | null)?.fontFamily;
  if (family) fontPicker?.setCurrent(family);
  spelling?.selectionChanged();
}

// ---- Paragraph indents and spacing --------------------------------------------

/** The selected paragraphs; with `placeCaret`, a document without a caret gets one at its start first. */
function selectedParagraphs(placeCaret = true) {
  if (!handle || !bridge) return [];
  let selection = bridge.getSelection();
  if (!selection && placeCaret) {
    focusEditor();
    selection = bridge.getSelection();
  }
  return selectionRanges(handle.getDocument() as unknown as DocumentLike, selection).map((range) => range.block);
}

/** Patches the given paragraphs' styles as one undoable edit. */
function patchParagraphs(ids: readonly string[], patchFor: (style: Partial<ParaStyle>) => Partial<ParaStyle>): boolean {
  return dispatchModelOps((doc) => {
    const ops: unknown[] = [];
    for (const id of ids) {
      const block = findParagraph(doc as unknown as DocumentLike, id);
      if (!block) continue;
      const patch = patchFor(block.style as Partial<ParaStyle>);
      if (Object.keys(patch).length) ops.push({ type: "setParaStyle", blockId: id, patch });
    }
    return ops;
  });
}

async function openParagraphSettings() {
  if (!canChangeDocument("change paragraph formatting")) return false;
  const paragraphs = selectedParagraphs();
  if (!paragraphs.length) return false;
  closeEditorOverlays();
  const returnFocus = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  const result = await openParagraphDialog({
    styles: paragraphs.map((block) => block.style as Partial<ParaStyle>),
    unit: unitForLocale(navigator.language),
    icon: icons.paragraph,
    onClosed: () => window.setTimeout(() => restoreFocus(returnFocus), 0),
  });
  if (!result) return false;
  const changed = patchParagraphs(paragraphs.map((block) => block.id), (style) => paragraphPatch(result.original, result.edited, style));
  window.setTimeout(() => focusEditor(), 0);
  return changed;
}

/** Word's line-spacing menu: Add (12 pt) or Remove space before or after the selected paragraphs. */
function toggleParagraphSpace(which: "before" | "after") {
  if (!canChangeDocument("change paragraph spacing")) return false;
  const paragraphs = selectedParagraphs();
  if (!paragraphs.length) return false;
  const toggle = spaceToggle(paragraphs.map((block) => block.style as Partial<ParaStyle>), which);
  const changed = patchParagraphs(paragraphs.map((block) => block.id), () => (which === "before" ? { spaceBeforePx: toggle.px } : { spaceAfterPx: toggle.px }));
  focusEditor();
  return changed;
}

/** The current Add/Remove labels (never moves focus: command search calls this while it is typed in). */
function paragraphSpaceLabels() {
  const styles = selectedParagraphs(false).map((block) => block.style as Partial<ParaStyle>);
  return { before: spaceToggle(styles, "before").label, after: spaceToggle(styles, "after").label };
}

// The ribbon's line-spacing menu gains Word's spacing toggles and the Indents and spacing dialog.
editorHost.addEventListener("click", (event) => {
  const button = event.target instanceof Element ? event.target.closest(`[data-ribbon-item="${RIBBON_ITEM_IDS.lineSpacing}"]`) : null;
  if (!button) return;
  // The editor opened its menu in its own click handler, just before this one.
  const menu = [...document.querySelectorAll<HTMLElement>(".cw-pop .cw-menu")].at(-1);
  if (!menu || menu.querySelector(".simple-spacing-item") || !/Line Spacing Options/i.test(menu.textContent ?? "")) return;
  const labels = paragraphSpaceLabels();
  const separator = document.createElement("div");
  separator.className = "simple-spacing-separator";
  menu.append(separator);
  for (const [label, run] of [
    [labels.before, () => toggleParagraphSpace("before")],
    [labels.after, () => toggleParagraphSpace("after")],
    ["Indents and spacing…", () => void openParagraphSettings()],
  ] as const) {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "simple-spacing-item";
    item.innerHTML = '<span class="check"></span><span></span>';
    item.lastElementChild!.textContent = label;
    item.addEventListener("click", () => {
      // An outside press closes the menu through the editor's own handler.
      document.body.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
      menu.closest(".cw-pop")?.remove();
      run();
    });
    menu.append(item);
  }
  // Keep the longer menu inside the window.
  const pop = menu.closest<HTMLElement>(".cw-pop");
  const rect = pop?.getBoundingClientRect();
  if (pop && rect && rect.bottom > window.innerHeight - 6) pop.style.top = `${Math.max(6, window.innerHeight - rect.height - 6)}px`;
});

// ---- Headers, footers and page numbers ------------------------------------------

const nextFrame = () => new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()));
const headerFooterDocument = () => handle!.getDocument() as unknown as HeaderFooterDocument;

/**
 * Opens a page's header or footer for typing, the way double-clicking it does (Word's
 * Edit Header): on the page with the caret, creating an empty one first if needed.
 * Esc or a double-click in the page text goes back to the document.
 */
async function editHeaderFooter(kind: BandKind) {
  if (!canChangeDocument(`edit the ${kind}`)) return false;
  const live = wordEditor();
  if (!live?.getLayoutTree) return false;
  const created = ensureBandOps(headerFooterDocument(), kind);
  if (created.length && !dispatchModelOps(() => created)) return false;
  await nextFrame();
  const tree = live.getLayoutTree();
  const pageIndex = Math.max(0, Math.min(tree.pages.length - 1, (live.getLayoutInfo?.().currentPage ?? 1) - 1));
  const page = tree.pages[pageIndex];
  const app = editorHost.querySelector<HTMLElement>(".cw-app");
  const element = app?.querySelector<HTMLElement>(`[data-page="${page?.index ?? pageIndex}"]`);
  if (!page || !app || !element) return false;
  const top = kind === "header" ? 0 : page.contentBottomPx;
  const bottom = kind === "header" ? page.contentTopPx : page.heightPx;
  let rect = element.getBoundingClientRect();
  let zoom = rect.width / page.widthPx;
  const view = app.getBoundingClientRect();
  const bandTop = rect.top + top * zoom;
  const bandBottom = rect.top + bottom * zoom;
  if (bandTop < view.top || bandBottom > view.bottom) {
    app.scrollTop += (bandTop + bandBottom) / 2 - (view.top + view.bottom) / 2;
    await nextFrame();
    rect = element.getBoundingClientRect();
    zoom = rect.width / page.widthPx;
  }
  const point = { clientX: rect.left + (page.marginPx.left + 4) * zoom, clientY: rect.top + ((top + bottom) / 2) * zoom };
  const init = { bubbles: true, cancelable: true, view: window, button: 0, buttons: 1, detail: 2, ...point };
  const source = ((kind === "header" ? live.getLayoutTree().pages[pageIndex]?.headerSource : live.getLayoutTree().pages[pageIndex]?.footerSource) ?? kind) as BandContainer;
  const inBand = () => {
    const selection = bridge!.getSelection();
    const band = ((headerFooterDocument().section as Record<string, unknown>)[source] ?? []) as Array<{ id?: string }>;
    return Boolean(selection && band.some((block) => block.id === selection.focus.blockId));
  };
  // From the other band, the first double-click only closes it (as a person's would).
  for (let attempt = 0; attempt < 2 && !inBand(); attempt += 1) {
    app.dispatchEvent(new MouseEvent("mousedown", init));
    window.dispatchEvent(new MouseEvent("mouseup", { ...init, buttons: 0 }));
  }
  const target = bandCaretTarget(headerFooterDocument(), source);
  if (!inBand()) {
    notify(`Double-click the ${kind === "header" ? "top" : "bottom"} margin of a page to edit its ${kind}.`);
    return false;
  }
  if (target) bridge!.setSelection({ anchor: target, focus: target });
  bridge!.focus();
  notify(`Editing the ${kind}. Press Esc to go back to the document.`);
  return true;
}

function removeHeaderFooter(kind: BandKind) {
  if (!canChangeDocument(`remove the ${kind}`)) return false;
  const removed = dispatchModelOps((doc) => removeBandOps(doc as unknown as HeaderFooterDocument, kind));
  notify(removed ? `${kind === "header" ? "Header" : "Footer"} removed.` : `This document has no ${kind}.`);
  focusEditor();
  return removed;
}

function toggleDifferentFirstPage() {
  if (!canChangeDocument("change the first page")) return false;
  const enable = !hasDifferentFirstPage(headerFooterDocument());
  const changed = dispatchModelOps((doc) => differentFirstPageOps(doc as unknown as HeaderFooterDocument, enable));
  if (changed) notify(enable ? "The first page now has its own header and footer." : "The first page uses the same header and footer as the others.");
  focusEditor();
  return changed;
}

function openBandMenu(kind: BandKind) {
  const anchor = ribbonButton(`simple.header-footer.${kind}`);
  if (!anchor || !handle || !canChangeDocument(`edit the ${kind}`)) return;
  const doc = headerFooterDocument();
  openPopupMenu(anchor, [
    { label: `Edit ${kind}`, run: () => void editHeaderFooter(kind) },
    { label: `Remove ${kind}`, disabled: !hasBand(doc, kind), run: () => removeHeaderFooter(kind) },
    "separator",
    { label: "Different first page", checked: hasDifferentFirstPage(doc), hint: "Give the first page its own header and footer (for a title page)", run: () => toggleDifferentFirstPage() },
  ], { label: kind === "header" ? "Header" : "Footer", onClose: (chosen) => { if (!chosen) focusEditor(); } });
}

async function openPageNumbers() {
  if (!canChangeDocument("insert page numbers")) return false;
  closeEditorOverlays();
  const doc = headerFooterDocument();
  const returnFocus = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  const result = await openPageNumberDialog({
    current: currentPageNumberOptions(doc),
    hasPageNumbers: documentHasPageNumbers(doc),
    icon: icons.pageNumber,
    onClosed: () => window.setTimeout(() => restoreFocus(returnFocus), 0),
  });
  if (!result || !handle) return false;
  const changed = dispatchModelOps((current) => result.action === "remove"
    ? removePageNumberOps(current as unknown as HeaderFooterDocument)
    : insertPageNumberOps(current as unknown as HeaderFooterDocument, result.options));
  if (result.action === "remove") notify(changed ? "Page numbers removed." : "This document has no page numbers.");
  else notify(changed ? `Page numbers are in the ${result.options.position === "top" ? "header" : "footer"}.` : "The page numbers already look like that.");
  window.setTimeout(() => focusEditor(), 0);
  return changed;
}

// ---- The font box ------------------------------------------------------------------

/** Families the editor has loaded (lower case): the startup set, plus any loaded later. */
const readyFonts = new Set(installedFonts.map((font) => familyKey(font.family)));
let installedFamilies: string[] = [];
let installedFamiliesLoad: Promise<void> | null = null;
let builtinFonts: Array<{ family: string; label: string }> = [];
const fontNotesShown = new Set<string>();

function fontView(): FontCatalogView {
  return { installed: installedFamilies, ready: readyFonts, builtin: builtinFonts };
}

function loadInstalledFonts(): Promise<void> {
  installedFamiliesLoad ||= Promise.resolve(fontBridge.getFontFamilies?.()).then((list) => {
    installedFamilies = (list ?? []).map((entry) => normalizeFamily(entry?.family)).filter(Boolean);
    fontPicker?.refresh();
  }).catch((error) => {
    console.warn("The installed fonts could not be listed", error);
    installedFamiliesLoad = null;
  });
  return installedFamiliesLoad;
}

/**
 * Loads installed fonts into the running editor when the engine can (a future engine
 * hook, feature-detected); otherwise they load the next time a window starts.
 */
async function loadFontsNow(families: readonly string[]): Promise<boolean> {
  const register = (handle as unknown as { registerFonts?: (fonts: CustomFontDef[]) => unknown } | null)?.registerFonts;
  if (typeof register !== "function" || !families.length) return false;
  const wanted = new Set(families.map(familyKey));
  const fonts = (await fontBridge.getDocumentFonts({ families: [...families] })).filter((font) => wanted.has(familyKey(font.family)) && !readyFonts.has(familyKey(font.family)));
  if (!fonts.length) return false;
  await register.call(handle, fonts);
  for (const font of fonts) readyFonts.add(familyKey(font.family));
  return true;
}

async function applyFont(family: string, entry: FontEntry): Promise<boolean> {
  if (!canChangeDocument("change the font")) return false;
  const live = wordEditor();
  if (!live?.setCharStyle) return false;
  let shown = entry.state !== "pending";
  if (!shown) shown = await loadFontsNow([family]).catch(() => false);
  live.setCharStyle({ fontFamily: family });
  rememberRecentFont(localSettings(), family);
  if (!shown && !fontNotesShown.has(familyKey(family))) {
    fontNotesShown.add(familyKey(family));
    notify(`“${family}” is applied and saved with the document. This window shows it with a similar font; new Simple Docs windows show the font itself.`);
  }
  return true;
}

/** Installed fonts the open document uses that this window could not load: loaded next time. */
async function noteDocumentFonts() {
  if (!handle) return;
  const unloaded = documentFontFamilies(handle.getDocument()).filter((family) => !readyFonts.has(familyKey(family)));
  if (!unloaded.length) return;
  await loadInstalledFonts();
  const installed = unloaded.filter((family) => installedFamilies.some((name) => familyKey(name) === familyKey(family)));
  if (!installed.length) return;
  rememberDocumentFonts(localSettings(), installed);
  if (await loadFontsNow(installed).catch(() => false)) fontPicker?.refresh();
}

function mountFontPicker() {
  const select = editorHost.querySelector<HTMLSelectElement>(`select[data-ribbon-item="home.font.font-family"]`)
    ?? editorHost.querySelector<HTMLElement>(`[data-ribbon-item="home.font.font-family"]`)?.querySelector("select")
    ?? null;
  if (!select) return;
  // The editor's own list offers its built-in look-alikes as CSS stacks ("Cambria, serif").
  builtinFonts = [...select.options].filter((option) => option.value.includes(",")).map((option) => ({ family: normalizeFamily(option.value), label: option.textContent?.trim() || normalizeFamily(option.value) }));
  fontPicker = createFontPicker({
    select,
    view: fontView,
    loadInstalled: loadInstalledFonts,
    documentFamilies: () => (handle ? documentFontFamilies(handle.getDocument()) : []),
    recentFamilies: () => readRecentFonts(localSettings()),
    apply: applyFont,
    onDone: () => {
      focusEditor();
    },
  });
  showDocumentFont();
}

/** The font box shows the caret's font, or the body text font when there is no caret yet. */
function showDocumentFont() {
  if (!handle || !fontPicker) return;
  const live = wordEditor() as { currentFormat?: () => { fontFamily?: string | null } } | null;
  const atCaret = bridge?.getSelection() ? live?.currentFormat?.()?.fontFamily : null;
  fontPicker.setCurrent(atCaret || bandCharStyle(headerFooterDocument()).fontFamily);
}

/** Starts spelling, AutoFormat and the font box once the editor is ready. */
function initializeWordFeatures() {
  if (!handle || !bridge) return;
  autoFormat = installAutoFormat(bridge, { locale: navigator.language || "en-US", enabled: autoFormatEnabled });
  spelling = createSpelling({
    bridge,
    handle,
    host: editorHost,
    api: (window.simpleDocs as unknown as { spell?: Parameters<typeof createSpelling>[0]["api"] }).spell,
    storage: localSettings(),
    notify,
    focusEditor,
    replaceText: replaceDocumentText,
    language: navigator.language || "en-US",
    onStateChange: refreshProofingButtons,
  });
  mountFontPicker();
  for (const id of ["simple.header-footer.header", "simple.header-footer.footer", "simple.header-footer.page-number", "simple.proofing.spelling", "simple.proofing.autoformat"]) {
    const button = ribbonButton(id);
    button?.classList.add("rib-big");
    if (id.startsWith("simple.header-footer.header") || id.startsWith("simple.header-footer.footer")) button?.setAttribute("aria-haspopup", "menu");
  }
  refreshProofingButtons();
  // The Windows spell checker starts in the background, so the first underlines come quickly.
  if (spelling.enabled) window.setTimeout(() => void (window.simpleDocs as unknown as { spell?: { checkWords?: (words: string[], language?: string) => Promise<unknown> } }).spell?.checkWords?.(["Simple"], spelling?.language).catch(() => {}), 1500);
  window.setTimeout(() => void loadInstalledFonts(), 2500);
}

async function initialize() {
  handle = await editor.whenReady();
  // Blank documents offer Heading 3 too (Ctrl+Alt+3 and the engine's "+" menu use it).
  const blank = DocumentBuilder.create();
  blank.style(headingStyleDefinition(3, { styles: [{ id: "Heading2", name: "Heading 2", char: {}, para: {} }], defaultStyleId: "Normal" }));
  handle.setDocument(
    blank
      .paragraph("", { fontFamily: "Calibri", fontSizePx: 14.667, color: "#111111" })
      .spacing({ lineHeight: 1.15, after: 10.667 })
      .build(),
  );
  bridge = createEngineBridge(editor, handle);
  // Loads (openDocx/setDocument) are not edits: each open adopts its own model.
  bridge.onDocChange((change) => {
    if (change.origin !== "load") syncDirty(change.origin);
    refreshHistoryButtons(change);
    // Any other edit ends the paste the chip could change.
    if (!applyingPasteOption && lastPaste && change.revision !== lastPaste.revision) hidePasteOptions();
    // A new document is checked from scratch; edits re-check only what changed.
    if (change.origin === "load") {
      spelling?.documentReplaced();
      showDocumentFont();
      void noteDocumentFonts();
    } else {
      spelling?.schedule();
    }
  });
  bridge.onReviewChange(() => {
    syncDirty("review");
    refreshHistoryButtons();
  });
  editor.on("modeChanged", (event) => {
    refreshHistoryButtons();
    noteSuggestingAuthor(event.mode);
    // Viewing hides the spelling marks; Editing and Suggesting show them again.
    spelling?.schedule(0);
  });
  commands = createEditorCommands({
    bridge,
    handle,
    notify,
    focusEditor,
    editorShown,
    author: {
      name: () => authorDisplayName(author),
      chosen: () => authorChosen,
      choose: (name) => setAuthorName(name, true),
    },
    icons: { link: icons.link, comment: icons.comment, bookmark: icons.bookmark, list: icons.list },
  });
  // The engine's hyperlink, bookmark and drop-down prompts open Simple's own small dialogs.
  bridge.setDialogHandler(async (request) => {
    if (request.kind === "alert") {
      notify(request.message, request.id === "document.open-failed" ? "error" : "normal");
      return;
    }
    if (request.kind === "confirm") {
      const choice = await askChoice({ title: request.title || "Simple Docs", message: request.message, choices: [{ id: "cancel", label: "Cancel" }, { id: "ok", label: "OK", primary: true }] });
      return choice === "ok";
    }
    return commands ? commands.answerEnginePrompt(request) : null;
  });
  initializeWordFeatures();
  annotateRibbonShortcuts();
  $("author-name-label").textContent = authorDisplayName(author);
  void adoptWindowsUserName();
  loadingOverlay.hidden = true;
  setTitle();
  const welcomeLists = refreshWelcomeLists().catch((error) => console.warn("Recent files could not be listed", error));
  // A file opened from Explorer or the command line reports its own errors
  // and never turns into a startup failure.
  if (pendingExternalPath) {
    const filePath = pendingExternalPath;
    pendingExternalPath = null;
    await openWith(() => window.simpleDocs.openPath(filePath));
  }
  await welcomeLists;
}

void initialize().catch((error) => {
  console.error("Simple Docs failed to initialize", error);
  loaderTitle.textContent = "Simple Docs could not start";
  loaderDetail.textContent = error instanceof Error ? error.message : "The page layout engine failed to load.";
});

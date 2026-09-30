export function documentExportDisabled(documentOpen, busy) {
  return !documentOpen || busy;
}

export function activeModal(root) {
  return root.querySelector(".modal-backdrop:not([hidden])");
}

export function isModalBlockedShortcut(event) {
  const modifier = Boolean(event.ctrlKey || event.metaKey);
  if (!modifier) return false;
  const key = String(event.key || "").toLowerCase();
  return ["s", "o", "n", "p"].includes(key) || (Boolean(event.shiftKey) && key === "e");
}

export function nextPrintPreviewGeneration(current) {
  return Number.isSafeInteger(current) && current >= 0 && current < Number.MAX_SAFE_INTEGER
    ? current + 1
    : 1;
}

export function canCommitPrintPreview(requestGeneration, currentGeneration, modalHidden) {
  return Number.isSafeInteger(requestGeneration)
    && requestGeneration === currentGeneration
    && modalHidden === false;
}

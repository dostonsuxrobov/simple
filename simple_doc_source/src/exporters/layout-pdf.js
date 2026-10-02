// The Office page view (a LibreOffice-rendered PDF of the unchanged file) is
// optional. Its conversion can time out on a cold start or reject a file. The
// failure must not make printing and PDF export of the unchanged document fail
// for the rest of the session: they fall back to the editor's own rendering.

/**
 * Wrap the page-view PDF request. `promise` feeds the page view; `pdfOr(render)`
 * returns the page-view PDF when it succeeded and otherwise the editor render.
 * A failure is remembered, so later calls go straight to the editor render.
 */
export function createLayoutPdfSource(load) {
  let failed = false;
  let error = null;
  const promise = Promise.resolve().then(load);
  promise.catch((reason) => {
    failed = true;
    error = reason;
  });
  return {
    promise,
    get failed() {
      return failed;
    },
    get error() {
      return error;
    },
    async pdfOr(render) {
      if (!failed) {
        try {
          return await promise;
        } catch (reason) {
          failed = true;
          error = reason;
        }
      }
      return render();
    },
  };
}

export const LAYOUT_FALLBACK_NOTE = "The Office page view could not be prepared. Printing and PDF export use the editor’s own layout.";

const test = require("node:test");
const assert = require("node:assert/strict");
const { simpleHooks } = require("../scripts/patch-wordcanvas.cjs");

const load = () => import("../src/engine-bridge.ts");
const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function fakeEditor() {
  const handlers = new Map();
  return {
    on(event, handler) {
      if (!handlers.has(event)) handlers.set(event, new Set());
      handlers.get(event).add(handler);
      return () => handlers.get(event)?.delete(handler);
    },
    emit(event, data) {
      for (const handler of [...(handlers.get(event) ?? [])]) handler(data);
    },
    count(event) {
      return handlers.get(event)?.size ?? 0;
    },
  };
}

function hookedHandle() {
  const calls = [];
  const record = (name, result) => (...args) => {
    calls.push([name, ...args]);
    return typeof result === "function" ? result(...args) : result;
  };
  return {
    calls,
    handle: {
      simpleHooks: 1,
      getDocument: () => ({ blocks: [] }),
      getSelection: () => ({ anchor: { blockId: "p", offset: 1 }, focus: { blockId: "p", offset: 1 } }),
      getModelRevision: () => 7,
      insertImageBytes: record("insertImageBytes", async () => true),
      insertBlocks: record("insertBlocks", true),
      replaceBlock: record("replaceBlock", true),
      replaceImage: record("replaceImage", async () => true),
      setDialogHandler: record("setDialogHandler"),
      setInsertTextHook: record("setInsertTextHook"),
      setBuiltinAutoCorrect: record("setBuiltinAutoCorrect"),
      deleteWord: record("deleteWord", true),
      undo: record("undo"),
      redo: record("redo"),
      canUndo: record("canUndo", true),
      canRedo: record("canRedo", false),
      setSelection: record("setSelection"),
      focus: record("focus"),
      seedReview: record("seedReview"),
      positionFromPoint: record("positionFromPoint", { blockId: "p", offset: 3 }),
      setDecorations: record("setDecorations"),
      clearDecorations: record("clearDecorations"),
      invalidateDecorations: record("invalidateDecorations"),
    },
  };
}

function captureConsoleError() {
  const original = console.error;
  const messages = [];
  console.error = (...args) => messages.push(args.map(String).join(" "));
  return { messages, restore: () => { console.error = original; } };
}

test("the bridge only relies on handle methods and events the SIMPLE_HOOKS patch defines", async () => {
  const { DOC_CHANGE_EVENT, SIMPLE_HOOKS_VERSION } = await load();
  const handleEdit = simpleHooks.edits.find(([, after]) => after.includes("simpleHooks: simpleHookVersion"));
  assert.ok(handleEdit, "patch extends the handle");
  for (const name of ["getModelRevision", "insertImageBytes", "insertBlocks", "replaceBlock", "replaceImage", "setDialogHandler", "setInsertTextHook", "setBuiltinAutoCorrect", "deleteWord", "undo", "redo", "canUndo", "canRedo", "setSelection", "focus", "seedReview", "positionFromPoint"]) {
    assert.match(handleEdit[1], new RegExp(`\\n    ${name}: `), `handle defines ${name}`);
  }
  assert.match(simpleHooks.helpers, new RegExp(`name: "${DOC_CHANGE_EVENT}"`));
  assert.match(simpleHooks.helpers, new RegExp(`simpleHookVersion = ${SIMPLE_HOOKS_VERSION},`));
  const dialogIds = [...simpleHooks.edits.map(([, after]) => after).join("\n").matchAll(/kind: "(?:prompt|alert)", id: "([^"]+)"/g)].map((match) => match[1]).sort();
  assert.deepEqual(dialogIds, ["bookmark.add", "bookmark.rename", "content-control.dropdown-items", "content-control.none", "document.open-failed", "hyperlink.edit", "hyperlink.insert"]);
});

test("onDocChange delivers only valid simple:docchange payloads and isolates listener failures", async () => {
  const { createEngineBridge } = await load();
  const editor = fakeEditor();
  const { handle } = hookedHandle();
  const warnings = [];
  const bridge = createEngineBridge(editor, handle, { warn: (message) => warnings.push(message) });
  assert.equal(bridge.hooked, true);
  const seen = [];
  const errors = captureConsoleError();
  try {
    bridge.onDocChange(() => { throw new Error("listener bug"); });
    const stop = bridge.onDocChange((change) => seen.push(change));
    const payload = { revision: 3, origin: "typing", canUndo: true, canRedo: false };
    editor.emit("custom", { name: "simple:docchange", payload });
    editor.emit("custom", { name: "other", payload });
    editor.emit("custom", { name: "simple:docchange", payload: { revision: "x" } });
    editor.emit("custom", { name: "simple:docchange" });
    assert.deepEqual(seen, [payload]);
    assert.equal(errors.messages.length, 1, "a throwing listener is reported, not propagated");
    stop();
    stop();
    editor.emit("custom", { name: "simple:docchange", payload: { ...payload, revision: 4 } });
    assert.equal(seen.length, 1, "unsubscribed listener stays silent");
  } finally {
    errors.restore();
  }
  assert.equal(bridge.revision(), 7);
  assert.deepEqual(warnings, []);
  bridge.dispose();
  assert.equal(editor.count("custom"), 0, "dispose removes every subscription");
});

test("onReviewChange forwards the review layer", async () => {
  const { createEngineBridge } = await load();
  const editor = fakeEditor();
  const bridge = createEngineBridge(editor, hookedHandle().handle, { warn: () => {} });
  const layers = [];
  bridge.onReviewChange((review) => layers.push(review));
  const review = { docId: "d", baseHead: 0, suggestions: [], threads: [] };
  editor.emit("reviewChanged", { review });
  assert.deepEqual(layers, [review]);
  bridge.dispose();
  editor.emit("reviewChanged", { review });
  assert.equal(layers.length, 1);
});

test("hooked handle: every operation forwards to the patched handle", async () => {
  const { createEngineBridge } = await load();
  const { handle, calls } = hookedHandle();
  const bridge = createEngineBridge(fakeEditor(), handle, { warn: () => { throw new Error("no warnings expected"); } });
  const bytes = new Uint8Array([1, 2, 3]);
  assert.equal(await bridge.insertImageBytes(bytes, "image/png", { maxWidthPx: 300 }), true);
  assert.equal(bridge.insertBlocks([{ kind: "paragraph" }], { index: 0 }), true);
  assert.equal(bridge.replaceBlock("b1", (current) => current), true);
  assert.equal(await bridge.replaceImage("img", bytes, "image/jpeg", { fit: "frame" }), true);
  assert.equal(bridge.deleteWord(-5), true);
  assert.equal(bridge.deleteWord(1), true);
  const dialog = async () => null;
  const hook = () => undefined;
  assert.equal(bridge.setDialogHandler(dialog), true);
  assert.equal(bridge.setInsertTextHook(hook), true);
  assert.equal(bridge.setBuiltinAutoCorrect(false), true);
  assert.equal(bridge.undo(), true);
  assert.equal(bridge.redo(), true);
  assert.equal(bridge.canUndo(), true);
  assert.equal(bridge.canRedo(), false);
  const selection = { anchor: { blockId: "p", offset: 0 }, focus: { blockId: "p", offset: 2 } };
  assert.equal(bridge.setSelection(selection), true);
  assert.deepEqual(bridge.getSelection(), { anchor: { blockId: "p", offset: 1 }, focus: { blockId: "p", offset: 1 } });
  assert.equal(bridge.focus(), true);
  assert.equal(bridge.seedReview({ threads: [] }), true);
  assert.deepEqual(bridge.positionFromPoint(10, 20), { blockId: "p", offset: 3 });
  const underline = { type: "underline", range: selection, color: "#d93025" };
  assert.equal(bridge.setDecorations([underline]), true);
  assert.equal(bridge.invalidateDecorations(), true);
  assert.equal(bridge.clearDecorations(), true);
  assert.deepEqual(calls.map(([name]) => name), ["insertImageBytes", "insertBlocks", "replaceBlock", "replaceImage", "deleteWord", "deleteWord", "setDialogHandler", "setInsertTextHook", "setBuiltinAutoCorrect", "undo", "redo", "canUndo", "canRedo", "setSelection", "focus", "seedReview", "positionFromPoint", "setDecorations", "invalidateDecorations", "clearDecorations"]);
  assert.deepEqual(calls[0].slice(1), [bytes, "image/png", { maxWidthPx: 300 }]);
  assert.deepEqual(calls[4].slice(1), [-1], "direction is normalised");
  assert.equal(calls[6][1], dialog);
  bridge.dispose();
  assert.deepEqual(calls.slice(-2).map((call) => call.slice(0, 2)), [["setDialogHandler", null], ["setInsertTextHook", null]], "dispose removes what this bridge installed");
});

test("engine failures surface as false/null instead of exceptions", async () => {
  const { createEngineBridge } = await load();
  const { handle } = hookedHandle();
  handle.insertImageBytes = async () => { throw new Error("decode"); };
  handle.replaceBlock = () => { throw new Error("bad block"); };
  handle.positionFromPoint = () => { throw new Error("no layout"); };
  const errors = captureConsoleError();
  try {
    const bridge = createEngineBridge(fakeEditor(), handle, { warn: () => {} });
    assert.equal(await bridge.insertImageBytes(new Uint8Array([1]), "image/png"), false);
    assert.equal(bridge.replaceBlock("x", {}), false);
    assert.equal(bridge.positionFromPoint(1, 2), null);
    assert.equal(errors.messages.length, 3);
  } finally {
    errors.restore();
  }
});

test("unpatched engine: warns once per feature and degrades safely", async () => {
  const { createEngineBridge } = await load();
  const runtimeCalls = [];
  const runtimeEditor = {
    undo: () => runtimeCalls.push("undo"),
    redo: () => runtimeCalls.push("redo"),
    setSelection: (selection) => runtimeCalls.push(["setSelection", selection]),
    focus: () => runtimeCalls.push("focus"),
    seedReview: () => runtimeCalls.push("seedReview"),
  };
  const decorations = [];
  // WordCanvas 0.12.0 without the patch: the runtime handle has decoration methods only.
  const handle = { getDocument: () => ({}), getSelection: () => null, setDecorations: (list) => decorations.push(list), clearDecorations: () => decorations.push("clear"), invalidateDecorations: () => decorations.push("invalidate") };
  const warnings = [];
  const bridge = createEngineBridge(fakeEditor(), handle, { warn: (message) => warnings.push(message), runtime: () => ({ editor: runtimeEditor }) });
  assert.equal(bridge.hooked, false);
  assert.equal(await bridge.insertImageBytes(new Uint8Array([1]), "image/png"), false);
  assert.equal(await bridge.insertImageBytes(new Uint8Array([1]), "image/png"), false);
  assert.equal(bridge.insertBlocks([]), false);
  assert.equal(bridge.replaceBlock("a", {}), false);
  assert.equal(await bridge.replaceImage("a", new Uint8Array([1]), "image/png"), false);
  assert.equal(bridge.setDialogHandler(async () => null), false);
  assert.equal(bridge.setInsertTextHook(() => undefined), false);
  assert.equal(bridge.setBuiltinAutoCorrect(false), false);
  assert.equal(bridge.deleteWord(-1), false);
  assert.equal(bridge.canUndo(), false);
  assert.equal(bridge.canRedo(), false);
  assert.equal(bridge.positionFromPoint(1, 1), null);
  assert.equal(bridge.undo(), true);
  assert.equal(bridge.redo(), true);
  assert.equal(bridge.setSelection(null), true);
  assert.equal(bridge.focus(), true);
  assert.equal(bridge.seedReview({}), true);
  assert.equal(bridge.setDecorations([]), true);
  assert.equal(bridge.invalidateDecorations(), true);
  assert.equal(bridge.clearDecorations(), true);
  assert.deepEqual(runtimeCalls, ["undo", "redo", ["setSelection", null], "focus", "seedReview"]);
  assert.deepEqual(decorations, [[], "invalidate", "clear"]);
  const features = warnings.map((message) => message.match(/\] (\S+) is unavailable/)?.[1]);
  assert.deepEqual(features, ["insertImageBytes", "insertBlocks", "replaceBlock", "replaceImage", "setDialogHandler", "setInsertTextHook", "setBuiltinAutoCorrect", "deleteWord", "canUndo", "canRedo", "positionFromPoint"]);
  assert.ok(warnings.every((message) => message.includes("SIMPLE_HOOKS")));
  bridge.dispose();
  const bare = createEngineBridge(fakeEditor(), { getDocument: () => ({}), getSelection: () => null }, { warn: (message) => warnings.push(message), runtime: () => undefined });
  assert.equal(bare.undo(), false);
  assert.equal(bare.setDecorations([]), false);
  assert.match(warnings.at(-1), /setDecorations is unavailable/);
});

test("unpatched engine: document changes are detected by polling the immutable model", async () => {
  const { createEngineBridge } = await load();
  let documentRef = { blocks: [] };
  const warnings = [];
  const bridge = createEngineBridge(fakeEditor(), { getDocument: () => documentRef, getSelection: () => null }, { warn: (message) => warnings.push(message), pollMs: 50 });
  const changes = [];
  const stop = bridge.onDocChange((change) => changes.push(change));
  await pause(130);
  assert.deepEqual(changes, [], "an unchanged reference is not a change");
  documentRef = { blocks: [] };
  await pause(130);
  assert.deepEqual(changes, [{ revision: 1, origin: "unknown", canUndo: false, canRedo: false }]);
  assert.equal(bridge.revision(), 1);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0], /simple:docchange is unavailable/);
  stop();
  documentRef = { blocks: [] };
  await pause(130);
  assert.equal(changes.length, 1, "stopping the last listener stops polling");
  bridge.dispose();
});

test("ribbon items: click the button or a split button's primary half, never disabled ones", async () => {
  const { createEngineBridge, RIBBON_ITEM_IDS } = await load();
  const clicks = [];
  const button = (title, extra = {}) => ({ tagName: "BUTTON", title, disabled: false, getAttribute: (name) => extra[name] ?? null, click: () => clicks.push(title), querySelector: () => null, ...extra });
  const primary = button("Bulleted list");
  const items = {
    "home.font.bold": button("Bold"),
    "home.paragraph.bulleted-list": { tagName: "DIV", getAttribute: () => null, querySelector: (selector) => selector === ":scope > button:first-child" ? primary : null },
    "home.font.font-size": { tagName: "DIV", getAttribute: () => null, querySelector: () => null },
    "home.font.superscript": button("Superscript", { disabled: true }),
    "home.font.subscript": button("Subscript", { "aria-disabled": "true" }),
    'odd"id': button("Odd"),
  };
  const selectors = [];
  const root = {
    querySelector(selector) {
      selectors.push(selector);
      const match = selector.match(/^\[data-ribbon-item="((?:[^"\\]|\\.)*)"\]$/);
      return match ? items[match[1].replace(/\\(.)/g, "$1")] ?? null : null;
    },
    querySelectorAll: () => Object.keys(items).map((id) => ({ getAttribute: (name) => (name === "data-ribbon-item" ? id : null) })),
  };
  const warnings = [];
  const bridge = createEngineBridge(fakeEditor(), hookedHandle().handle, { root, warn: (message) => warnings.push(message) });
  assert.equal(bridge.clickRibbonItem(RIBBON_ITEM_IDS.bold), true);
  assert.equal(bridge.clickRibbonItem(RIBBON_ITEM_IDS.bulletedList), true);
  assert.equal(bridge.clickRibbonItem("home.font.font-size"), false, "no primary button");
  assert.equal(bridge.clickRibbonItem(RIBBON_ITEM_IDS.superscript), false, "disabled");
  assert.equal(bridge.clickRibbonItem(RIBBON_ITEM_IDS.subscript), false, "aria-disabled");
  assert.equal(bridge.clickRibbonItem('odd"id'), true, "attribute values are escaped");
  assert.ok(selectors.includes('[data-ribbon-item="odd\\"id"]'));
  assert.deepEqual(clicks, ["Bold", "Bulleted list", "Odd"]);
  assert.equal(bridge.clickRibbonItem("file.undo.undo"), false);
  assert.equal(bridge.clickRibbonItem("file.undo.undo"), false);
  assert.equal(warnings.length, 1, "a missing item warns once");
  assert.equal(bridge.hasRibbonItem("home.font.bold"), true);
  assert.equal(bridge.hasRibbonItem("missing"), false);
  assert.deepEqual(bridge.ribbonItemIds(), Object.keys(items));
});

test("ribbon id constants are unique and well formed", async () => {
  const { RIBBON_ITEM_IDS } = await load();
  const ids = Object.values(RIBBON_ITEM_IDS);
  assert.equal(new Set(ids).size, ids.length);
  for (const id of ids) assert.match(id, /^(home|insert|layout|table|view)\.[a-z0-9-]+\.[a-z0-9-]+$/);
  assert.ok(!ids.some((id) => id.startsWith("file.")), "Simple removes the File tab");
});

test("isHookedHandle feature-detects the patch version", async () => {
  const { isHookedHandle } = await load();
  assert.equal(isHookedHandle({ simpleHooks: 1 }), true);
  assert.equal(isHookedHandle({ simpleHooks: 2 }), true);
  assert.equal(isHookedHandle({ simpleHooks: "1" }), false);
  assert.equal(isHookedHandle({}), false);
});

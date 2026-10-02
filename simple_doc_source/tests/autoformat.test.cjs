const test = require("node:test");
const assert = require("node:assert/strict");

const load = () => import("../src/autoformat.ts");

/** The hook context right after `typed` was inserted after `before` (and before `after`). */
function context(before, typed, extra = {}) {
  const { after = "", ...rest } = extra;
  return { kind: "text", text: typed, blockId: "p", offset: before.length + typed.length, paragraphText: before + typed + after, paragraphStyle: {}, mode: "edit", ...rest };
}

/**
 * A tiny model of the engine: every keystroke is one undo step, and the hook's edits are
 * applied as a separate step after it (what the SIMPLE_HOOKS patch does).
 */
async function editor(options = {}) {
  const { autoFormat, applyTextEdits } = await load();
  const state = { text: "", list: null };
  const undoStack = [];
  const snapshot = () => ({ text: state.text, list: state.list });
  return {
    state,
    type(keys) {
      for (const key of keys) {
        undoStack.push(snapshot());
        state.text += key;
        const edits = autoFormat({ kind: "text", text: key, blockId: "p", offset: state.text.length, paragraphText: state.text, paragraphStyle: state.list ? { list: { listId: "1", level: 0 } } : {}, mode: "edit" }, options);
        if (edits.length) {
          undoStack.push(snapshot());
          state.text = applyTextEdits(state.text, edits);
          for (const edit of edits) if (edit.type === "list") state.list = edit.numberFormat ?? edit.kind;
        }
      }
      return this;
    },
    undo() {
      Object.assign(state, undoStack.pop());
      return this;
    },
  };
}

const typed = async (keys, options) => (await editor(options)).type(keys).state.text;

test("smart quotes follow the locale and the position", async () => {
  assert.equal(await typed(`He said "hi" and 'yes' (don't) "ok."`), "He said “hi” and ‘yes’ (don’t) “ok.”");
  assert.equal(await typed(`"(nested 'one')"`), "“(Nested ‘one’)”");
  assert.equal(await typed(`Er sagte "Hallo" und 'ja', geht's`, { locale: "de-DE" }), "Er sagte „Hallo“ und ‚ja‘, geht’s");
  assert.equal(await typed(`Il a dit "bonjour"`, { locale: "fr-FR" }), "Il a dit «bonjour»");
  assert.equal(await typed(`Он сказал "привет"`, { locale: "ru-RU" }), "Он сказал «привет»");
  assert.equal(await typed(`"Cześć"`, { locale: "pl" }), "„Cześć”");
  assert.equal(await typed(`"konnichiwa"`, { locale: "ja-JP" }), "「konnichiwa」");
  // Uzbek Latin: oʻ and gʻ take the turned comma, other apostrophes the modifier apostrophe.
  assert.equal(await typed(`O'zbekiston g'alaba ma'lumot "salom"`, { locale: "uz-Latn-UZ" }), "Oʻzbekiston gʻalaba maʼlumot “salom”");
  assert.equal(await typed(`"салом"`, { locale: "uz-Cyrl" }), "«салом»");
  assert.equal(await typed(`Say "hi"`, { smartQuotes: false }), `Say "hi"`);
  const { quoteStyle } = await load();
  assert.equal(quoteStyle("pt-BR").open, "“");
  assert.equal(quoteStyle("pt-PT").open, "«");
  assert.equal(quoteStyle("xx").open, "“");
});

test("-- and --- become en and em dashes once the next character is typed", async () => {
  assert.equal(await typed("Pages 10--12"), "Pages 10–12");
  assert.equal(await typed("Wait---what"), "Wait—what");
  assert.equal(await typed("Word -- word"), "Word – word");
  assert.equal(await typed("Cats - dogs"), "Cats – dogs");
  // Left alone: four hyphens, a leading run, HTML comments and URLs.
  assert.equal(await typed("a----b"), "a----b");
  assert.equal(await typed("--flag"), "--flag");
  assert.equal(await typed("<!--x"), "<!--x");
  assert.equal(await typed("http://a--b.example"), "http://a--b.example");
  assert.equal(await typed("Well-known x-ray"), "Well-known x-ray");
  assert.equal(await typed("a--b", { dashes: false }), "a--b");
});

test("(c) (r) (tm), arrows and ellipsis become symbols", async () => {
  assert.equal(await typed("Acme (c) 2026 Brand(R) Name(tm)"), "Acme © 2026 Brand® Name™");
  assert.equal(await typed("A -> b --> c <-- d ==> e <== f <=> g <-> h"), "A → b → c ← d ⇒ e ⇐ f ⇔ g ↔ h");
  assert.equal(await typed("Wait... what"), "Wait… what");
  assert.equal(await typed("(c)", { symbols: false }), "(c)");
});

test("fractions and English ordinals", async () => {
  const { autoFormat } = await load();
  assert.equal(await typed("Add 1/2 cup and 3/4, then 1/4."), "Add ½ cup and ¾, then ¼.");
  assert.equal(await typed("On 11/2 and 1/2/2026 "), "On 11/2 and 1/2/2026 ");
  assert.deepEqual(autoFormat(context("the 21st", " ")), [{ type: "format", blockId: "p", start: 6, end: 8, style: { verticalAlign: "super" } }]);
  assert.deepEqual(autoFormat(context("the 112th", " ")), [{ type: "format", blockId: "p", start: 7, end: 9, style: { verticalAlign: "super" } }]);
  assert.deepEqual(autoFormat(context("the 1th", " ")), []);
  assert.deepEqual(autoFormat(context("the 21st", " "), { locale: "de-DE" }), []);
  assert.deepEqual(autoFormat(context("add 1/2", " "), { fractions: false }), []);
});

test("list markers at the start of a paragraph start lists", async () => {
  const { autoFormat } = await load();
  assert.deepEqual(autoFormat(context("1.", " ")), [
    { type: "replace", blockId: "p", start: 0, end: 3, text: "" },
    { type: "list", blockId: "p", kind: "number", numberFormat: "decimal", marker: "1." },
  ]);
  assert.deepEqual(autoFormat(context("a)", " ")).at(-1), { type: "list", blockId: "p", kind: "number", numberFormat: "lowerLetter", marker: "a)" });
  assert.deepEqual(autoFormat(context("A.", "\t")).at(-1), { type: "list", blockId: "p", kind: "number", numberFormat: "upperLetter", marker: "A." });
  for (const marker of ["-", "*", "•"]) assert.deepEqual(autoFormat(context(marker, " ")).at(-1), { type: "list", blockId: "p", kind: "bullet", marker });
  // Text already after the caret moves into the new list item.
  assert.deepEqual(autoFormat(context("-", " ", { after: "existing" }))[0], { type: "replace", blockId: "p", start: 0, end: 2, text: "" });
  // Not a list start.
  assert.deepEqual(autoFormat(context("2.", " ")), []);
  assert.deepEqual(autoFormat(context("x 1.", " ")), []);
  assert.deepEqual(autoFormat(context("1.", " ", { paragraphStyle: { list: { listId: "1", level: 0 } } })), []);
  assert.deepEqual(autoFormat(context("1.", " "), { lists: false }), []);
  const page = await editor();
  page.type("1. ");
  assert.equal(page.state.text, "");
  assert.equal(page.state.list, "decimal");
});

test("URLs and e-mail addresses are linked on space, Tab and Enter", async () => {
  const { autoFormat, linkTarget } = await load();
  assert.deepEqual(autoFormat(context("see www.example.com", " ")), [{ type: "format", blockId: "p", start: 4, end: 19, style: { link: "https://www.example.com" } }]);
  assert.deepEqual(autoFormat(context("(at https://example.org/a?b=1).", " ")), [{ type: "format", blockId: "p", start: 4, end: 29, style: { link: "https://example.org/a?b=1" } }]);
  assert.deepEqual(autoFormat(context("Wiki https://en.wikipedia.org/wiki/Foo_(bar)", "\t"))[0].style, { link: "https://en.wikipedia.org/wiki/Foo_(bar)" });
  assert.deepEqual(autoFormat(context("write to me@example.co.uk,", " ")), [{ type: "format", blockId: "p", start: 9, end: 25, style: { link: "mailto:me@example.co.uk" } }]);
  // Enter: the link goes into the paragraph that was split.
  assert.deepEqual(autoFormat({ kind: "paragraph", text: "\n", blockId: "p2", offset: 0, paragraphText: "", paragraphStyle: {}, mode: "edit", previousBlockId: "p1", previousText: "Visit https://simple.local" }), [
    { type: "format", blockId: "p1", start: 6, end: 26, style: { link: "https://simple.local" } },
  ]);
  assert.deepEqual(autoFormat(context("not.a.link", " ")), []);
  assert.deepEqual(autoFormat(context("see www.example.com", " "), { links: false }), []);
  assert.equal(linkTarget("ftp://files.example.com/x"), "ftp://files.example.com/x");
  assert.equal(linkTarget("mailto:a@b.cd"), "mailto:a@b.cd");
  assert.equal(linkTarget("example"), null);
});

test("sentences start with a capital letter", async () => {
  assert.equal(await typed("hello world. this is it! and so? yes "), "Hello world. This is it! And so? Yes ");
  assert.equal(await typed("see e.g. this and Mr. smith or U.S. economy at 5 p.m. today "), "See e.g. this and Mr. smith or U.S. economy at 5 p.m. today ");
  assert.equal(await typed("well... maybe "), "Well… maybe ");
  assert.equal(await typed("so i think i'm right "), "So I think I’m right ");
  assert.equal(await typed("THe IDs of iPhone owners "), "The IDs of iPhone owners ");
  assert.equal(await typed("he said “yes.” then "), "He said “yes.” Then ");
  assert.equal(await typed("istanbul ", { locale: "tr-TR" }), "İstanbul ");
  assert.equal(await typed("привет. мир ", { locale: "ru-RU" }), "Привет. Мир ");
  assert.equal(await typed("hello world ", { capitalizeSentences: false }), "hello world ");
  const { autoFormat } = await load();
  // Enter after a one-word paragraph capitalizes it.
  assert.deepEqual(autoFormat({ kind: "paragraph", text: "\n", blockId: "p2", offset: 0, paragraphText: "", paragraphStyle: {}, mode: "edit", previousBlockId: "p1", previousText: "hello" }), [
    { type: "replace", blockId: "p1", start: 0, end: 1, text: "H" },
  ]);
  // Numbers, addresses and words without case are left alone.
  assert.deepEqual(autoFormat(context("3rd", " ", { }), { ordinals: false }), []);
  assert.deepEqual(autoFormat(context("你好", " ")), []);
});

test("every AutoFormat change is one undo step that restores what was typed", async () => {
  const cases = [
    ["Acme (c)", "Acme ©", "Acme (c)"],
    ["A->", "A→", "A->"],
    ["So x--y", "So x–y", "So x--y"],
    [`So "`, "So “", `So "`],
    ["hello ", "Hello ", "hello "],
    ["1/2 ", "½ ", "1/2 "],
    ["Wait...", "Wait…", "Wait..."],
  ];
  for (const [keys, formatted, restored] of cases) {
    const page = await editor();
    page.type(keys);
    assert.equal(page.state.text, formatted, keys);
    page.undo();
    assert.equal(page.state.text, restored, `Ctrl+Z after ${keys}`);
    // Typing on does not redo the correction.
    page.type(" z");
    assert.equal(page.state.text, `${restored} z`, `typing after undo of ${keys}`);
    // A second Ctrl+Z removes the typed characters, as usual.
    page.undo().undo().undo();
    assert.equal(page.state.text, restored.slice(0, -1), `second undo after ${keys}`);
  }
  // A list start undoes to the typed marker.
  const list = await editor();
  list.type("- ");
  assert.equal(list.state.list, "bullet");
  list.undo();
  assert.deepEqual(list.state, { text: "- ", list: null });
  list.type("x");
  assert.deepEqual(list.state, { text: "- x", list: null });
});

test("edits of one keystroke can be inverted exactly", async () => {
  const { autoFormat, applyTextEdits, invertTextEdits } = await load();
  const ctx = context("wait--", " ");
  const edits = autoFormat(ctx);
  assert.deepEqual(edits.map((edit) => edit.type), ["replace", "replace"]);
  const after = applyTextEdits(ctx.paragraphText, edits);
  assert.equal(after, "Wait– ");
  assert.equal(applyTextEdits(after, invertTextEdits(ctx.paragraphText, edits)), ctx.paragraphText);
});

test("AutoFormat stays out of code, Viewing mode and composed input", async () => {
  const { autoFormat, autoFormatChanges } = await load();
  assert.deepEqual(autoFormat(context("x--y", " ", { paragraphStyle: { namedStyle: "HTMLCode" } })), []);
  assert.deepEqual(autoFormat(context("hello", " ", { mode: "view" })), []);
  assert.deepEqual(autoFormat(context("hello", "  ")), []);
  assert.deepEqual(autoFormat({ ...context("hello", " "), offset: 3 }), []);
  assert.deepEqual(autoFormat({ kind: "paragraph", text: "\n", blockId: "p2", offset: 0, paragraphText: "", paragraphStyle: {}, mode: "edit" }), []);
  // Suggesting mode is fine: the engine records the corrections as suggestions.
  assert.equal(autoFormat(context("hello", " ", { mode: "suggest" })).length, 1);
  const [change] = autoFormatChanges(context("Brand(tm", ")"));
  assert.deepEqual({ rule: change.rule, typed: change.typed, result: change.result }, { rule: "symbols", typed: "(tm)", result: "™" });
});

test("installAutoFormat owns the hook and switches the engine's own conversions off", async () => {
  const { installAutoFormat, createAutoFormatHook } = await load();
  const calls = [];
  let hook = null;
  const bridge = {
    setInsertTextHook(next) { calls.push(["hook", next === null ? null : "fn"]); hook = next; return true; },
    setBuiltinAutoCorrect(enabled) { calls.push(["builtin", enabled]); return true; },
  };
  const changes = [];
  const installed = installAutoFormat(bridge, { onChange: (change) => changes.push(change.rule) });
  assert.equal(installed.active, true);
  assert.deepEqual(calls, [["hook", "fn"], ["builtin", false]]);
  assert.deepEqual(hook(context("Acme (c", ")")), [{ type: "replace", blockId: "p", start: 5, end: 8, text: "©" }]);
  assert.deepEqual(changes, ["symbols"]);
  assert.equal(hook(context("plain", "x")), null);
  installed.setOptions({ symbols: false });
  assert.equal(hook(context("Acme (c", ")")), null);
  installed.dispose();
  installed.dispose();
  assert.deepEqual(calls.slice(2), [["hook", null], ["builtin", true]]);

  // Without the engine hook nothing is switched off.
  const plain = [];
  const inactive = installAutoFormat({ setInsertTextHook: () => false, setBuiltinAutoCorrect: (value) => plain.push(value) });
  assert.equal(inactive.active, false);
  inactive.dispose();
  assert.deepEqual(plain, []);

  // A throwing listener never blocks typing.
  const safe = createAutoFormatHook({ onChange: () => { throw new Error("listener"); } });
  assert.equal(safe(context("Acme (c", ")")).length, 1);
});

test("the single AutoFormat switch turns every rule off without giving typing back to the engine's own conversions", async () => {
  const { autoFormat, installAutoFormat } = await load();
  assert.deepEqual(autoFormat(context("", '"'), { enabled: false }), []);
  assert.deepEqual(autoFormat(context("1.", " "), { enabled: false }), []);
  assert.equal(autoFormat(context("", '"'), { enabled: true }).length, 1);
  const calls = [];
  let hook = null;
  const installed = installAutoFormat({
    setInsertTextHook(next) { hook = next; calls.push(["hook", next === null ? null : "fn"]); return true; },
    setBuiltinAutoCorrect(enabled) { calls.push(["builtin", enabled]); return true; },
  }, { locale: "en-US" });
  assert.deepEqual(hook(context("1.", " ")).map((edit) => edit.type), ["replace", "list"]);
  installed.setOptions({ locale: "en-US", enabled: false });
  assert.equal(hook(context("1.", " ")), null, "off: '1. ' stays text");
  assert.equal(hook(context("say ", '"')), null, "off: straight quotes stay straight");
  assert.deepEqual(calls, [["hook", "fn"], ["builtin", false]], "turning AutoFormat off keeps the engine's built-in conversions off too");
  installed.setOptions({ locale: "en-US", enabled: true });
  assert.deepEqual(hook(context("say ", '"')), [{ type: "replace", blockId: "p", start: 4, end: 5, text: "“" }]);
});

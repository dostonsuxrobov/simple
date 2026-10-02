const test = require("node:test");
const assert = require("node:assert/strict");
const { EventEmitter } = require("node:events");
const { PassThrough } = require("node:stream");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { spawn, spawnSync } = require("node:child_process");
const spelling = require("../electron/spelling.cjs");

const load = () => import("../src/spellcheck.ts");

const words = (tokens) => tokens.map((token) => token.word);

function paragraph(id, text, extra = {}) {
  return { kind: "paragraph", id, revision: 1, runs: [{ text, style: {} }], style: { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0, ...extra } };
}

/** A dictionary that knows a few words and records every request. */
function fakeBackend(known = ["the", "a", "is", "this", "receive", "well", "known", "don't", "test", "word", "words", "simple", "and", "of"]) {
  const dictionary = new Set(known);
  const calls = [];
  const suggestionCalls = [];
  return {
    calls,
    suggestionCalls,
    async checkWords(list, language) {
      calls.push({ words: [...list], language });
      return { source: "windows", language, misspelled: list.map((word) => !dictionary.has(word.toLowerCase())) };
    },
    async getSuggestions(word) {
      suggestionCalls.push(word);
      return word.toLowerCase() === "recieve" ? ["receive", "relieve", "receive", "Recieve"] : word === "teh" ? ["the", "ten", "tech"] : [];
    },
  };
}

test("tokenizer keeps apostrophes inside words and splits hyphenated compounds", async () => {
  const { tokenizeSpelling } = await load();
  const text = "Don't stop the well-knwon students' rock’n’roll";
  const tokens = tokenizeSpelling(text);
  assert.deepEqual(words(tokens), ["Don't", "stop", "the", "well", "knwon", "students", "rock’n’roll"]);
  const knwon = tokens.find((token) => token.word === "knwon");
  assert.equal(text.slice(knwon.start, knwon.end), "knwon");
  assert.equal(knwon.compound, "well-knwon");
  assert.equal(tokens.find((token) => token.word === "stop").compound, undefined);
  for (const token of tokens) assert.equal(text.slice(token.start, token.end), token.word);
});

test("tokenizer skips numbers, single letters, UPPERCASE and inner capitals", async () => {
  const { tokenizeSpelling } = await load();
  assert.deepEqual(words(tokenizeSpelling("Page A4 has 2nd v8_x items, 3.5 kg, x y NASA NASA's iPhone JavaScript Wordd")), ["Page", "has", "items", "kg", "Wordd"]);
  // Each check can be switched off.
  assert.deepEqual(words(tokenizeSpelling("NASA wrd4 iPhone", { ignoreUppercase: false, ignoreWordsWithNumbers: false, ignoreMixedCase: false })), ["NASA", "wrd4", "iPhone"]);
});

test("tokenizer ignores Internet and file addresses", async () => {
  const { tokenizeSpelling, addressRanges } = await load();
  const text = "Mail jhon.doe@exmaple.com or see https://exmaple.org/pathh?q=wrod, www.exmaple.net, C:\\Usres\\notse.docx, \\\\srver\\shaer\\fiel, reprot.pdf, exmaple.io/x #hashtg @mentoin end.";
  assert.deepEqual(words(tokenizeSpelling(text)), ["Mail", "or", "see", "end"]);
  const urls = addressRanges("see https://example.org/a).");
  assert.equal("see https://example.org/a).".slice(urls[0].start, urls[0].end), "https://example.org/a");
  // A sentence that lacks a space after a full stop is still checked when the next word is capitalized.
  assert.deepEqual(words(tokenizeSpelling("Frist.Second")), ["Frist", "Second"]);
  assert.deepEqual(words(tokenizeSpelling("https://x.org", { ignoreInternetAndFileAddresses: false })), ["https", "x", "org"].filter((word) => word.length > 1));
});

test("tokenizer handles scripts by language: foreign scripts skipped, look-alike mixes flagged", async () => {
  const { tokenizeSpelling, scriptForLanguage } = await load();
  // English checking: Cyrillic, Greek, Arabic and CJK words are not English typos.
  assert.deepEqual(words(tokenizeSpelling("Hello привет γεια مرحبا 你好 world")), ["Hello", "world"]);
  // Russian checking keeps Cyrillic words and skips Latin ones.
  assert.deepEqual(words(tokenizeSpelling("Привет hello мир", { language: "ru-RU" })), ["Привет", "мир"]);
  // Uzbek Latin oʻ/gʻ (U+02BB) and ʼ (U+02BC) are letters of the word.
  assert.deepEqual(words(tokenizeSpelling("Oʻzbekiston maʼlumot gʻalaba", { language: "uz-Latn-UZ" })), ["Oʻzbekiston", "maʼlumot", "gʻalaba"]);
  assert.equal(scriptForLanguage("uz-Cyrl"), "Cyrillic");
  assert.equal(scriptForLanguage("sr-Latn"), "Latin");
  assert.equal(scriptForLanguage("el"), "Greek");
  // "сat" with a Cyrillic с is a look-alike typo.
  const mixed = tokenizeSpelling("the сat sat");
  assert.deepEqual(mixed.map((token) => [token.word, Boolean(token.mixedScript)]), [["the", false], ["сat", true], ["sat", false]]);
  assert.deepEqual(words(tokenizeSpelling("the сat", { flagMixedScripts: false })), ["the"]);
});

test("repeated words are found on their second occurrence", async () => {
  const { repeatedWords } = await load();
  const text = "This is the the test. The the end, end. 10 10 and and";
  const issues = repeatedWords(text);
  assert.deepEqual(issues.map((issue) => text.slice(issue.start, issue.end)), ["the", "the", "and"]);
  const first = issues[0];
  assert.equal(text.slice(first.deleteStart, first.deleteEnd), " the");
});

test("document paragraphs include table cells; hidden text and code styles are left out", async () => {
  const { collectSpellParagraphs, paragraphSpellText } = await load();
  const hidden = { kind: "paragraph", id: "h", revision: 1, style: {}, runs: [{ text: "vis", style: {} }, { text: "XX", style: { hidden: true } }, { text: "ible", style: {} }] };
  assert.equal(paragraphSpellText(hidden), "vis\u0000\u0000ible");
  const table = { kind: "table", id: "t", revision: 1, rows: [{ cells: [{ id: "c1", blocks: [paragraph("p2", "in cell")] }, { id: "c2", blocks: [{ kind: "table", id: "t2", revision: 1, rows: [{ cells: [{ id: "c3", blocks: [paragraph("p3", "nested")] }] }] }] }] }] };
  const doc = { blocks: [paragraph("p1", "body"), { kind: "image", id: "i", revision: 1 }, table, paragraph("code", "x = frobnicate()", { namedStyle: "HTMLCode" }), hidden] };
  const paragraphs = collectSpellParagraphs(doc);
  assert.deepEqual(paragraphs.map((item) => [item.id, Boolean(item.skip)]), [["p1", false], ["p2", false], ["p3", false], ["code", true], ["h", false]]);
});

test("checker batches unknown words, caches results and reports decoration ranges", async () => {
  const { SpellChecker } = await load();
  const backend = fakeBackend();
  const checker = new SpellChecker(backend, { batchSize: 2 });
  const doc = { blocks: [paragraph("p1", "This is a tset of the wrods"), paragraph("p2", "Recieve the wrods")] };
  const issues = await checker.checkDocument(doc);
  assert.deepEqual(issues.map((issue) => [issue.blockId, issue.word, issue.start, issue.end, issue.kind]), [
    ["p1", "tset", 10, 14, "spelling"],
    ["p1", "wrods", 22, 27, "spelling"],
    ["p2", "Recieve", 0, 7, "spelling"],
    ["p2", "wrods", 12, 17, "spelling"],
  ]);
  // Unique words only, in batches of two.
  const requested = backend.calls.flatMap((call) => call.words);
  assert.equal(new Set(requested).size, requested.length);
  assert.ok(backend.calls.every((call) => call.words.length <= 2 && call.language === "en-US"));
  // A second check with the same text asks nothing.
  const before = backend.calls.length;
  await checker.checkDocument(doc);
  assert.equal(backend.calls.length, before);
  // Only the new word of an edited paragraph is looked up.
  await checker.checkDocument({ blocks: [paragraph("p1", "This is a tset of the wrods tpyo"), paragraph("p2", "Recieve the wrods")] });
  assert.deepEqual(backend.calls.slice(before).flatMap((call) => call.words), ["tpyo"]);

  const decorations = checker.decorations();
  assert.equal(decorations.length, 5);
  assert.deepEqual(decorations[0], { type: "underline", range: { anchor: { blockId: "p1", offset: 10 }, focus: { blockId: "p1", offset: 14 } }, color: "#d93025", thickness: 1.5 });
  let clicked = null;
  const clickable = checker.decorations(undefined, (issue, event) => { clicked = [issue.word, event.clientX]; });
  clickable[1].onClick({ clientX: 7, clientY: 9 });
  assert.deepEqual(clicked, ["wrods", 7]);
});

test("concurrent checks share in-flight lookups", async () => {
  const { SpellChecker } = await load();
  const backend = fakeBackend();
  const checker = new SpellChecker(backend);
  const doc = { blocks: [paragraph("p", "tset wrods")] };
  const [first, second] = await Promise.all([checker.checkDocument(doc), checker.checkDocument(doc)]);
  assert.equal(first.length, 2);
  assert.equal(second.length, 2);
  assert.equal(backend.calls.length, 1);
});

test("the word at the caret is not flagged while it is being typed", async () => {
  const { SpellChecker } = await load();
  const checker = new SpellChecker(fakeBackend());
  const doc = { blocks: [paragraph("p", "tset wrod")] };
  const issues = await checker.checkDocument(doc, { caret: { blockId: "p", offset: 9 } });
  assert.deepEqual(issues.map((issue) => issue.word), ["tset"]);
});

test("user dictionary follows Word's case rules; Ignore all accepts the exact word", async () => {
  const { SpellChecker } = await load();
  const checker = new SpellChecker(fakeBackend());
  checker.setUserDictionary({ words: ["simplex", "Tashkent"], ignored: ["Qwerty"] });
  assert.equal(checker.isAccepted("simplex"), true);
  assert.equal(checker.isAccepted("Simplex"), true);
  assert.equal(checker.isAccepted("SIMPLEX"), true);
  assert.equal(checker.isAccepted("sImplex"), false);
  assert.equal(checker.isAccepted("Tashkent"), true);
  assert.equal(checker.isAccepted("TASHKENT"), true);
  assert.equal(checker.isAccepted("tashkent"), false);
  assert.equal(checker.isAccepted("Qwerty"), true);
  assert.equal(checker.isAccepted("qwerty"), false);
  checker.addToDictionary("dont’ve");
  assert.equal(checker.isAccepted("dont've"), true);
  checker.ignoreAll("zorp", { persist: false });
  checker.addToDictionary("ad-hok");
  const issues = await checker.checkParagraphs([{ id: "p", text: "simplex Tashkent tashkent zorp Qwerty blarg ad-hok hok" }]);
  assert.deepEqual(issues.map((issue) => issue.word), ["tashkent", "blarg", "hok"]);
});

test("context menu offers case-matched suggestions, Add to dictionary and Ignore all", async () => {
  const { SpellChecker } = await load();
  const backend = fakeBackend();
  const checker = new SpellChecker(backend, { maxSuggestions: 2 });
  await checker.checkParagraphs([{ id: "p", text: "Recieve the the teh" }]);
  const issue = checker.issueAt({ blockId: "p", offset: 3 });
  assert.equal(issue.word, "Recieve");
  const menu = await checker.menuFor(issue);
  assert.deepEqual(menu.suggestions.map((item) => item.label), ["Receive", "Relieve"]);
  assert.deepEqual(menu.suggestions[0], { label: "Receive", replacement: "Receive", blockId: "p", start: 0, end: 7 });
  assert.equal(menu.canAddToDictionary, true);
  assert.equal(menu.canIgnore, true);
  // Suggestions are cached.
  await checker.menuFor(issue);
  assert.equal(backend.suggestionCalls.length, 1);
  const repeated = checker.issueAt({ blockId: "p", offset: 13 });
  assert.equal(repeated.kind, "repeated");
  const repeatedMenu = await checker.menuFor(repeated);
  assert.deepEqual(repeatedMenu.suggestions, [{ label: "Delete repeated word", replacement: "", blockId: "p", start: 11, end: 15 }]);
  assert.equal(repeatedMenu.canAddToDictionary, false);
  assert.equal(checker.issueAt({ blockId: "p", offset: 9 }), null);
  assert.equal(checker.issueAt({ blockId: "other", offset: 0 }), null);
});

test("an unavailable dictionary flags nothing and is not asked again", async () => {
  const { SpellChecker, preloadSpellBackend } = await load();
  let calls = 0;
  const checker = new SpellChecker({ async checkWords() { calls++; return { source: null, misspelled: [] }; }, async getSuggestions() { return []; } });
  assert.deepEqual(await checker.checkParagraphs([{ id: "p", text: "tset wrods" }]), []);
  assert.equal(checker.available, false);
  await checker.checkParagraphs([{ id: "p", text: "tset wrods more" }]);
  assert.equal(calls, 1);
  // Changing the language tries again.
  checker.setOptions({ language: "en-GB" });
  assert.equal(checker.available, true);

  assert.equal(preloadSpellBackend(undefined), null);
  assert.equal(preloadSpellBackend({ checkWords() {} }), null);
  const backend = preloadSpellBackend({
    async checkWords(list, language) { return { source: "windows", language, misspelled: list.map((word) => word === "tset") }; },
    async getWordSuggestions() { return ["test"]; },
  });
  assert.deepEqual(await backend.checkWords(["tset", "test"], "en-US"), { source: "windows", language: "en-US", misspelled: [true, false] });
  assert.deepEqual(await backend.getSuggestions("tset"), ["test"]);
});

test("cache is bounded and evicts the least recently used lookups", async () => {
  const { SpellChecker } = await load();
  const backend = fakeBackend();
  const checker = new SpellChecker(backend, { cacheSize: 100 });
  const list = Array.from({ length: 150 }, (_, index) => `wrd${String.fromCharCode(97 + (index % 26))}${String.fromCharCode(97 + Math.floor(index / 26))}`);
  await checker.lookup(list);
  const calls = backend.calls.length;
  await checker.lookup(list.slice(-10));
  assert.equal(backend.calls.length, calls);
  await checker.lookup(list.slice(0, 1));
  assert.equal(backend.calls.length, calls + 1);
});

// ---------------------------------------------------------------------------------------
// Main-process side (electron/spelling.cjs)

test("spell words are normalized and validated", () => {
  assert.equal(spelling.normalizeSpellWord("  Tashkent "), "Tashkent");
  assert.equal(spelling.normalizeSpellWord("e\u0301te"), "éte");
  assert.equal(spelling.normalizeSpellWord("two words"), null);
  assert.equal(spelling.normalizeSpellWord("tab\tword"), null);
  assert.equal(spelling.normalizeSpellWord(""), null);
  assert.equal(spelling.normalizeSpellWord("x".repeat(65)), null);
  assert.equal(spelling.normalizeSpellWord(42), null);
});

test("requested languages resolve to an installed Windows checker", () => {
  const supported = ["en-CA", "en-GB", "en-US", "ru-RU", "uz-Latn"];
  assert.equal(spelling.resolveSpellLanguage("en-US", supported), "en-US");
  assert.equal(spelling.resolveSpellLanguage("en_gb", supported), "en-GB");
  assert.equal(spelling.resolveSpellLanguage("en", supported), "en-US");
  assert.equal(spelling.resolveSpellLanguage("en-AU", supported), "en-US");
  assert.equal(spelling.resolveSpellLanguage("en-CA", supported), "en-CA");
  assert.equal(spelling.resolveSpellLanguage("ru", supported), "ru-RU");
  assert.equal(spelling.resolveSpellLanguage("uz-Latn-UZ", supported), "uz-Latn");
  assert.equal(spelling.resolveSpellLanguage("de-DE", supported), null);
  assert.equal(spelling.resolveSpellLanguage("en-US", null), null);
  assert.equal(spelling.canonicalTag("uz-latn-uz"), "uz-Latn-UZ");
});

test("user dictionary updates are validated, sorted and serialized", async () => {
  let state = spelling.updateDictionary(null, "add", "Zeta");
  state = spelling.updateDictionary(state, "add", "alpha");
  state = spelling.updateDictionary(state, "ignore", "qwrt");
  assert.deepEqual(state, { version: 1, words: ["alpha", "Zeta"], ignored: ["qwrt"] });
  assert.deepEqual(spelling.updateDictionary(state, "remove", "Zeta").words, ["alpha"]);
  assert.deepEqual(spelling.updateDictionary(state, "unignore", "qwrt").ignored, []);
  assert.deepEqual(spelling.updateDictionary({ ignored: ["word"] }, "add", "word"), { version: 1, words: ["word"], ignored: [] });
  assert.throws(() => spelling.updateDictionary(state, "add", "two words"));
  assert.throws(() => spelling.updateDictionary(state, "explode", "word"));
  assert.deepEqual(spelling.sanitizeDictionary({ words: ["ok", 5, "bad word", "ok"], ignored: "nope" }), { version: 1, words: ["ok"], ignored: [] });

  let stored = { words: ["kept"] };
  const writes = [];
  const dictionary = spelling.createUserDictionary({
    read: async () => stored,
    write: async (value) => { await new Promise((resolve) => setTimeout(resolve, 5)); writes.push(value); stored = value; },
  });
  await Promise.all(["one", "two", "three"].map((word) => dictionary.apply("add", word)));
  assert.deepEqual((await dictionary.get()).words, ["kept", "one", "three", "two"]);
  assert.equal(writes.length, 3);
  await assert.rejects(dictionary.apply("add", " "));
  assert.deepEqual((await dictionary.get()).words, ["kept", "one", "three", "two"]);
});

function fakeHostProcess(responder, { ready = "ready\ten-US\ten-GB\tru-RU" } = {}) {
  const spawned = [];
  const spawn = (command, args, options) => {
    const child = new EventEmitter();
    child.stdin = new PassThrough();
    child.stdout = new PassThrough();
    child.killed = false;
    child.kill = () => { child.killed = true; setImmediate(() => child.emit("exit", 1)); };
    spawned.push({ command, args, options, child, lines: [] });
    const record = spawned[spawned.length - 1];
    let buffer = "";
    child.stdin.on("data", (chunk) => {
      buffer += chunk.toString("utf8");
      let newline;
      while ((newline = buffer.indexOf("\n")) >= 0) {
        const line = buffer.slice(0, newline);
        buffer = buffer.slice(newline + 1);
        record.lines.push(line);
        const reply = responder(line.split("\t"));
        if (reply !== undefined) child.stdout.write(`${reply}\n`);
      }
    });
    child.stdin.on("finish", () => setImmediate(() => child.emit("exit", 0)));
    if (ready) setImmediate(() => child.stdout.write(`${ready}\r\n`));
    return child;
  };
  return { spawn, spawned };
}

test("Windows spell host speaks a line protocol to one hidden helper", async () => {
  const fake = fakeHostProcess(([id, op, language, ...rest]) => {
    if (op === "check") return `${id}\tok\t${rest.map((word) => (word === "recieve" ? "1" : "0")).join("")}`;
    if (op === "suggest") return `${id}\tok\treceive\trelieve`;
    return undefined;
  });
  const host = spelling.createWindowsSpellHost({ spawn: fake.spawn, platform: "win32", command: "powershell-test.exe" });
  assert.deepEqual(await host.languages(), ["en-US", "en-GB", "ru-RU"]);
  assert.deepEqual(await host.check(["recieve", "receive", "tab\tbed"], "en"), { language: "en-US", misspelled: [true, false, false] });
  assert.deepEqual(await host.suggest("recieve", "en-GB"), { language: "en-GB", suggestions: ["receive", "relieve"] });
  assert.equal(await host.check(["x"], "de-DE"), null);
  assert.equal(fake.spawned.length, 1);
  const [{ command, args, options, lines }] = fake.spawned;
  assert.equal(command, "powershell-test.exe");
  assert.equal(options.windowsHide, true);
  assert.ok(args.includes("-EncodedCommand") && args.includes("-NonInteractive") && args.includes("-NoProfile"));
  assert.match(Buffer.from(args[args.length - 1], "base64").toString("utf16le"), /ISpellCheckerFactory/);
  assert.equal(lines[0], "1\tcheck\ten-US\trecieve\treceive\ttab bed");
  host.dispose();
  assert.equal(host.available, false);
  assert.equal(await host.check(["recieve"], "en-US"), null);
});

test("Windows spell host gives up cleanly when the helper fails", async () => {
  // Not Windows: never spawns.
  let spawned = 0;
  const other = spelling.createWindowsSpellHost({ spawn: () => { spawned++; }, platform: "linux" });
  assert.equal(await other.check(["word"], "en-US"), null);
  assert.equal(spawned, 0);
  // PowerShell reports a fatal error (for example Add-Type blocked by policy).
  const fatal = fakeHostProcess(() => undefined, { ready: "fatal\tAdd-Type is blocked" });
  const blocked = spelling.createWindowsSpellHost({ spawn: fatal.spawn, platform: "win32" });
  assert.equal(await blocked.check(["word"], "en-US"), null);
  assert.equal(blocked.available, false);
  // Unsupported languages and request timeouts resolve to null.
  const slow = fakeHostProcess(([id, op]) => (op === "check" ? `${id}\tunsupported` : undefined));
  const host = spelling.createWindowsSpellHost({ spawn: slow.spawn, platform: "win32", requestTimeoutMs: 30, maxFailures: 5 });
  assert.equal(await host.check(["word"], "en-US"), null);
  assert.equal(await host.suggest("word", "en-US"), null);
  host.dispose();
  // A helper that never says ready is abandoned after the start timeout.
  const silent = fakeHostProcess(() => undefined, { ready: null });
  const quiet = spelling.createWindowsSpellHost({ spawn: silent.spawn, platform: "win32", startTimeoutMs: 30 });
  assert.equal(await quiet.check(["word"], "en-US"), null);
  assert.equal(silent.spawned[0].child.killed, true);
  quiet.dispose();
});

test("spell IPC validates input, persists the dictionary and broadcasts changes", async () => {
  const handlers = new Map();
  const ipcMain = { handle: (channel, handler) => handlers.set(channel, handler) };
  const host = {
    languages: async () => ["en-US", "ru-RU"],
    check: async (list, language) => ({ language, misspelled: list.map((word) => word === "recieve") }),
    suggest: async (word, language) => ({ language, suggestions: ["receive"] }),
  };
  let stored = null;
  const sessionWords = [];
  const broadcasts = [];
  spelling.registerSpellingIpc(ipcMain, {
    host,
    dictionary: spelling.createUserDictionary({ read: async () => stored, write: async (value) => { stored = value; } }),
    session: { addWordToSpellCheckerDictionary: (word) => sessionWords.push(["add", word]), removeWordFromSpellCheckerDictionary: (word) => sessionWords.push(["remove", word]) },
    preferredLanguages: () => ["uz-Latn-UZ", "ru-RU"],
    broadcast: (channel, value) => broadcasts.push([channel, value]),
  });
  const call = (channel, ...args) => handlers.get(channel)({}, ...args);
  assert.deepEqual(await call("spell:languages"), { available: true, languages: ["en-US", "ru-RU"], preferred: "ru-RU" });
  assert.deepEqual(await call("spell:check", { words: ["recieve", "receive"] }), { source: "windows", language: "ru-RU", misspelled: [true, false] });
  assert.deepEqual(await call("spell:check", { words: ["recieve"], language: "en-US" }), { source: "windows", language: "en-US", misspelled: [true] });
  await assert.rejects(call("spell:check", { words: "recieve" }));
  await assert.rejects(call("spell:check", { words: Array(spelling.MAX_WORDS_PER_CHECK + 1).fill("a") }));
  assert.deepEqual(await call("spell:suggest", { word: "recieve", language: "en-US" }), { source: "windows", language: "en-US", suggestions: ["receive"] });
  assert.deepEqual(await call("spell:suggest", { word: "two words" }), { source: null, language: null, suggestions: [] });
  assert.deepEqual(await call("spell:add-word", "Simplex"), { words: ["Simplex"], ignored: [] });
  assert.deepEqual(await call("spell:ignore-word", "zorp"), { words: ["Simplex"], ignored: ["zorp"] });
  assert.deepEqual(await call("spell:remove-word", "Simplex"), { words: [], ignored: ["zorp"] });
  assert.deepEqual(await call("spell:unignore-word", "zorp"), { words: [], ignored: [] });
  assert.deepEqual(await call("spell:dictionary"), { words: [], ignored: [] });
  assert.deepEqual(stored, { version: 1, words: [], ignored: [] });
  assert.deepEqual(sessionWords, [["add", "Simplex"], ["remove", "Simplex"]]);
  assert.equal(broadcasts.length, 4);
  assert.equal(broadcasts[0][0], "spell:dictionary-changed");
  await assert.rejects(call("spell:add-word", "two words"));
});

test("Chromium's dictionary downloader is pointed at a local folder", () => {
  const calls = [];
  const session = { setSpellCheckerEnabled: (value) => calls.push(["enabled", value]), setSpellCheckerDictionaryDownloadURL: (url) => calls.push(["url", url]) };
  const folder = path.join(os.tmpdir(), "Simple Docs", "offline-dictionaries");
  spelling.configureSpellingSession(session, folder);
  assert.deepEqual(calls[0], ["enabled", true]);
  assert.match(calls[1][1], /^file:\/\/\/.*offline-dictionaries\/$/);
  assert.ok(!/^https?:/i.test(calls[1][1]));
});

// ---------------------------------------------------------------------------------------
// Electron check: real preload, real Windows checker, networking blocked. Opt-in because it
// starts Electron and PowerShell: set SIMPLE_ELECTRON_TESTS=1.

const electronCheck = process.platform === "win32" && process.env.SIMPLE_ELECTRON_TESTS === "1";

test("Electron: isWordMisspelled('recieve') is true and suggests 'receive' with networking blocked", { skip: electronCheck ? false : "set SIMPLE_ELECTRON_TESTS=1 on Windows to run", timeout: 120000 }, async () => {
  const electron = require("electron");
  const work = fs.mkdtempSync(path.join(os.tmpdir(), "simple-docs-spell-"));
  const harness = path.join(work, "harness.cjs");
  const netLog = path.join(work, "net-log.json");
  fs.writeFileSync(harness, `
const { app, BrowserWindow, ipcMain, session } = require('electron')
const fs = require('node:fs/promises')
const path = require('node:path')
const spelling = require(${JSON.stringify(path.resolve(__dirname, "../electron/spelling.cjs"))})
const events = []
app.whenReady().then(async () => {
  const ses = session.defaultSession
  spelling.configureSpellingSession(ses, path.join(app.getPath('userData'), 'offline-dictionaries'))
  for (const name of ['spellcheck-dictionary-download-begin', 'spellcheck-dictionary-download-success', 'spellcheck-dictionary-download-failure']) ses.on(name, (_event, code) => events.push(name + ':' + code))
  // Ask for a language Windows cannot check here, which makes Chromium look for a Hunspell file.
  ses.setSpellCheckerLanguages(['en-US', 'hy'])
  const host = spelling.createWindowsSpellHost()
  const file = path.join(app.getPath('userData'), 'spelling-dictionary.json')
  spelling.registerSpellingIpc(ipcMain, {
    host, session: ses, preferredLanguages: () => ['en-US'],
    dictionary: spelling.createUserDictionary({ read: async () => JSON.parse(await fs.readFile(file, 'utf8')), write: (value) => fs.writeFile(file, JSON.stringify(value)) }),
  })
  const win = new BrowserWindow({ show: false, webPreferences: { preload: ${JSON.stringify(path.resolve(__dirname, "../electron/preload.cjs"))}, sandbox: true, contextIsolation: true, nodeIntegration: false } })
  await win.loadURL('data:text/html,<p>spell</p>')
  let result
  try {
    result = await win.webContents.executeJavaScript(\`(async () => {
      const spell = window.simpleDocs.spell
      return {
        misspelled: await spell.isWordMisspelled('recieve'),
        correct: await spell.isWordMisspelled('receive'),
        suggestions: await spell.getWordSuggestions('recieve'),
        batch: await spell.checkWords(['teh', 'the', "don't"], 'en-US'),
        added: await spell.addWord('Simplexity'),
        dictionary: await spell.getUserDictionary(),
        languages: await spell.getLanguages(),
      }
    })()\`)
    result.stored = JSON.parse(await fs.readFile(file, 'utf8'))
  } catch (error) {
    result = { error: String(error && error.stack || error) }
  }
  await new Promise((resolve) => setTimeout(resolve, 1500))
  result.events = events
  process.stdout.write('RESULT ' + JSON.stringify(result) + '\\n')
  host.dispose()
  app.quit()
})
`);
  const child = spawn(electron, [harness, `--user-data-dir=${path.join(work, "profile")}`, `--log-net-log=${netLog}`, "--host-resolver-rules=MAP * ~NOTFOUND"], { windowsHide: true, stdio: ["ignore", "pipe", "pipe"] });
  let output = "";
  child.stdout.on("data", (chunk) => { output += chunk; });
  child.stderr.on("data", () => {});
  const killTree = () => { if (child.exitCode === null) spawnSync("taskkill", ["/pid", String(child.pid), "/T", "/F"], { windowsHide: true }); };
  const timer = setTimeout(killTree, 90000);
  try {
    await new Promise((resolve) => child.on("exit", resolve));
  } finally {
    clearTimeout(timer);
    killTree();
  }
  const line = output.split(/\r?\n/).find((item) => item.startsWith("RESULT "));
  assert.ok(line, `no result from Electron: ${output.slice(0, 2000)}`);
  const result = JSON.parse(line.slice(7));
  assert.equal(result.error, undefined, result.error);
  assert.equal(result.misspelled, true);
  assert.equal(result.correct, false);
  assert.ok(result.suggestions.includes("receive"), JSON.stringify(result.suggestions));
  assert.equal(result.batch.source, "windows");
  assert.deepEqual(result.batch.misspelled, [true, false, false]);
  assert.deepEqual(result.added, { words: ["Simplexity"], ignored: [] });
  assert.deepEqual(result.stored.words, ["Simplexity"]);
  assert.ok(result.languages.available && result.languages.languages.includes("en-US"));
  // Chromium may look for the Hunspell file, but only in the local folder: no network request at all.
  const log = fs.existsSync(netLog) ? fs.readFileSync(netLog, "utf8") : "";
  const urls = [...log.matchAll(/"url":"([^"]*)"/g)].map((match) => match[1]);
  const network = urls.filter((url) => /^(?:https?|wss?|ftp):/i.test(url));
  assert.deepEqual(network, [], `network requests: ${network.join(", ")}`);
  assert.ok(!result.events.includes("spellcheck-dictionary-download-success:hy"), JSON.stringify(result.events));
  if (result.events.includes("spellcheck-dictionary-download-begin:hy")) {
    // The Hunspell lookup happened and stayed on disk.
    assert.ok(urls.some((url) => url.startsWith("file:") && url.includes("offline-dictionaries")), JSON.stringify(urls.slice(0, 20)));
  }
  if (process.env.SIMPLE_SPELL_DEBUG) console.log(JSON.stringify({ events: result.events, urls: [...new Set(urls)] }));
  fs.rmSync(work, { recursive: true, force: true });
});

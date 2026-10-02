'use strict'
/**
 * Local spell checking for Simple Docs.
 *
 * Electron's own checker cannot answer word lookups on Windows: languages Windows supports
 * are checked asynchronously inside Chromium, so webFrame.isWordMisspelled() is always
 * false, and other languages make Chromium download Hunspell dictionaries from Google.
 * Simple therefore asks the Windows Spell Checking API (ISpellCheckerFactory, offline,
 * built into Windows 8+) through one hidden, long-lived PowerShell helper, and points
 * Chromium's dictionary downloader at a local folder so it never goes online.
 *
 * The user dictionary (Add to dictionary, Ignore all) is a JSON file in Simple's user data.
 * Windows' own user dictionary is never written.
 */
const childProcess = require('node:child_process')
const path = require('node:path')
const { pathToFileURL } = require('node:url')

const MAX_WORD_LENGTH = 64
const MAX_WORDS_PER_CHECK = 1000
const MAX_DICTIONARY_WORDS = 20000
const MAX_SUGGESTIONS = 8

// ---------------------------------------------------------------------------------------
// Words and languages

/** A word fit for the dictionary or a lookup: NFC, trimmed, no whitespace or controls, <= 64 chars. */
function normalizeSpellWord(word) {
  if (typeof word !== 'string') return null
  const value = word.normalize('NFC').trim()
  if (!value || value.length > MAX_WORD_LENGTH) return null
  if (/[\s\u0000-\u001f\u007f\u2028\u2029]/u.test(value)) return null
  return value
}

/** Canonical BCP 47 casing: en-us -> en-US, uz-latn-uz -> uz-Latn-UZ. */
function canonicalTag(tag) {
  return String(tag || '').replace(/_/g, '-').split('-').filter(Boolean).map((part, index) => {
    if (index === 0) return part.toLowerCase()
    if (part.length === 4 && /^[a-z]+$/i.test(part)) return part[0].toUpperCase() + part.slice(1).toLowerCase()
    if (part.length === 2 || /^\d{3}$/.test(part)) return part.toUpperCase()
    return part.toLowerCase()
  }).join('-')
}

/**
 * The supported checker language for a requested tag: exact match, then the tag with
 * fewer subtags (uz-Latn-UZ -> uz-Latn -> uz), then any region of the same language
 * (preferring the requested region, then en-US style defaults). Null when none fits.
 */
function resolveSpellLanguage(requested, supported) {
  const list = (supported || []).map(canonicalTag).filter(Boolean)
  if (!list.length) return null
  const lower = new Map(list.map((tag) => [tag.toLowerCase(), tag]))
  const parts = canonicalTag(requested).split('-').filter(Boolean)
  for (let length = parts.length; length > 0; length--) {
    const hit = lower.get(parts.slice(0, length).join('-').toLowerCase())
    if (hit) return hit
  }
  if (!parts.length) return null
  const base = parts[0].toLowerCase()
  const family = list.filter((tag) => tag.split('-')[0].toLowerCase() === base)
  if (!family.length) return null
  const region = parts.find((part, index) => index > 0 && /^[A-Z]{2}$/.test(part))
  const defaults = { en: 'en-US', es: 'es-ES', fr: 'fr-FR', de: 'de-DE', pt: 'pt-BR', zh: 'zh-CN', ru: 'ru-RU', it: 'it-IT' }
  return family.find((tag) => region && tag.endsWith(`-${region}`))
    || family.find((tag) => tag === defaults[base])
    || family[0]
}

// ---------------------------------------------------------------------------------------
// User dictionary

function emptyDictionary() {
  return { version: 1, words: [], ignored: [] }
}

function cleanList(list) {
  const seen = new Set()
  const result = []
  for (const item of Array.isArray(list) ? list : []) {
    const word = normalizeSpellWord(item)
    if (word && !seen.has(word)) { seen.add(word); result.push(word) }
  }
  return result.sort((a, b) => a.localeCompare(b)).slice(0, MAX_DICTIONARY_WORDS)
}

/** A valid dictionary from whatever was stored (bad entries dropped). */
function sanitizeDictionary(value) {
  const source = value && typeof value === 'object' ? value : {}
  return { version: 1, words: cleanList(source.words), ignored: cleanList(source.ignored) }
}

/** The dictionary after one action: 'add' | 'remove' | 'ignore' | 'unignore'. */
function updateDictionary(state, action, word) {
  const current = sanitizeDictionary(state)
  const value = normalizeSpellWord(word)
  if (!value) throw new Error('This word cannot be added to the dictionary.')
  const words = new Set(current.words)
  const ignored = new Set(current.ignored)
  if (action === 'add') { words.add(value); ignored.delete(value) } else if (action === 'remove') words.delete(value)
  else if (action === 'ignore') ignored.add(value)
  else if (action === 'unignore') ignored.delete(value)
  else throw new Error(`Unknown dictionary action: ${action}`)
  if (words.size > MAX_DICTIONARY_WORDS || ignored.size > MAX_DICTIONARY_WORDS) throw new Error('The dictionary is full.')
  return sanitizeDictionary({ words: [...words], ignored: [...ignored] })
}

/** Serialized read-modify-write access to the stored dictionary. */
function createUserDictionary({ read, write }) {
  let queue = Promise.resolve()
  let cached = null
  const load = async () => (cached ||= sanitizeDictionary(await read().catch(() => emptyDictionary())))
  const run = (task) => {
    const next = queue.then(task, task)
    queue = next.catch(() => {})
    return next
  }
  return {
    get: () => run(load),
    apply: (action, word) => run(async () => {
      const next = updateDictionary(await load(), action, word)
      await write(next)
      cached = next
      return next
    }),
  }
}

// ---------------------------------------------------------------------------------------
// Windows Spell Checking API host

const HOST_SOURCE = String.raw`
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
namespace SimpleDocsSpelling {
  [ComImport, Guid("00000101-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IEnumString {
    [PreserveSig] int Next(int celt, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 0), Out] string[] rgelt, IntPtr fetched);
    [PreserveSig] int Skip(int celt);
    void Reset();
    void Clone(out IEnumString clone);
  }
  [ComImport, Guid("B7C82D61-FBE8-4B47-9B27-6C0D2E0DE0A3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface ISpellingError {
    uint StartIndex { get; }
    uint Length { get; }
    int CorrectiveAction { get; }
    string Replacement { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
  }
  [ComImport, Guid("803E3BD4-2828-4410-8290-418D1D73C762"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IEnumSpellingError {
    [PreserveSig] int Next(out ISpellingError value);
  }
  [ComImport, Guid("B6FD0B71-E2BC-4653-8D05-F197E412770B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface ISpellChecker {
    string LanguageTag { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
    IEnumSpellingError Check([MarshalAs(UnmanagedType.LPWStr)] string text);
    IEnumString Suggest([MarshalAs(UnmanagedType.LPWStr)] string word);
  }
  [ComImport, Guid("8E018A9D-2415-4677-BF08-794EA61F94BB"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface ISpellCheckerFactory {
    IEnumString SupportedLanguages { get; }
    int IsSupported([MarshalAs(UnmanagedType.LPWStr)] string languageTag);
    ISpellChecker CreateSpellChecker([MarshalAs(UnmanagedType.LPWStr)] string languageTag);
  }
  [ComImport, Guid("7AB36653-1796-484B-BDFA-E74F1DB7C1DC")]
  public class SpellCheckerFactoryClass {}
  public static class Host {
    static List<string> Read(IEnumString items, int limit) {
      var list = new List<string>();
      var buffer = new string[1];
      while (items != null && list.Count < limit && items.Next(1, buffer, IntPtr.Zero) == 0) list.Add(Clean(buffer[0]));
      return list;
    }
    static string Clean(string value) {
      return (value ?? "").Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');
    }
    static bool Misspelled(ISpellChecker checker, string word) {
      if (word.Length == 0) return false;
      ISpellingError error;
      return checker.Check(word).Next(out error) == 0;
    }
    public static void Run() {
      var input = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
      var output = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false));
      output.AutoFlush = true;
      output.NewLine = "\n";
      ISpellCheckerFactory factory;
      try { factory = (ISpellCheckerFactory)new SpellCheckerFactoryClass(); }
      catch (Exception error) { output.WriteLine("fatal\t" + Clean(error.Message)); return; }
      output.WriteLine("ready\t" + string.Join("\t", Read(factory.SupportedLanguages, 500)));
      var checkers = new Dictionary<string, ISpellChecker>(StringComparer.OrdinalIgnoreCase);
      string line;
      while ((line = input.ReadLine()) != null) {
        var parts = line.Split('\t');
        if (parts.Length < 3) continue;
        string id = parts[0], op = parts[1], language = parts[2];
        if (op == "quit") break;
        try {
          ISpellChecker checker;
          if (!checkers.TryGetValue(language, out checker)) {
            if (factory.IsSupported(language) == 0) { output.WriteLine(id + "\tunsupported"); continue; }
            checker = factory.CreateSpellChecker(language);
            checkers[language] = checker;
          }
          if (op == "check") {
            var bits = new StringBuilder();
            for (int index = 3; index < parts.Length; index++) bits.Append(Misspelled(checker, parts[index]) ? '1' : '0');
            output.WriteLine(id + "\tok\t" + bits);
          } else if (op == "suggest") {
            var list = parts.Length > 3 && parts[3].Length > 0 ? Read(checker.Suggest(parts[3]), ${MAX_SUGGESTIONS}) : new List<string>();
            output.WriteLine(id + "\tok" + (list.Count > 0 ? "\t" + string.Join("\t", list) : ""));
          } else {
            output.WriteLine(id + "\terror\tUnknown request");
          }
        } catch (Exception error) {
          output.WriteLine(id + "\terror\t" + Clean(error.Message));
        }
      }
    }
  }
}
`

const HOST_SCRIPT = [
  "$ErrorActionPreference = 'Stop'",
  'try {',
  `  Add-Type -TypeDefinition @'\n${HOST_SOURCE}\n'@`,
  '} catch {',
  '  [Console]::Out.WriteLine("fatal`t" + ($_.Exception.Message -replace "[`t`r`n]", " "))',
  '  exit 3',
  '}',
  '[SimpleDocsSpelling.Host]::Run()',
].join('\n')

function powershellPath(env = process.env) {
  return path.join(env.SystemRoot || env.windir || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
}

function hostArguments() {
  return ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', Buffer.from(HOST_SCRIPT, 'utf16le').toString('base64')]
}

function cleanField(value) {
  return String(value).replace(/[\t\r\n]/g, ' ')
}

/**
 * The Windows Spell Checking API in a hidden helper process, started on first use and
 * stopped after `idleMs` without requests. Every method resolves to null when no
 * checker is available (not Windows, PowerShell blocked, or the language is missing).
 */
function createWindowsSpellHost(options = {}) {
  const spawn = options.spawn || childProcess.spawn
  const platform = options.platform || process.platform
  const startTimeoutMs = options.startTimeoutMs ?? 20000
  const requestTimeoutMs = options.requestTimeoutMs ?? 10000
  const idleMs = options.idleMs ?? 5 * 60 * 1000
  const maxFailures = options.maxFailures ?? 3
  let child = null
  let starting = null
  let languages = null
  let disabled = platform !== 'win32'
  let failures = 0
  let nextId = 1
  let idleTimer = null
  const pending = new Map()

  const settleAll = () => {
    for (const [id, request] of pending) { clearTimeout(request.timer); request.resolve(null); pending.delete(id) }
  }
  const fail = (fatal) => {
    failures++
    if (fatal || failures >= maxFailures) disabled = true
  }
  const stop = () => {
    clearTimeout(idleTimer)
    idleTimer = null
    const current = child
    child = null
    starting = null
    settleAll()
    if (current) {
      try { current.stdin.end('0\tquit\t-\n') } catch {}
      const killer = setTimeout(() => { try { current.kill() } catch {} }, 1500)
      killer.unref?.()
      current.once('exit', () => clearTimeout(killer))
    }
  }
  const touch = () => {
    clearTimeout(idleTimer)
    idleTimer = setTimeout(() => { if (!pending.size) stop(); else touch() }, idleMs)
    idleTimer.unref?.()
  }

  const start = () => {
    if (disabled) return Promise.resolve(null)
    if (child && languages) return Promise.resolve(languages)
    if (starting) return starting
    starting = new Promise((resolve) => {
      let process_
      try {
        process_ = spawn(options.command || powershellPath(), hostArguments(), { windowsHide: true, stdio: ['pipe', 'pipe', 'ignore'] })
      } catch {
        fail(true)
        starting = null
        resolve(null)
        return
      }
      child = process_
      let buffer = ''
      let ready = false
      const timer = setTimeout(() => {
        if (ready) return
        fail(false)
        try { process_.kill() } catch {}
        resolve(null)
      }, startTimeoutMs)
      timer.unref?.()
      process_.stdout.setEncoding('utf8')
      process_.stdout.on('data', (chunk) => {
        buffer += chunk
        let newline
        while ((newline = buffer.indexOf('\n')) >= 0) {
          const line = buffer.slice(0, newline).replace(/\r$/, '')
          buffer = buffer.slice(newline + 1)
          const fields = line.split('\t')
          if (!ready) {
            if (fields[0] === 'ready') {
              ready = true
              clearTimeout(timer)
              failures = 0
              languages = fields.slice(1).filter(Boolean)
              touch()
              resolve(languages)
            } else if (fields[0] === 'fatal') {
              clearTimeout(timer)
              fail(true)
              resolve(null)
            }
            continue
          }
          const request = pending.get(fields[0])
          if (!request) continue
          pending.delete(fields[0])
          clearTimeout(request.timer)
          request.resolve(fields[1] === 'ok' ? fields.slice(2) : fields[1] === 'unsupported' ? { unsupported: true } : null)
        }
      })
      process_.on('error', () => {
        clearTimeout(timer)
        fail(true)
        if (child === process_) { child = null; starting = null; languages = null }
        settleAll()
        resolve(null)
      })
      process_.on('exit', () => {
        clearTimeout(timer)
        if (!ready) fail(false)
        if (child === process_) { child = null; starting = null; languages = null }
        settleAll()
        resolve(null)
      })
      process_.stdin.on('error', () => {})
    }).then((result) => {
      if (!result) starting = null
      return result
    })
    return starting
  }

  const request = async (op, language, args) => {
    const supported = await start()
    if (!supported || !child) return null
    const id = String(nextId++)
    touch()
    return new Promise((resolve) => {
      const timer = setTimeout(() => {
        pending.delete(id)
        resolve(null)
        // A stuck helper is replaced on the next request.
        fail(false)
        stop()
      }, requestTimeoutMs)
      timer.unref?.()
      pending.set(id, { resolve, timer })
      try {
        child.stdin.write(`${id}\t${op}\t${cleanField(language)}${args.map((arg) => `\t${cleanField(arg)}`).join('')}\n`)
      } catch {
        pending.delete(id)
        clearTimeout(timer)
        resolve(null)
      }
    })
  }

  return {
    /** Languages Windows can check, or null. */
    languages: () => start(),
    /** { language, misspelled: boolean[] } or null. */
    async check(words, language) {
      const supported = await start()
      const resolved = resolveSpellLanguage(language, supported)
      if (!resolved) return null
      const result = await request('check', resolved, words)
      if (!Array.isArray(result)) return null
      const bits = result[0] || ''
      return { language: resolved, misspelled: words.map((_word, index) => bits[index] === '1') }
    },
    /** { language, suggestions: string[] } or null. */
    async suggest(word, language) {
      const supported = await start()
      const resolved = resolveSpellLanguage(language, supported)
      if (!resolved) return null
      const result = await request('suggest', resolved, [word])
      if (!Array.isArray(result)) return null
      return { language: resolved, suggestions: result.filter(Boolean).slice(0, MAX_SUGGESTIONS) }
    },
    get available() { return !disabled },
    get running() { return Boolean(child) },
    dispose() { disabled = true; stop() },
  }
}

// ---------------------------------------------------------------------------------------
// Electron wiring

/**
 * Keeps Chromium's spell checker on (native squiggles in ordinary text fields) and points
 * its Hunspell downloader at a local folder, so no dictionary is ever fetched from the
 * network. Hunspell .bdic files placed in that folder are used as-is.
 */
function configureSpellingSession(session, dictionariesDirectory) {
  session.setSpellCheckerEnabled(true)
  const folder = dictionariesDirectory.endsWith(path.sep) ? dictionariesDirectory : `${dictionariesDirectory}${path.sep}`
  session.setSpellCheckerDictionaryDownloadURL(pathToFileURL(folder).href)
}

function wordList(value) {
  if (!Array.isArray(value)) throw new Error('Words must be a list.')
  if (value.length > MAX_WORDS_PER_CHECK) throw new Error(`At most ${MAX_WORDS_PER_CHECK} words can be checked at once.`)
  return value.map((word) => normalizeSpellWord(word) || '')
}

function languageOf(value, fallback) {
  return typeof value === 'string' && /^[a-z]{2,3}(?:[-_][a-z0-9]{2,8})*$/i.test(value) ? value : fallback()
}

/**
 * Registers the spell:* IPC channels. `preferredLanguages()` lists the user's languages
 * (app.getPreferredSystemLanguages); `broadcast(channel, value)` reaches every window.
 */
function registerSpellingIpc(ipcMain, { host, dictionary, session = null, preferredLanguages = () => [], broadcast = () => {} }) {
  let preferred = null
  const defaultLanguage = async () => {
    if (preferred) return preferred
    const supported = await host.languages()
    for (const tag of [...preferredLanguages(), 'en-US']) {
      const resolved = resolveSpellLanguage(tag, supported)
      if (resolved) return (preferred = resolved)
    }
    return null
  }
  const pick = (value) => languageOf(value, () => null)

  ipcMain.handle('spell:languages', async () => {
    const languages = await host.languages()
    const available = Boolean(languages && languages.length)
    return { available, languages: languages || [], preferred: available ? await defaultLanguage() : null }
  })
  ipcMain.handle('spell:check', async (_event, input = {}) => {
    const words = wordList(input.words)
    const language = pick(input.language) || await defaultLanguage()
    const result = language ? await host.check(words, language) : null
    return result ? { source: 'windows', language: result.language, misspelled: result.misspelled } : { source: null, language: null, misspelled: [] }
  })
  ipcMain.handle('spell:suggest', async (_event, input = {}) => {
    const word = normalizeSpellWord(input.word)
    if (!word) return { source: null, language: null, suggestions: [] }
    const language = pick(input.language) || await defaultLanguage()
    const result = language ? await host.suggest(word, language) : null
    return result ? { source: 'windows', language: result.language, suggestions: result.suggestions } : { source: null, language: null, suggestions: [] }
  })
  ipcMain.handle('spell:dictionary', async () => {
    const { words, ignored } = await dictionary.get()
    return { words, ignored }
  })
  const change = (action) => async (_event, word) => {
    const next = await dictionary.apply(action, word)
    const value = normalizeSpellWord(word)
    try {
      if (action === 'add') session?.addWordToSpellCheckerDictionary(value)
      if (action === 'remove') session?.removeWordFromSpellCheckerDictionary(value)
    } catch {}
    const result = { words: next.words, ignored: next.ignored }
    broadcast('spell:dictionary-changed', result)
    return result
  }
  ipcMain.handle('spell:add-word', change('add'))
  ipcMain.handle('spell:remove-word', change('remove'))
  ipcMain.handle('spell:ignore-word', change('ignore'))
  ipcMain.handle('spell:unignore-word', change('unignore'))
}

module.exports = {
  MAX_WORDS_PER_CHECK,
  canonicalTag,
  configureSpellingSession,
  createUserDictionary,
  createWindowsSpellHost,
  hostArguments,
  normalizeSpellWord,
  powershellPath,
  registerSpellingIpc,
  resolveSpellLanguage,
  sanitizeDictionary,
  updateDictionary,
}

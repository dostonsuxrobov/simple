'use strict'

/**
 * Type-checks the shared renderer TypeScript (simple/shared/renderer/*.ts) the
 * way each workspace will compile its vendored copy in src/simple-io/: with the
 * workspace's own TypeScript (<workspace>/node_modules/typescript), the
 * compiler options of the tsconfig that covers its src/ folder, and the @types
 * packages that configuration loads, on a temporary copy of the files.
 *
 *   node scripts/check-shared-renderer.cjs           check every workspace that receives renderer files
 *   --module=pdf[,calc]                              limit to these workspaces
 *   --target-root <dir>                              folder holding the simple_*_source workspaces (default: repo root)
 *   --shared-root <dir>                              folder holding manifest.json and renderer/ (default: simple/shared)
 *   --temp-root <dir>                                where the temporary copy goes (default: the system temp folder)
 *   --keep                                           keep the temporary copy for debugging
 *
 * Every workspace is checked twice: with its configuration as it is (including
 * @types it picks up automatically, such as @types/node) and with DOM types
 * only (`types: []`), so shared code can't lean on Node globals. One more pass
 * adds the strict flags common in Vite templates (noUnusedLocals,
 * noUnusedParameters, noFallthroughCasesInSwitch, noImplicitReturns,
 * noImplicitOverride, noUncheckedIndexedAccess) with the newest TypeScript
 * found, so a workspace can tighten its tsconfig without breaking shared code.
 *
 * Workspaces are checked whether or not they are enabled in manifest.json, so
 * shared code is known to compile before an integrator enables a workspace.
 * Nothing inside a workspace folder is written.
 * Exit codes: 0 clean, 1 type errors, 2 setup problem (missing TypeScript, unreadable tsconfig, …).
 */

const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { loadManifest } = require('./sync-shared.cjs')

const ROOT = path.resolve(__dirname, '..')
const REPO_ROOT = path.resolve(ROOT, '..')
const DEFAULT_SHARED_ROOT = path.join(ROOT, 'shared')
const DEFAULT_TARGET_ROOT = REPO_ROOT
const VENDORED_FOLDER = path.join('src', 'simple-io')
const STRICT_EXTRAS = Object.freeze({
  noUnusedLocals: true,
  noUnusedParameters: true,
  noFallthroughCasesInSwitch: true,
  noImplicitReturns: true,
  noImplicitOverride: true,
  noUncheckedIndexedAccess: true,
})
// Emit and project options that don't apply to a one-off, no-emit check of copied files.
const DROPPED_OPTIONS = Object.freeze([
  'composite', 'incremental', 'tsBuildInfoFile', 'declaration', 'declarationMap', 'declarationDir',
  'emitDeclarationOnly', 'outDir', 'outFile', 'out', 'rootDir', 'rootDirs', 'sourceMap', 'inlineSourceMap',
  'inlineSources', 'build',
])
// "No inputs were found" is expected for solution-style tsconfig files.
const IGNORED_CONFIG_CODES = new Set([18003])

/** Error for a missing TypeScript, an unreadable tsconfig or a bad option (exit code 2). */
class CheckSetupError extends Error {}

function toPosix(value) {
  return String(value).replaceAll('\\', '/')
}

function compareVersions(a, b) {
  const left = String(a).split(/[.-]/).map((part) => Number.parseInt(part, 10) || 0)
  const right = String(b).split(/[.-]/).map((part) => Number.parseInt(part, 10) || 0)
  for (let index = 0; index < Math.max(left.length, right.length); index += 1) {
    const difference = (left[index] || 0) - (right[index] || 0)
    if (difference) return difference
  }
  return 0
}

/**
 * Loads the TypeScript compiler a workspace installed for itself.
 * @param {string} workspaceRoot
 * @param {string} folder workspace folder name, for messages
 * @returns {typeof import('typescript')}
 * @throws {CheckSetupError} when the workspace has no TypeScript of its own
 */
function loadWorkspaceTypeScript(workspaceRoot, folder) {
  const packageFolder = path.join(workspaceRoot, 'node_modules', 'typescript')
  if (!fs.existsSync(path.join(packageFolder, 'package.json'))) {
    throw new CheckSetupError(`${folder} has no node_modules/typescript. Run npm ci in ${folder} first.`)
  }
  return require(packageFolder)
}

function formatDiagnosticText(ts, diagnostic) {
  return ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n')
}

/**
 * Parses a tsconfig (following "extends") with the workspace's TypeScript.
 * @param {typeof import('typescript')} ts
 * @param {string} configPath
 * @param {string} workspaceRoot
 * @param {string} folder workspace folder name, for messages
 * @returns {import('typescript').ParsedCommandLine}
 */
function parseConfig(ts, configPath, workspaceRoot, folder) {
  const fatal = []
  const host = { ...ts.sys, onUnRecoverableConfigFileDiagnostic: (diagnostic) => fatal.push(diagnostic) }
  const parsed = ts.getParsedCommandLineOfConfigFile(configPath, {}, host)
  const errors = [...fatal, ...((parsed && parsed.errors) || [])].filter((diagnostic) => !IGNORED_CONFIG_CODES.has(diagnostic.code))
  if (!parsed || errors.length) {
    const detail = errors.map((diagnostic) => `TS${diagnostic.code}: ${formatDiagnosticText(ts, diagnostic)}`).join('\n  ')
    throw new CheckSetupError(`${folder}/${toPosix(path.relative(workspaceRoot, configPath))} can't be read by TypeScript ${ts.version}:\n  ${detail}`)
  }
  return parsed
}

function coversSourceFolder(parsed, workspaceRoot) {
  const source = (path.join(workspaceRoot, 'src') + path.sep).toLowerCase()
  return parsed.fileNames.some((file) => path.resolve(file).toLowerCase().startsWith(source))
}

/**
 * Finds the tsconfig that compiles <workspace>/src: tsconfig.json itself, or
 * the first project it references that does (Vite's tsconfig.app.json layout).
 * @returns {{configPath: string, parsed: import('typescript').ParsedCommandLine}}
 */
function rendererConfig(ts, workspaceRoot, folder) {
  const queue = [path.join(workspaceRoot, 'tsconfig.json')]
  const seen = new Set()
  let first = null
  while (queue.length) {
    const configPath = path.resolve(queue.shift())
    if (seen.has(configPath.toLowerCase())) continue
    seen.add(configPath.toLowerCase())
    if (!fs.existsSync(configPath)) continue
    const parsed = parseConfig(ts, configPath, workspaceRoot, folder)
    if (!first) first = { configPath, parsed }
    if (coversSourceFolder(parsed, workspaceRoot)) return { configPath, parsed }
    for (const reference of parsed.projectReferences || []) queue.push(ts.resolveProjectReferencePath(reference))
  }
  if (!first) throw new CheckSetupError(`${folder} has no tsconfig.json.`)
  return first
}

/**
 * Compiler options for one pass: the workspace's options with emit settings
 * removed. `configFilePath` stays, so automatic @types and `types` entries
 * resolve from the workspace folder exactly as in the workspace's own build.
 */
function passOptions(parsed, extra) {
  const options = { ...parsed.options, ...extra, noEmit: true }
  for (const key of DROPPED_OPTIONS) delete options[key]
  return options
}

/**
 * Copies the files a workspace receives into <temp>/<module>/src/simple-io/.
 * @returns {{folder: string, roots: string[], missing: string[]}}
 */
function copyRendererFiles(sharedRoot, sources, destination) {
  fs.mkdirSync(destination, { recursive: true })
  const roots = []
  const missing = []
  for (const relative of sources) {
    const from = path.join(sharedRoot, relative)
    if (!fs.existsSync(from)) {
      missing.push(relative)
      continue
    }
    const to = path.join(destination, path.basename(relative))
    fs.copyFileSync(from, to)
    if (/\.(ts|tsx|mts|cts)$/i.test(to) && !/\.d\.ts$/i.test(to)) roots.push(to)
  }
  return { folder: destination, roots, missing }
}

/**
 * Runs one type-check pass.
 * @returns {{errors: string[], fileCount: number}}
 */
function runPass(ts, rootNames, options, copyFolder, sharedRoot) {
  const program = ts.createProgram({ rootNames, options })
  const diagnostics = ts.getPreEmitDiagnostics(program)
  const copyPrefix = (path.resolve(copyFolder) + path.sep).toLowerCase()
  const errors = []
  for (const diagnostic of diagnostics) {
    if (diagnostic.category !== ts.DiagnosticCategory.Error) continue
    const text = formatDiagnosticText(ts, diagnostic)
    if (!diagnostic.file) {
      errors.push(`TS${diagnostic.code}: ${text}`)
      continue
    }
    const fileName = path.resolve(diagnostic.file.fileName)
    let where = toPosix(fileName)
    if (fileName.toLowerCase().startsWith(copyPrefix)) {
      // Report the shared source, not the temporary copy.
      const shared = path.join(sharedRoot, 'renderer', fileName.slice(copyPrefix.length))
      const relative = path.relative(REPO_ROOT, shared)
      where = toPosix(relative && !relative.startsWith('..') && !path.isAbsolute(relative) ? relative : shared)
    }
    if (typeof diagnostic.start === 'number') {
      const { line, character } = diagnostic.file.getLineAndCharacterOfPosition(diagnostic.start)
      where += `:${line + 1}:${character + 1}`
    }
    errors.push(`${where} TS${diagnostic.code}: ${text}`)
  }
  return { errors, fileCount: program.getSourceFiles().length }
}

/**
 * Type-checks the shared renderer files for each workspace that receives them.
 *
 * @param {object} [options]
 * @param {string[]} [options.modules] workspace names (default: every workspace with renderer files)
 * @param {string} [options.sharedRoot] folder holding manifest.json and renderer/
 * @param {string} [options.targetRoot] folder holding the workspace folders
 * @param {string} [options.tempRoot] where the temporary copy is made
 * @param {boolean} [options.keep] keep the temporary copy
 * @returns {{ok: boolean, results: Array<{module: string, folder: string, typescript: string, config: string,
 *   passes: Array<{name: string, errors: string[]}>}>, warnings: string[], temp: string, message: string}}
 * @throws {CheckSetupError} when a workspace can't be checked at all
 */
function checkSharedRenderer(options = {}) {
  const sharedRoot = path.resolve(options.sharedRoot || DEFAULT_SHARED_ROOT)
  const targetRoot = path.resolve(options.targetRoot || DEFAULT_TARGET_ROOT)
  const manifest = loadManifest(sharedRoot)
  const names = Object.keys(manifest.workspaces).filter((name) => manifest.workspaces[name].renderer.length)
  const modules = options.modules && options.modules.length ? [...new Set(options.modules)] : names
  for (const name of modules) {
    if (!manifest.workspaces[name]) throw new CheckSetupError(`Unknown module "${name}". Known modules: ${Object.keys(manifest.workspaces).join(', ')}.`)
    if (!manifest.workspaces[name].renderer.length) throw new CheckSetupError(`${name} receives no renderer files, so there is nothing to check.`)
  }

  const warnings = []
  const rendererFolder = path.join(sharedRoot, 'renderer')
  const sharedSources = fs.existsSync(rendererFolder)
    ? fs.readdirSync(rendererFolder).filter((file) => /\.ts$/i.test(file) && !/\.d\.ts$/i.test(file)).map((file) => `renderer/${file}`)
    : []

  const tempParent = path.resolve(options.tempRoot || os.tmpdir())
  fs.mkdirSync(tempParent, { recursive: true })
  const temp = fs.mkdtempSync(path.join(tempParent, 'simple-shared-renderer-'))
  const results = []
  let newest = null
  try {
    for (const name of modules) {
      const workspace = manifest.workspaces[name]
      const workspaceRoot = path.join(targetRoot, workspace.folder)
      if (!fs.existsSync(workspaceRoot)) throw new CheckSetupError(`Workspace folder for ${name} is missing: ${workspaceRoot}`)
      for (const source of sharedSources) {
        if (!workspace.renderer.includes(source)) warnings.push(`simple/shared/${source} is not vendored to ${workspace.folder}; add it to manifest.json if that workspace needs it.`)
      }
      const ts = loadWorkspaceTypeScript(workspaceRoot, workspace.folder)
      const { configPath, parsed } = rendererConfig(ts, workspaceRoot, workspace.folder)
      const copy = copyRendererFiles(sharedRoot, workspace.renderer, path.join(temp, name, VENDORED_FOLDER))
      for (const relative of copy.missing) warnings.push(`simple/shared/${relative} is listed for ${name} but does not exist yet.`)
      const passes = []
      if (copy.roots.length) {
        passes.push({ name: 'as configured', ...runPass(ts, copy.roots, passOptions(parsed, {}), copy.folder, sharedRoot) })
        passes.push({ name: 'DOM types only', ...runPass(ts, copy.roots, passOptions(parsed, { types: [] }), copy.folder, sharedRoot) })
      }
      const result = {
        module: name,
        folder: workspace.folder,
        typescript: ts.version,
        config: toPosix(path.relative(workspaceRoot, configPath)),
        passes,
        roots: copy.roots,
        copyFolder: copy.folder,
        parsed,
        ts,
      }
      results.push(result)
      if (copy.roots.length && (!newest || compareVersions(ts.version, newest.typescript) > 0)) newest = result
    }
    if (newest) {
      const strict = runPass(newest.ts, newest.roots, passOptions(newest.parsed, { types: [], ...STRICT_EXTRAS }), newest.copyFolder, sharedRoot)
      newest.passes.push({ name: 'strict extras', ...strict })
    }
  } finally {
    if (!options.keep) fs.rmSync(temp, { recursive: true, force: true })
  }

  const report = results.map(({ module, folder, typescript, config, passes }) => ({
    module,
    folder,
    typescript,
    config,
    passes: passes.map(({ name, errors, fileCount }) => ({ name, errors, fileCount })),
  }))
  const ok = report.every((result) => result.passes.every((pass) => !pass.errors.length))
  const lines = []
  for (const result of report) {
    const summary = result.passes.length
      ? result.passes.map((pass) => `${pass.name} ${pass.errors.length ? `${pass.errors.length} error${pass.errors.length === 1 ? '' : 's'}` : 'OK'}`).join(', ')
      : 'no TypeScript files listed'
    lines.push(`${result.module} (${result.folder}, TypeScript ${result.typescript}, ${result.config}): ${summary}`)
    for (const pass of result.passes) for (const error of pass.errors) lines.push(`  [${pass.name}] ${error}`)
  }
  for (const warning of warnings) lines.push(`Warning: ${warning}`)
  if (ok) lines.push(`Shared renderer code type-checks in ${report.map((result) => result.module).join(', ') || 'no workspaces'}.`)
  return { ok, results: report, warnings, temp: options.keep ? temp : '', message: lines.join('\n') }
}

function parseArguments(argv) {
  const options = { modules: [] }
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index]
    const valued = (flag) => {
      if (argument === flag) {
        const value = argv[index + 1]
        if (!value || value.startsWith('--')) throw new CheckSetupError(`${flag} needs a value.`)
        index += 1
        return value
      }
      if (argument.startsWith(`${flag}=`)) return argument.slice(flag.length + 1)
      return undefined
    }
    let value
    if (argument === '--keep') options.keep = true
    else if (argument === '--help' || argument === '-h') options.help = true
    else if ((value = valued('--module')) !== undefined) options.modules.push(...value.split(',').map((item) => item.trim()).filter(Boolean))
    else if ((value = valued('--target-root')) !== undefined) options.targetRoot = value
    else if ((value = valued('--shared-root')) !== undefined) options.sharedRoot = value
    else if ((value = valued('--temp-root')) !== undefined) options.tempRoot = value
    else throw new CheckSetupError(`Unknown option: ${argument}`)
  }
  return options
}

/**
 * Command-line entry point.
 * @param {string[]} argv arguments after the script name
 * @returns {number} process exit code
 */
function main(argv) {
  try {
    const options = parseArguments(argv)
    if (options.help) {
      console.log('Usage: node scripts/check-shared-renderer.cjs [--module=<name>[,<name>]] [--target-root <dir>] [--shared-root <dir>] [--temp-root <dir>] [--keep]')
      return 0
    }
    const result = checkSharedRenderer(options)
    if (result.ok) console.log(result.message)
    else console.error(result.message)
    if (result.temp) console.log(`Temporary copy kept at ${result.temp}`)
    return result.ok ? 0 : 1
  } catch (error) {
    console.error(error instanceof CheckSetupError ? error.message : (error && error.stack) || error)
    return 2
  }
}

if (require.main === module) process.exitCode = main(process.argv.slice(2))

module.exports = { CheckSetupError, checkSharedRenderer, main }

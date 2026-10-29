/**
 * Portability test: run the Host half as if it had just been installed on a
 * different machine.
 *
 * The bundle is copied to a temporary directory that has no sibling `bridge`
 * folder, no virtual environment and no saved configuration, then mounted. This
 * is the first-run path a second machine or the Desktop profile takes, so it
 * must resolve everything from inside the package. Run with: node findtest.mjs
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { pathToFileURL } from 'node:url'

const here = path.dirname(fileURLToPath(import.meta.url))
const sandbox = fs.mkdtempSync(path.join(os.tmpdir(), 'autocad-bridge-portable-'))

// A bare install: the two halves, the packaged bridge sources and no .venv.
for (const file of ['index.js', 'client.js', 'package.json']) {
  fs.copyFileSync(path.join(here, file), path.join(sandbox, file))
}
fs.mkdirSync(path.join(sandbox, 'bridge'), { recursive: true })
for (const file of ['autocad_server.py', 'cad_com.py', 'config.py', 'requirements.txt', '.env']) {
  fs.copyFileSync(path.join(here, 'bridge', file), path.join(sandbox, 'bridge', file))
}

// No overrides, no saved configuration: exactly what a fresh install sees.
delete process.env.AUTOCAD_BRIDGE_DIR
delete process.env.AUTOCAD_BRIDGE_PYTHON

const mod = await import(pathToFileURL(path.join(sandbox, 'index.js')).href)

let route
const ctx = {
  effect(factory) {
    factory()
  },
  inject(names, callback) {
    callback({
      effect: (factory) => factory(),
      connection: {
        fetch: {
          register(definition) {
            route = definition
            return () => {}
          },
        },
      },
      subprocess: {
        async resolveExecutable(command) {
          return command
        },
        spawn() {
          throw new Error('spawn must not run during discovery')
        },
      },
    })
  },
}

mod.apply(ctx)

const response = await route.fetch(new Request('http://127.0.0.1:3080/api/autocad-bridge?action=status'))
const status = await response.json()

console.log('sandbox           =', sandbox)
console.log('found.bridgeDir   =', status.found.bridgeDir)
console.log('found.bundled     =', status.found.bundled)
console.log('found.hasVenv     =', status.found.hasVenv)
console.log('found.pythonSource=', status.found.pythonSource)
console.log('found.basePython  =', status.found.basePython)
console.log('found.hasRequirements =', status.found.hasRequirements)
console.log('config.pythonPath =', status.config.pythonPath)
console.log('config.serverScript =', status.config.serverScript)

const checks = [
  ['resolves the bridge inside the package', status.found.bundled === true],
  ['resolved directory is the sandbox copy', path.resolve(status.found.bridgeDir) === path.resolve(path.join(sandbox, 'bridge'))],
  ['server script exists there', fs.existsSync(status.config.serverScript)],
  ['no virtual environment yet', status.found.hasVenv === false],
  ['python falls back to PATH', status.found.pythonSource === 'path'],
  ['requirement list is reachable', status.found.hasRequirements === true],
  ['reports a port', Number.isInteger(status.config.httpPort)],
]
let failed = 0
for (const [label, ok] of checks) {
  console.log(`${ok ? 'ok  ' : 'FAIL'} ${label}`)
  if (!ok) failed += 1
}

// An override that points at a prepared directory must win over the package.
const overrideDir = path.join(sandbox, 'custom-bridge')
fs.mkdirSync(overrideDir, { recursive: true })
fs.copyFileSync(path.join(here, 'bridge', 'autocad_server.py'), path.join(overrideDir, 'autocad_server.py'))
process.env.AUTOCAD_BRIDGE_DIR = overrideDir

let overrideRoute
const overrideCtx = {
  effect(factory) {
    factory()
  },
  inject(names, callback) {
    callback({
      effect: (factory) => factory(),
      connection: {
        fetch: {
          register(definition) {
            overrideRoute = definition
            return () => {}
          },
        },
      },
      subprocess: { async resolveExecutable(c) { return c }, spawn() { throw new Error('no spawn') } },
    })
  },
}
const fresh = await import(pathToFileURL(path.join(sandbox, 'index.js')).href + '?override=1')
fresh.apply(overrideCtx)
const overrideStatus = await (await overrideRoute.fetch(new Request('http://127.0.0.1:3080/api/autocad-bridge?action=status'))).json()
const overrideOk = path.resolve(overrideStatus.found.bridgeDir) === path.resolve(overrideDir) && overrideStatus.found.bundled === false
console.log(`${overrideOk ? 'ok  ' : 'FAIL'} AUTOCAD_BRIDGE_DIR override wins over the packaged copy`)
if (!overrideOk) failed += 1

// A fresh install has no virtual environment, so installing dependencies must
// create it in the plugin data directory rather than beside the sources.
delete process.env.AUTOCAD_BRIDGE_DIR
// Point the data directory at the sandbox so the run cannot disturb a real
// installation living under the user's home.
const fakeHome = path.join(sandbox, 'home')
fs.mkdirSync(fakeHome, { recursive: true })
process.env.USERPROFILE = fakeHome
process.env.HOME = fakeHome
const recorded = []
let bootRoute
const bootCtx = {
  effect(factory) {
    factory()
  },
  inject(names, callback) {
    callback({
      effect: (factory) => factory(),
      connection: {
        fetch: {
          register(definition) {
            bootRoute = definition
            return () => {}
          },
        },
      },
      subprocess: {
        async resolveExecutable(command) {
          return command
        },
        spawn(spec) {
          recorded.push(spec.argv)
          // Stand in for `python -m venv`: produce the interpreter the next step
          // looks for, so the dependency install is actually reached.
          const target = spec.argv[spec.argv.indexOf('venv') + 1]
          if (target !== undefined) {
            const python = path.join(target, 'Scripts', 'python.exe')
            fs.mkdirSync(path.dirname(python), { recursive: true })
            fs.writeFileSync(python, '')
          }
          return {
            stdout: undefined,
            stderr: undefined,
            done: Promise.resolve({ exitCode: 0, signal: null }),
            terminate() {},
            async waitForExit() {
              return true
            },
          }
        },
      },
    })
  },
}
const bootMod = await import(pathToFileURL(path.join(sandbox, 'index.js')).href + '?boot=1')
bootMod.apply(bootCtx)
await (await bootRoute.fetch(new Request('http://127.0.0.1:3080/api/autocad-bridge?action=bootstrap', { method: 'POST', body: '{"action":"bootstrap"}' }))).json()
for (let i = 0; i < 60 && !recorded.some((argv) => argv.includes('pip')); i += 1) await new Promise((r) => setTimeout(r, 100))

const venvCall = recorded.find((argv) => argv.includes('venv'))
const expectedVenv = path.join(fakeHome, '.dsh', 'autocad-bridge', '.venv')
const venvOk = venvCall !== undefined && venvCall.some((arg) => path.resolve(arg) === path.resolve(expectedVenv))
console.log(`${venvOk ? 'ok  ' : 'FAIL'} creates the virtual environment in the plugin data directory`)
console.log('     venv call =', JSON.stringify(venvCall))
if (!venvOk) failed += 1
const pipCall = recorded.find((argv) => argv.includes('pip'))
const reqIndex = pipCall === undefined ? -1 : pipCall.indexOf('-r')
const requirementsArg = reqIndex >= 0 ? pipCall[reqIndex + 1] : undefined
const pipOk = pipCall !== undefined && requirementsArg !== undefined && fs.existsSync(requirementsArg)
console.log(`${pipOk ? 'ok  ' : 'FAIL'} installs from a reachable requirement list`)
console.log('     requirements =', requirementsArg, '| exists =', requirementsArg !== undefined && fs.existsSync(requirementsArg))
if (!pipOk) failed += 1

fs.rmSync(sandbox, { recursive: true, force: true })
console.log(failed === 0 ? '\nall portability checks passed' : `\n${failed} portability check(s) failed`)
process.exit(failed === 0 ? 0 : 1)

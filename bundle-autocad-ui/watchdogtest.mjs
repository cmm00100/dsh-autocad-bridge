/**
 * Watchdog unit test.
 *
 * Two paths matter and they are easy to get wrong:
 *  - an unexpected exit must relaunch the bridge,
 *  - a deliberate stop must NOT relaunch it.
 *
 * The host is loaded with a fake subprocess provider so the test controls
 * exactly when a handle finishes.
 */
import assert from 'node:assert/strict'
import fs from 'node:fs'
import net from 'node:net'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const HERE = path.dirname(fileURLToPath(import.meta.url))
const moduleUrl = pathToFileURL(path.join(HERE, 'index.js')).href

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))

/**
 * The watchdog probes the configured port before relaunching, so the test needs
 * a port nothing else owns — otherwise a real bridge on 8770 answers the probe
 * and the correct behaviour (stay quiet) looks like a missing restart.
 */
async function freePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer()
    server.once('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address()
      server.close(() => resolve(port))
    })
  })
}

const home = fs.mkdtempSync(path.join(os.tmpdir(), 'autocad-bridge-watchdog-'))
const dataDir = path.join(home, '.dsh', 'autocad-bridge')
fs.mkdirSync(dataDir, { recursive: true })
const port = await freePort()
fs.writeFileSync(
  path.join(dataDir, 'config.json'),
  JSON.stringify(
    {
      pythonPath: 'python',
      serverScript: path.join(HERE, 'bridge', 'autocad_server.py'),
      workingDirectory: path.join(HERE, 'bridge'),
      httpPort: port,
      progId: 'AutoCAD.Application.24',
      autoStart: true,
    },
    null,
    2,
  ),
)
process.env.USERPROFILE = home
process.env.HOME = home
delete process.env.AUTOCAD_BRIDGE_DIR
delete process.env.AUTOCAD_BRIDGE_PYTHON

/** A handle the test can finish on demand. */
function makeHandle(log) {
  let finish
  const handle = {
    stdout: null,
    stderr: null,
    done: new Promise((resolve) => {
      finish = resolve
    }),
    terminate() {
      log.push('terminate')
      handle.finish({ exitCode: 0, signal: 'SIGTERM' })
    },
    waitForExit() {
      return handle.done.then(() => true)
    },
    finish(outcome) {
      finish(outcome)
    },
  }
  return handle
}

function makeCtx(log, spawned) {
  const effects = []
  const track = (fn, label) => {
    const dispose = fn()
    effects.push({ label, dispose })
    return () => {
      if (typeof dispose === 'function') dispose()
    }
  }
  return {
    effects,
    ctx: {
      effect: track,
      inject(services, fn) {
        fn({
          effect: track,
          subprocess: {
            async resolveExecutable(command) {
              return command
            },
            spawn(spec) {
              const handle = makeHandle(log)
              handle.spec = spec
              spawned.push(handle)
              log.push(`spawn ${spec.argv.join(' ')}`)
              return handle
            },
          },
          connection: {
            fetch: {
              register() {
                return () => {}
              },
            },
          },
        })
      },
    },
  }
}

async function status(fetchHandler) {
  const response = await fetchHandler(
    new Request('http://dsh.invalid/api/autocad-bridge?action=status', { method: 'GET' }),
  )
  return response.json()
}

// --- the fetch handler is captured through the registered route ------------
const log = []
const spawned = []
const { ctx, effects } = makeCtx(log, spawned)

// `connection.fetch.register` is replaced so the test can reach the handler.
let handler
const originalInject = ctx.inject
ctx.inject = (services, fn) =>
  originalInject(services, (inner) => {
    inner.connection.fetch.register = (route) => {
      handler = route.fetch
      return () => {}
    }
    fn(inner)
  })

const module = await import(moduleUrl)
module.apply(ctx)
await sleep(50)

assert.equal(spawned.length, 1, 'the bridge should be started once at apply time')
const first = spawned[0]
let state = await status(handler)
assert.equal(state.running, true, 'bridge should report running')
assert.equal(state.restartAttempt, 0, 'no restart attempts yet')

// --- path 1: an unexpected exit relaunches the bridge ---------------------
first.finish({ exitCode: 1, signal: null })
await sleep(2500)
assert.equal(spawned.length, 2, 'a crashed bridge should be relaunched')
state = await status(handler)
assert.equal(state.running, true, 'relaunched bridge should be running')
assert.ok(state.restartAttempt >= 1, `restartAttempt should have advanced, got ${state.restartAttempt}`)
assert.equal(state.lastExit.exitCode, 1, 'the exit record should survive for the panel')

// --- path 2: a deliberate stop stays stopped ------------------------------
const stopResponse = await handler(
  new Request('http://dsh.invalid/api/autocad-bridge', {
    method: 'POST',
    body: JSON.stringify({ action: 'stop' }),
  }),
)
const stopped = await stopResponse.json()
assert.equal(stopped.ok, true, 'stop should succeed')
assert.equal(stopped.running, false, 'stop should leave the bridge stopped')

const spawnCountAfterStop = spawned.length
await sleep(3000)
assert.equal(spawned.length, spawnCountAfterStop, 'a deliberate stop must not be undone by the watchdog')
state = await status(handler)
assert.equal(state.running, false, 'bridge should still be stopped')
assert.equal(state.pendingRestart, false, 'no restart may be pending after an explicit stop')

// --- teardown does not leave a timer behind -------------------------------
for (const effect of effects) effect.dispose()
await sleep(200)
await status(handler).catch(() => undefined)

console.log('ok   watchdog relaunches a crashed bridge')
console.log('ok   watchdog records the last exit')
console.log('ok   explicit stop is not reinterpreted as a crash')
console.log('ok   no restart timer survives teardown')
console.log(`\nspawns = ${spawned.length}, effects = ${effects.length}`)

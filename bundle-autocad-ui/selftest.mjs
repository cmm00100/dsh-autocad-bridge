/**
 * Headless self-test for the Host half of the AutoCAD bridge bundle.
 *
 * Mounts `apply()` against a minimal fake Cordis context (real subprocess-free
 * spawn is NOT exercised here — that needs the Harness provider) so the API
 * route, the log ring buffer and the action dispatch can be verified without a
 * browser. Run with: node selftest.mjs
 */
import { spawn } from 'node:child_process'
import fs from 'node:fs'
import net from 'node:net'
import os from 'node:os'
import path from 'node:path'

const mod = await import(new URL('./index.js', import.meta.url).href)

let route

const fakeSubprocess = {
  async resolveExecutable(command) {
    if (!fs.existsSync(command)) throw new Error(`executable not found: ${command}`)
    return command
  },
  spawn(spec) {
    const child = spawn(spec.argv[0], spec.argv.slice(1), {
      cwd: spec.cwd,
      env: { ...process.env, ...(spec.env ?? {}) },
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    return {
      stdin: undefined,
      stdout: child.stdout,
      stderr: child.stderr,
      control: undefined,
      collected: {},
      done: new Promise((resolve) => {
        child.once('exit', (code, signal) => resolve({ exitCode: code, signal }))
        child.once('error', (error) => resolve({ exitCode: null, signal: null, error }))
      }),
      terminate() {
        child.kill()
      },
      async waitForExit() {
        await new Promise((resolve) => {
          if (child.exitCode !== null) return resolve()
          child.once('exit', resolve)
          setTimeout(resolve, 6000)
        })
        return true
      },
    }
  },
}

const effects = []
const ctx = {
  effect(factory) {
    effects.push(factory())
  },
  inject(names, callback) {
    callback({
      effect: (factory) => effects.push(factory()),
      connection: {
        fetch: {
          register(definition) {
            route = definition
            return () => {
              route = undefined
            }
          },
        },
      },
      subprocess: fakeSubprocess,
    })
  },
}

mod.apply(ctx)

const call = async (action, body) => {
  const request = new Request(`http://127.0.0.1:3080/api/autocad-bridge?action=${action}`, {
    method: body ? 'POST' : 'GET',
    ...(body ? { headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) } : {}),
  })
  const response = await route.fetch(request)
  return { status: response.status, body: await response.json() }
}

const waitPort = async (port, deadline) => {
  while (Date.now() < deadline) {
    const ok = await new Promise((resolve) => {
      const socket = net.connect({ host: '127.0.0.1', port })
      socket.once('connect', () => {
        socket.destroy()
        resolve(true)
      })
      socket.once('error', () => resolve(false))
    })
    if (ok) return true
    await new Promise((r) => setTimeout(r, 300))
  }
  return false
}

const step = (title) => console.log(`\n=== ${title} ===`)

step('1. module exports')
console.log('name =', mod.name, '| apply =', typeof mod.apply, '| route registered =', Boolean(route))

step('2. initial status')
const first = await call('status')
console.log('HTTP', first.status, '| running =', first.body.running, '| healthy =', first.body.healthy)
console.log('mcpUrl =', first.body.mcpUrl, '| progId =', first.body.config?.progId)

step('3. start')
const started = await call('start', { action: 'start' })
console.log('HTTP', started.status, '| ok =', started.body.ok, '| running =', started.body.running)

step('4. wait for HTTP endpoint')
const up = await waitPort(8770, Date.now() + 25000)
console.log('port 8770 listening =', up)

step('5. status after start')
const after = await call('status')
console.log('running =', after.body.running, '| healthy =', after.body.healthy, '| uptimeMs =', after.body.uptimeMs)
console.log('lastError =', after.body.lastError)
console.log('found =', JSON.stringify(after.body.found))
console.log('progIds (before detect) =', JSON.stringify(after.body.progIds))

step('5b. detect ProgIDs')
const detected = await call('progids', { action: 'progids' })
console.log('ok =', detected.body.ok, '| progIds =', JSON.stringify(detected.body.progIds))
if (!Array.isArray(detected.body.progIds) || detected.body.progIds.length === 0) {
  console.log('!! no ProgID detected on this machine')
}

step('5c. redetect')
const redetected = await call('redetect', { action: 'redetect' })
console.log('ok =', redetected.body.ok, '| bridgeDir =', redetected.body.found?.bridgeDir, '| bundled =', redetected.body.found?.bundled)

step('6. logs')
const logs = await call('logs', { action: 'logs', since: 0 })
console.log('log lines =', logs.body.lines.length)
for (const line of logs.body.lines.slice(0, 24)) console.log(`  [${line.level}] ${line.at} ${line.text}`)

step('7. saveConfig')
const saved = await call('saveConfig', { action: 'saveConfig', config: { httpPort: 8770, autoStart: true } })
console.log('saved =', saved.body.saved, '| file =', path.join(os.homedir(), '.dsh', 'autocad-bridge', 'config.json'))
console.log('exists =', fs.existsSync(path.join(os.homedir(), '.dsh', 'autocad-bridge', 'config.json')))

step('7b. bootstrap (reuses the prepared virtual environment)')
const boot = await call('bootstrap', { action: 'bootstrap' })
console.log('accepted =', boot.body.ok, '| started =', boot.body.started)
let bootDone = false
for (let i = 0; i < 120; i += 1) {
  await new Promise((r) => setTimeout(r, 1000))
  const probe = await call('status')
  if (probe.body.installing !== true) {
    bootDone = true
    console.log('installing finished after ~' + (i + 1) + 's | running =', probe.body.running, '| lastError =', probe.body.lastError)
    break
  }
}
if (!bootDone) console.log('!! bootstrap still running after 120s')
const bootLogs = await call('logs', { action: 'logs', since: 0 })
for (const line of bootLogs.body.lines.filter((l) => /虚拟环境|安装|pip|already satisfied|Requirement/i.test(l.text))) {
  console.log(`  [${line.level}] ${line.text}`)
}

step('8. stop')
const stopped = await call('stop', { action: 'stop' })
console.log('ok =', stopped.body.ok, '| running =', stopped.body.running)
await new Promise((r) => setTimeout(r, 1500))
const stoppedPort = await new Promise((resolve) => {
  const socket = net.connect({ host: '127.0.0.1', port: 8770 })
  socket.once('connect', () => {
    socket.destroy()
    resolve(true)
  })
  socket.once('error', () => resolve(false))
})
console.log('port still listening after stop =', stoppedPort)

step('9. teardown')
for (const dispose of effects) if (typeof dispose === 'function') dispose()
console.log('effects disposed =', effects.length)
process.exit(0)

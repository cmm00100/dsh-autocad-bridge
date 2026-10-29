/**
 * Host half of the AutoCAD bridge control panel.
 *
 * Owns the Python MCP bridge process (start / stop / restart / dependency
 * bootstrap), keeps a bounded ring buffer of its stderr/stdout, persists the
 * user-facing parameters, and exposes everything to the Client half through one
 * exact `/api` Fetch route. The Client module owns all rendering.
 *
 * Everything resolves at runtime — bundle directory, bridge directory, Python
 * interpreter, AutoCAD ProgID — so the same bundle works on another machine and
 * under another DSH profile (the web surface and the Electron desktop surface
 * both load the Client half through the same `dsh.client` declaration).
 */
import fs from 'node:fs'
import net from 'node:net'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const name = 'autocad-bridge'

const API_PATH = '/api/autocad-bridge'
const LOG_CAPACITY = 600
/** How many times a crashed bridge is relaunched before the watchdog stands down. */
const RESTART_ATTEMPTS = 8
const RESTART_BASE_DELAY_MS = 1000
const RESTART_MAX_DELAY_MS = 30000
/** A bridge that lives this long is healthy; the back-off ladder then resets. */
const RESTART_SETTLE_MS = 20000

/** Absolute directory of this bundle: the anchor every discovery path derives from. */
const BUNDLE_DIR = path.dirname(fileURLToPath(import.meta.url))

const SERVER_SCRIPT = 'autocad_server.py'
const REQUIREMENTS = 'requirements.txt'

function isFile(candidate) {
  if (typeof candidate !== 'string' || candidate === '') return false
  try {
    return fs.statSync(candidate).isFile()
  } catch {
    return false
  }
}

function isDirectory(candidate) {
  if (typeof candidate !== 'string' || candidate === '') return false
  try {
    return fs.statSync(candidate).isDirectory()
  } catch {
    return false
  }
}

/** The bundled, self-contained bridge directory shipped inside this package. */
function bundledBridgeDir() {
  return path.join(BUNDLE_DIR, 'bridge')
}

/**
 * Locate the Python bridge. An explicit environment override wins, then the copy
 * shipped inside the package, then a sibling `bridge` directory (the development
 * layout), then the current working directory.
 */
function findBridgeDir() {
  // An explicit override is authoritative: honour it whenever it holds the
  // server script, without second-guessing the user.
  const override = process.env.AUTOCAD_BRIDGE_DIR
  if (typeof override === 'string' && override.trim() !== '') {
    const dir = path.resolve(override)
    if (isFile(path.join(dir, SERVER_SCRIPT))) return dir
  }
  const candidates = [bundledBridgeDir(), path.join(BUNDLE_DIR, '..', 'bridge'), process.cwd()]
  const usable = []
  for (const candidate of candidates) {
    const dir = path.resolve(candidate)
    if (isFile(path.join(dir, SERVER_SCRIPT))) usable.push(dir)
  }
  // A directory that already carries its own prepared virtual environment wins
  // over a bare copy, so an existing installation keeps being used after an
  // upgrade instead of switching to fresher sources with no dependencies.
  for (const dir of usable) {
    if (treeVenvPython(dir) !== undefined) return dir
  }
  return usable[0] ?? bundledBridgeDir()
}

/** The virtual environment that lives beside a bridge directory, if any. */
function treeVenvPython(bridgeDir) {
  const candidates = [
    path.join(bridgeDir, '.venv', 'Scripts', 'python.exe'),
    path.join(bridgeDir, '.venv', 'bin', 'python'),
  ]
  for (const candidate of candidates) {
    if (isFile(candidate)) return candidate
  }
  return undefined
}

/**
 * Virtual environments this plugin may use, most preferred first.
 *
 * The data directory comes first because a bundle installed from a package
 * registry lives inside a read-only (or shared) module tree — an environment
 * written next to the sources would be lost on upgrade and could not be created
 * at all. The in-tree location stays supported so an existing checkout keeps
 * working without reinstalling its dependencies.
 */
function venvCandidates(bridgeDir) {
  const data = path.join(os.homedir(), '.dsh', 'autocad-bridge', '.venv')
  const trees = [data, path.join(bridgeDir, '.venv')]
  const candidates = []
  for (const tree of trees) {
    candidates.push(path.join(tree, 'Scripts', 'python.exe'), path.join(tree, 'bin', 'python'))
  }
  return candidates
}

function venvPython(bridgeDir) {
  for (const candidate of venvCandidates(bridgeDir)) {
    if (isFile(candidate)) return candidate
  }
  return undefined
}

/** Where a new virtual environment is created. */
function venvTargetDir() {
  return path.join(os.homedir(), '.dsh', 'autocad-bridge', '.venv')
}

/**
 * Resolve the interpreter used to run the bridge itself: the environment
 * override, the bridge virtual environment (preferred — it carries pywin32 and
 * the MCP SDK), then whatever `python` the PATH offers.
 */
function resolvePython(bridgeDir) {
  const override = process.env.AUTOCAD_BRIDGE_PYTHON
  if (typeof override === 'string' && override.trim() !== '' && isFile(override)) {
    return { pythonPath: override, pythonSource: 'env' }
  }
  const venv = venvPython(bridgeDir)
  if (venv !== undefined) return { pythonPath: venv, pythonSource: 'venv' }
  return { pythonPath: process.platform === 'win32' ? 'python' : 'python3', pythonSource: 'path' }
}

/** Any interpreter able to create the virtual environment when none exists yet. */
function resolveBasePython() {
  const override = process.env.AUTOCAD_BRIDGE_PYTHON
  if (typeof override === 'string' && override.trim() !== '' && isFile(override)) return override
  return process.platform === 'win32' ? 'python' : 'python3'
}

/** Everything the plugin needs, derived from the environment rather than hard-coded. */
function discover() {
  const bridgeDir = findBridgeDir()
  const { pythonPath, pythonSource } = resolvePython(bridgeDir)
  return {
    bridgeDir,
    serverScript: path.join(bridgeDir, SERVER_SCRIPT),
    workingDirectory: bridgeDir,
    pythonPath,
    pythonSource,
    bundled: path.resolve(bridgeDir) === path.resolve(bundledBridgeDir()),
    hasVenv: venvPython(bridgeDir) !== undefined,
    basePython: resolveBasePython(),
    progId: process.env.CAD_PROG_ID?.trim() || 'AutoCAD.Application.24',
    httpPort: 8770,
    autoStart: true,
  }
}

const PORTABLE_FIELDS = ['pythonPath', 'serverScript', 'workingDirectory', 'progId']

function stateFile() {
  return path.join(os.homedir(), '.dsh', 'autocad-bridge', 'config.json')
}

/**
 * Merge a persisted configuration onto the discovered defaults, repairing any
 * path the current machine cannot satisfy. A `config.json` carried over from
 * another machine therefore cannot pin the plugin to a directory that is gone.
 */
function reconcile(persisted) {
  const found = discover()
  const config = {
    pythonPath: found.pythonPath,
    serverScript: found.serverScript,
    workingDirectory: found.workingDirectory,
    httpPort: found.httpPort,
    progId: found.progId,
    autoStart: found.autoStart,
  }
  if (persisted === null || typeof persisted !== 'object') return { config, found }
  if (isFile(persisted.pythonPath)) {
    config.pythonPath = persisted.pythonPath
  } else if (typeof persisted.pythonPath === 'string' && persisted.pythonPath.trim() !== '' && !path.isAbsolute(persisted.pythonPath)) {
    config.pythonPath = persisted.pythonPath
  }
  if (isFile(persisted.serverScript)) config.serverScript = persisted.serverScript
  if (isDirectory(persisted.workingDirectory)) config.workingDirectory = persisted.workingDirectory
  const port = Number(persisted.httpPort)
  if (Number.isInteger(port) && port > 0 && port < 65536) config.httpPort = port
  if (typeof persisted.progId === 'string' && persisted.progId.trim() !== '') config.progId = persisted.progId.trim()
  if (typeof persisted.autoStart === 'boolean') config.autoStart = persisted.autoStart
  return { config, found }
}

function loadConfig() {
  let persisted
  try {
    persisted = JSON.parse(fs.readFileSync(stateFile(), 'utf8'))
  } catch {
    persisted = undefined
  }
  return reconcile(persisted)
}

function saveConfig(config) {
  const file = stateFile()
  fs.mkdirSync(path.dirname(file), { recursive: true })
  fs.writeFileSync(file, JSON.stringify(config, null, 2), 'utf8')
  return file
}

function stamp() {
  const d = new Date()
  const pad = (n, w = 2) => String(n).padStart(w, '0')
  return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`
}

function probePort(port, timeoutMs = 900) {
  return new Promise((resolve) => {
    const socket = net.connect({ host: '127.0.0.1', port })
    let settled = false
    const finish = (ok) => {
      if (settled) return
      settled = true
      socket.destroy()
      resolve(ok)
    }
    socket.setTimeout(timeoutMs)
    socket.once('connect', () => finish(true))
    socket.once('timeout', () => finish(false))
    socket.once('error', () => finish(false))
  })
}

function json(status, body) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' },
  })
}

/** A one-liner that lists the AutoCAD ProgIDs this machine actually registers. */
const PROGID_SCRIPT = [
  'import winreg',
  'root = winreg.HKEY_CLASSES_ROOT',
  'out = []',
  'i = 0',
  'while True:',
  '    try:',
  '        key = winreg.EnumKey(root, i)',
  '    except OSError:',
  '        break',
  '    i += 1',
  '    if key.startswith("AutoCAD.Application.") and key.count(".") == 2:',
  '        out.append(key)',
  'print("\\n".join(sorted(out)))',
].join('\n')

/**
 * @param {object} ctx Cordis host context.
 */
function apply(ctx) {
  const loaded = loadConfig()
  const state = {
    config: loaded.config,
    found: loaded.found,
    lines: [],
    seq: 0,
    handle: undefined,
    exit: undefined,
    startedAt: undefined,
    lastError: undefined,
    busy: false,
    installing: false,
    progIds: undefined,
    /** Handles stopped on purpose; their exit must not trigger the watchdog. */
    retired: new Set(),
    restartAttempt: 0,
    restartTimer: undefined,
    lastExit: undefined,
  }

  /** Filled by the `ctx.inject` scope below; process actions are refused until then. */
  let sp

  const push = (level, text) => {
    for (const raw of String(text).split(/\r?\n/)) {
      const line = raw.trimEnd()
      if (line === '') continue
      state.seq += 1
      state.lines.push({ n: state.seq, at: stamp(), level, text: line })
    }
    if (state.lines.length > LOG_CAPACITY) state.lines.splice(0, state.lines.length - LOG_CAPACITY)
  }

  push('info', 'autocad-bridge 插件已加载，桥进程待命。')
  push('info', `桥目录：${state.found.bridgeDir}（${state.found.bundled ? '插件内置' : '外部目录'}）`)
  push('info', `Python：${state.config.pythonPath}（来源：${state.found.hasVenv ? '虚拟环境' : 'PATH'}）`)

  const attached = (stream, level) => {
    if (!stream) return
    stream.setEncoding?.('utf8')
    stream.on('data', (chunk) => push(level, chunk))
  }

  /** Child environment: the host environment plus the bridge's own overrides. */
  const childEnv = (extra) => ({
    ...process.env,
    PYTHONUTF8: '1',
    PYTHONIOENCODING: 'utf-8',
    ...extra,
  })

  const spawnStep = (argv, cwd, level = 'out') => {
    if (!sp) throw new Error('subprocess 服务尚未就绪，请稍后重试。')
    push('info', `$ ${argv.join(' ')}`)
    const handle = sp.spawn({
      argv,
      cwd,
      stdio: { stdin: 'ignore', stdout: 'pipe', stderr: 'pipe' },
      graceMs: 5000,
      env: childEnv(),
    })
    attached(handle.stdout, level)
    attached(handle.stderr, level)
    return handle
  }

  const stop = async () => {
    const handle = state.handle
    state.handle = undefined
    state.startedAt = undefined
    if (state.restartTimer !== undefined) {
      clearTimeout(state.restartTimer)
      state.restartTimer = undefined
    }
    if (!handle) return false
    // Mark the handle itself: the exit hook may run after this function returns,
    // so a global flag would already be cleared and look like a crash.
    state.retired.add(handle)
    push('info', '正在终止桥进程…')
    try {
      handle.terminate()
      await handle.waitForExit(new AbortController().signal)
    } catch {
      /* the provider owns teardown; a refused wait is not an error here */
    }
    push('info', '桥进程已终止。')
    return true
  }

  /**
   * Bring the bridge back after an unexpected exit. Backs off up to
   * `RESTART_MAX_DELAY_MS` so a permanently broken configuration cannot spin.
   */
  const scheduleRestart = (reason) => {
    if (!state.config.autoStart) return
    if (state.restartTimer !== undefined) return
    if (state.restartAttempt >= RESTART_ATTEMPTS) {
      push('err', `桥进程已连续失败 ${RESTART_ATTEMPTS} 次，已停止自动重启。请检查参数后点「重启」。`)
      return
    }
    const delay = Math.min(RESTART_BASE_DELAY_MS * 2 ** state.restartAttempt, RESTART_MAX_DELAY_MS)
    state.restartAttempt += 1
    push('warn', `${reason}，${Math.round(delay / 100) / 10} 秒后第 ${state.restartAttempt} 次自动重启…`)
    state.restartTimer = setTimeout(() => {
      state.restartTimer = undefined
      if (state.handle !== undefined) return
      start().catch((error) => {
        state.lastError = String(error?.message ?? error)
        push('err', `自动重启失败：${state.lastError}`)
        scheduleRestart('自动重启未成功')
      })
    }, delay)
    state.restartTimer.unref?.()
  }

  const start = async () => {
    if (state.handle) return false
    if (!sp) throw new Error('subprocess 服务尚未就绪，请稍后重试。')
    const { config } = state
    if (!isFile(config.serverScript)) {
      throw new Error(`找不到桥服务脚本：${config.serverScript}。请在“连接参数”里改正，或点“安装依赖”重建。`)
    }
    state.lastError = undefined
    push('info', `启动桥进程：${config.pythonPath} ${config.serverScript}`)
    const executable = await sp.resolveExecutable(config.pythonPath)
    const handle = sp.spawn({
      argv: [executable, config.serverScript],
      cwd: config.workingDirectory,
      stdio: { stdin: 'ignore', stdout: 'pipe', stderr: 'pipe' },
      graceMs: 5000,
      env: childEnv({
        MCP_TRANSPORT: 'http',
        MCP_HTTP_HOST: '127.0.0.1',
        MCP_HTTP_PORT: String(config.httpPort),
        CAD_PROG_ID: config.progId,
      }),
    })
    state.handle = handle
    state.startedAt = Date.now()
    attached(handle.stdout, 'out')
    attached(handle.stderr, 'err')
    // A process that survives this long counts as healthy, so the watchdog's
    // back-off ladder starts over instead of exhausting its attempts.
    const settle = setTimeout(() => {
      if (state.handle === handle) state.restartAttempt = 0
    }, RESTART_SETTLE_MS)
    settle.unref?.()
    state.exit = handle.done.then(
      (outcome) => {
        state.lastExit = { exitCode: outcome.exitCode, signal: outcome.signal ?? null, at: stamp() }
        push('info', `桥进程退出：exitCode=${outcome.exitCode} signal=${outcome.signal ?? 'null'}`)
        if (state.handle === handle) {
          state.handle = undefined
          state.startedAt = undefined
        }
        onUnexpectedExit(handle)
        return outcome
      },
      (error) => {
        state.lastError = String(error?.message ?? error)
        push('err', `桥进程启动失败：${state.lastError}`)
        state.lastExit = { exitCode: null, signal: null, at: stamp() }
        onUnexpectedExit(handle)
        return undefined
      },
    )
    return true
  }

  /**
   * A bridge launcher may hand the listening socket to its own child, so the
   * handle ending is not proof that the endpoint died. Probe before relaunching,
   * otherwise the watchdog would fight a perfectly healthy bridge over the port.
   */
  const onUnexpectedExit = (handle) => {
    if (state.retired.has(handle)) {
      state.retired.delete(handle)
      return
    }
    if (state.handle !== undefined && state.handle !== handle) return
    probePort(state.config.httpPort)
      .then((alive) => {
        if (state.handle !== undefined) return
        if (alive) {
          push('warn', `启动器已退出，但 ${state.config.httpPort} 端口仍在服务，视为正常运行。`)
          state.restartAttempt = 0
          return
        }
        scheduleRestart('桥进程已停止')
      })
      .catch(() => scheduleRestart('桥进程已停止'))
  }

  const restart = async () => {
    await stop()
    return start()
  }

  const snapshot = async () => {
    const port = state.config.httpPort
    const healthy = await probePort(port)
    return {
      running: state.handle !== undefined,
      healthy,
      busy: state.busy,
      installing: state.installing,
      startedAt: state.startedAt ?? null,
      uptimeMs: state.startedAt ? Date.now() - state.startedAt : 0,
      lastError: state.lastError ?? null,
      config: state.config,
      found: {
        bridgeDir: state.found.bridgeDir,
        bundled: state.found.bundled,
        hasVenv: state.found.hasVenv,
        pythonSource: state.found.pythonSource,
        basePython: state.found.basePython,
        stateFile: stateFile(),
        // True when a requirement list is reachable — either next to the server
        // script, or the copy shipped inside this bundle.
        hasRequirements:
          isFile(path.join(state.found.bridgeDir, REQUIREMENTS)) || isFile(path.join(bundledBridgeDir(), REQUIREMENTS)),
      },
      progIds: state.progIds ?? null,
      /** Watchdog state: how many relaunches happened, and the last exit seen. */
      restartAttempt: state.restartAttempt,
      pendingRestart: state.restartTimer !== undefined,
      lastExit: state.lastExit ?? null,
      mcpUrl: `http://127.0.0.1:${port}/mcp`,
      log: { seq: state.seq, lines: state.lines },
    }
  }

  const logSince = (since) => ({ seq: state.seq, lines: state.lines.filter((l) => l.n > since) })

  const guard = async (work) => {
    if (state.busy) return { ok: false, error: '上一个操作尚未完成，请稍候。' }
    state.busy = true
    try {
      return await work()
    } catch (error) {
      const message = String(error?.message ?? error)
      state.lastError = message
      push('err', `操作失败：${message}`)
      return { ok: false, error: message }
    } finally {
      state.busy = false
    }
  }

  /** Enumerate the AutoCAD ProgIDs this machine registers, through a throwaway Python run. */
  const detectProgIds = async () => {
    const bridgeDir = state.found.bridgeDir
    const python = isFile(state.config.pythonPath) ? state.config.pythonPath : state.found.basePython
    const executable = await sp.resolveExecutable(python)
    const handle = sp.spawn({
      argv: [executable, '-c', PROGID_SCRIPT],
      cwd: bridgeDir,
      stdio: { stdin: 'ignore', stdout: 'pipe', stderr: 'pipe' },
      graceMs: 5000,
      env: childEnv(),
    })
    let buffer = ''
    handle.stdout?.setEncoding?.('utf8')
    handle.stdout?.on('data', (chunk) => {
      buffer += String(chunk)
    })
    handle.stderr?.setEncoding?.('utf8')
    handle.stderr?.on('data', () => {})
    await handle.done
    const progIds = buffer
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => /^AutoCAD\.Application\.\d{2}(\.\d)?$/.test(line))
    state.progIds = progIds
    push('info', progIds.length > 0 ? `检测到 AutoCAD ProgID：${progIds.join(', ')}` : '未检测到任何 AutoCAD ProgID，请确认已安装 AutoCAD。')
    return progIds
  }

  /**
   * Create the virtual environment and install the bridge dependencies. Runs in
   * the background so the request returns at once and progress arrives as log
   * lines; `state.installing` reports it in every status snapshot.
   */
  const bootstrap = async () => {
    if (state.installing) return { ok: false, error: '依赖正在安装中，请稍候。' }
    if (!sp) throw new Error('subprocess 服务尚未就绪，请稍后重试。')
    const bridgeDir = state.found.bridgeDir
    // An external bridge directory may be a bare copy of the server script, so
    // fall back to the requirement list shipped inside this bundle.
    const localRequirements = path.join(bridgeDir, REQUIREMENTS)
    const requirements = isFile(localRequirements) ? localRequirements : path.join(bundledBridgeDir(), REQUIREMENTS)
    if (!isFile(requirements)) {
      return { ok: false, error: `找不到 ${REQUIREMENTS}，无法安装依赖。` }
    }
    state.installing = true
    await stop().catch(() => false)
    const run = async () => {
      try {
        // Create the environment in the plugin data directory, never next to the
        // sources, so an upgrade of the package cannot take the dependencies away.
        const venvDir = venvTargetDir()
        let python = venvPython(bridgeDir)
        if (python === undefined) {
          const base = await sp.resolveExecutable(state.found.basePython)
          push('info', `创建虚拟环境：${venvDir}`)
          fs.mkdirSync(path.dirname(venvDir), { recursive: true })
          const created = await spawnStep([base, '-m', 'venv', venvDir], bridgeDir).done
          if (created.exitCode !== 0) throw new Error(`创建虚拟环境失败（exitCode=${created.exitCode}）`)
          python = venvPython(bridgeDir)
          if (python === undefined) throw new Error('虚拟环境已创建，但找不到其中的 python。')
        }
        const resolved = await sp.resolveExecutable(python)
        push('info', '安装桥依赖：pywin32 / mcp / python-dotenv')
        const installed = await spawnStep([resolved, '-m', 'pip', 'install', '--disable-pip-version-check', '-r', requirements], bridgeDir).done
        if (installed.exitCode !== 0) throw new Error(`依赖安装失败（exitCode=${installed.exitCode}）`)
        const after = discover()
        state.found = after
        state.config = { ...state.config, pythonPath: after.pythonPath, serverScript: after.serverScript, workingDirectory: after.workingDirectory }
        saveConfig(state.config)
        push('info', `依赖安装完成，Python：${after.pythonPath}`)
        if (state.config.autoStart) await start()
      } catch (error) {
        const message = String(error?.message ?? error)
        state.lastError = message
        push('err', `安装失败：${message}`)
      } finally {
        state.installing = false
      }
    }
    run()
    return { ok: true, started: true }
  }

  const dispatch = async (action, payload) => {
    switch (action) {
      case 'status':
        return snapshot()
      case 'logs':
        return logSince(Number(payload?.since ?? 0))
      case 'start':
        return guard(async () => {
          await start()
          return { ok: true, ...(await snapshot()) }
        })
      case 'stop':
        return guard(async () => {
          await stop()
          return { ok: true, ...(await snapshot()) }
        })
      case 'restart':
        return guard(async () => {
          await restart()
          return { ok: true, ...(await snapshot()) }
        })
      case 'redetect':
        return guard(async () => {
          const found = discover()
          const config = reconcile(state.config).config
          state.found = found
          state.config = { ...config, pythonPath: config.pythonPath }
          saveConfig(state.config)
          push('info', `重新检测完成：${state.found.bridgeDir} / ${state.config.pythonPath}`)
          return { ok: true, ...(await snapshot()) }
        })
      case 'progids':
        return guard(async () => {
          await detectProgIds()
          return { ok: true, ...(await snapshot()) }
        })
      case 'bootstrap':
        return guard(() => bootstrap())
      case 'saveConfig': {
        const next = { ...state.config }
        const incoming = payload?.config ?? {}
        for (const field of PORTABLE_FIELDS) {
          const value = incoming[field]
          if (typeof value === 'string' && value.trim() !== '') next[field] = value.trim()
        }
        const port = Number(incoming.httpPort)
        if (Number.isInteger(port) && port > 0 && port < 65536) next.httpPort = port
        if (typeof incoming.autoStart === 'boolean') next.autoStart = incoming.autoStart
        state.config = next
        saveConfig(next)
        push('info', `参数已保存到 ${stateFile()}`)
        return { saved: true, config: next }
      }
      default:
        return { error: `未知操作：${action}` }
    }
  }

  ctx.inject(['connection', 'subprocess'], (inner) => {
    sp = inner.subprocess
    inner.effect(
      () =>
        inner.connection.fetch.register({
          path: API_PATH,
          methods: ['GET', 'POST'],
          requestBody: 'buffered',
          fetch: async (request) => {
            try {
              const url = new URL(request.url, 'http://dsh.invalid')
              let action = url.searchParams.get('action') ?? 'status'
              let payload
              if (request.method === 'POST') {
                const text = await request.text()
                if (text.trim() !== '') {
                  payload = JSON.parse(text)
                  if (typeof payload?.action === 'string') action = payload.action
                }
              }
              const value = await dispatch(action, payload)
              return json(200, { action, ...value })
            } catch (error) {
              return json(500, { error: String(error?.message ?? error) })
            }
          },
        }),
      'autocad-bridge: api route',
    )

    inner.effect(() => {
      if (state.config.autoStart) {
        start().catch((error) => push('err', `自动启动失败：${String(error?.message ?? error)}`))
      }
      return () => {
        stop().catch(() => {})
      }
    }, 'autocad-bridge: process lifecycle')
  })
}

export { apply, name }

/**
 * Client half of the AutoCAD 2024 bridge control panel.
 *
 * Renders the central `main` panel and the sidebar entry, and talks to the Host
 * half through the exact `/api/autocad-bridge` Fetch route. Only theme tokens are
 * used for colour; all markup and behaviour are local to this package.
 */
window.__ModuleLoader__.load({
  id: '@local/autocad-bridge',
  factory(require) {
    const React = require('react')
    const h = React.createElement

    const NS = 'autocad-bridge'
    const PANEL_ID = 'autocad-bridge'
    const API = '/api/autocad-bridge'

    const zh = {
      panel: 'AutoCAD 桥接',
      title: 'AutoCAD 2024 桥接',
      subtitle: '管理连接本机 AutoCAD 2024 的 MCP 服务进程：启动、重启、查看实时日志。',
      stateRunning: '运行中',
      stateStopped: '已停止',
      stateDegraded: '进程在但不通',
      stateChecking: '检测中…',
      start: '启动',
      stop: '停止',
      restart: '重启',
      saving: '保存中…',
      save: '保存参数',
      saved: '参数已保存',
      runStatus: '运行状态',
      setting: '连接参数',
      log: '实时日志',
      labelProcess: '进程',
      labelPort: '端口',
      labelEndpoint: 'MCP 端点',
      labelUptime: '运行时长',
      labelTransport: '传输方式',
      labelLastError: '最近错误',
      valueYes: '已启动',
      valueNo: '未启动',
      labelPython: 'Python 解释器',
      labelScript: '服务脚本',
      labelCwd: '工作目录',
      labelPortField: 'HTTP 端口',
      labelProgId: 'ProgID',
      labelAutoStart: '随插件自动启动',
      empty: '暂无日志。点击「启动」拉起桥进程。',
      none: '无',
      logCount: '共 {n} 行',
      agentHint: 'Agent 侧工具：mcp__autocad__cad_*',
      labelBridgeDir: '桥目录',
      labelPythonSource: '解释器来源',
      labelStateFile: '配置文件',
      labelProgIds: '已检测 CAD',
      bundledYes: '插件内置',
      bundledNo: '外部目录',
      sourceVenv: '虚拟环境',
      sourceEnv: '环境变量',
      sourcePath: '系统 PATH',
      builtin: '内置',
      bootstrap: '安装依赖',
      bootstrapping: '安装中…',
      redetect: '重新检测',
      detectProgIds: '检测 CAD',
      installHint: '换机器或首次使用：点「安装依赖」自动建虚拟环境并装 pywin32 / mcp / python-dotenv。上方的自动发现可被环境变量 AUTOCAD_BRIDGE_PYTHON / AUTOCAD_BRIDGE_DIR 覆盖。',
      progIdsEmpty: '未检测',
      labelWatchdog: '自动重启',
      watchdogIdle: '待命',
      watchdogWaiting: '等待第 {n} 次重启',
      watchdogRetried: '已重启 {n} 次',
      portHint: '端口要改在插件的 cordis.patch.yml 里（autocad-mcp 行的 url），本页只负责启动进程；改完重载插件才生效。',
    }

    const en = {
      panel: 'AutoCAD Bridge',
      title: 'AutoCAD 2024 Bridge',
      subtitle: 'Manage the MCP service process that drives the local AutoCAD 2024 instance.',
      stateRunning: 'Running',
      stateStopped: 'Stopped',
      stateDegraded: 'Up but unreachable',
      stateChecking: 'Checking…',
      start: 'Start',
      stop: 'Stop',
      restart: 'Restart',
      saving: 'Saving…',
      save: 'Save settings',
      saved: 'Settings saved',
      runStatus: 'Runtime',
      setting: 'Connection',
      log: 'Live log',
      labelProcess: 'Process',
      labelPort: 'Port',
      labelEndpoint: 'MCP endpoint',
      labelUptime: 'Uptime',
      labelTransport: 'Transport',
      labelLastError: 'Last error',
      valueYes: 'Running',
      valueNo: 'Not started',
      labelPython: 'Python executable',
      labelScript: 'Server script',
      labelCwd: 'Working directory',
      labelPortField: 'HTTP port',
      labelProgId: 'ProgID',
      labelAutoStart: 'Start with the plugin',
      empty: 'No output yet. Press Start to launch the bridge.',
      none: 'none',
      logCount: '{n} lines',
      agentHint: 'Agent tools: mcp__autocad__cad_*',
      labelBridgeDir: 'Bridge directory',
      labelPythonSource: 'Interpreter from',
      labelStateFile: 'State file',
      labelProgIds: 'Detected CAD',
      bundledYes: 'Bundled',
      bundledNo: 'External',
      sourceVenv: 'Virtual environment',
      sourceEnv: 'Environment override',
      sourcePath: 'System PATH',
      builtin: 'built-in',
      bootstrap: 'Install deps',
      bootstrapping: 'Installing…',
      redetect: 'Re-detect',
      detectProgIds: 'Detect CAD',
      installHint: 'New machine or first run: press Install deps to create the virtual environment and install pywin32 / mcp / python-dotenv. Discovery above can be overridden with the AUTOCAD_BRIDGE_PYTHON / AUTOCAD_BRIDGE_DIR environment variables.',
      progIdsEmpty: 'not detected',
      labelWatchdog: 'Auto restart',
      watchdogIdle: 'standing by',
      watchdogWaiting: 'restart #{n} pending',
      watchdogRetried: 'restarted {n}x',
      portHint: 'The port lives in this bundle\'s cordis.patch.yml (the autocad-mcp row url); this page only launches the process. Reload the plugin after editing it.',
    }

    const CSS = `
.acb-root {
  display: flex;
  flex-direction: column;
  gap: 16px;
  box-sizing: border-box;
  height: 100%;
  min-height: 0;
  padding: 20px 24px 24px;
  overflow: auto;
  background: var(--dsw-alias-bg-base);
  color: var(--dsw-alias-label-primary);
  font-size: 13px;
  line-height: 20px;
}
.acb-head {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
  flex-wrap: wrap;
}
.acb-head-main { display: flex; flex-direction: column; gap: 6px; min-width: 260px; }
.acb-title-row { display: flex; align-items: center; gap: 10px; }
.acb-title { margin: 0; font-size: 16px; font-weight: 600; line-height: 24px; }
.acb-sub { margin: 0; color: var(--dsw-alias-label-secondary); font-size: 12px; line-height: 18px; max-width: 62ch; }
.acb-badge {
  display: inline-flex;
  align-items: center;
  height: 20px;
  padding: 0 8px;
  border-radius: 10px;
  border: 1px solid var(--dsw-alias-border-l1);
  color: var(--dsw-alias-label-secondary);
  font-size: 11px;
  white-space: nowrap;
}
.acb-dot { width: 8px; height: 8px; border-radius: 50%; flex: none; background: var(--dsw-alias-state-idle-primary); }
.acb-dot[data-state="running"] { background: var(--dsw-alias-state-success-primary); }
.acb-dot[data-state="degraded"] { background: var(--dsw-alias-state-warn-primary); }
.acb-dot[data-state="stopped"] { background: var(--dsw-alias-state-idle-primary); }
.acb-actions { display: flex; align-items: center; gap: 8px; }
.acb-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  height: 30px;
  padding: 0 14px;
  border-radius: 8px;
  border: 1px solid var(--dsw-alias-border-l1);
  background: transparent;
  color: var(--dsw-alias-label-primary);
  font-size: 13px;
  font-family: inherit;
  cursor: pointer;
}
.acb-btn:hover:not(:disabled) { background: var(--dsw-alias-bg-layer-2); }
.acb-btn:disabled { opacity: 0.5; cursor: default; }
.acb-btn[data-variant="primary"] {
  border-color: transparent;
  background: var(--dsw-alias-brand-primary);
  color: var(--dsw-alias-bg-base);
  font-weight: 600;
}
.acb-grid { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1.15fr); gap: 16px; }
@media (max-width: 900px) { .acb-grid { grid-template-columns: minmax(0, 1fr); } }
.acb-card {
  display: flex;
  flex-direction: column;
  gap: 12px;
  min-width: 0;
  padding: 16px;
  border-radius: 12px;
  border: 1px solid var(--dsw-alias-border-l1);
  background: var(--dsw-alias-bg-layer-1);
}
.acb-card-title {
  margin: 0;
  font-size: 12px;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--dsw-alias-label-secondary);
}
.acb-kv { display: grid; grid-template-columns: auto minmax(0, 1fr); gap: 8px 14px; margin: 0; }
.acb-kv dt { color: var(--dsw-alias-label-secondary); white-space: nowrap; }
.acb-kv dd { margin: 0; min-width: 0; overflow: hidden; text-overflow: ellipsis; }
.acb-mono { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size: 12px; }
.acb-field { display: flex; flex-direction: column; gap: 6px; min-width: 0; }
.acb-label { color: var(--dsw-alias-label-secondary); font-size: 12px; }
.acb-input {
  box-sizing: border-box;
  width: 100%;
  height: 30px;
  padding: 0 10px;
  border-radius: 8px;
  border: 1px solid var(--dsw-alias-border-l1);
  background: var(--dsw-alias-bg-base);
  color: var(--dsw-alias-label-primary);
  font-size: 12px;
  font-family: inherit;
}
.acb-input:focus { outline: none; border-color: var(--dsw-alias-brand-primary); }
.acb-check { display: flex; align-items: center; gap: 8px; color: var(--dsw-alias-label-primary); }
.acb-tag-ok { color: var(--dsw-alias-state-success-primary); }
.acb-tag-warn { color: var(--dsw-alias-state-warn-primary); }
.acb-tag-idle { color: var(--dsw-alias-label-secondary); }
.acb-tag-bad { color: var(--dsw-alias-state-error-primary); }
.acb-logs { flex: 1 1 auto; min-height: 220px; }
.acb-logs-head { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
.acb-pre {
  box-sizing: border-box;
  flex: 1 1 auto;
  min-height: 160px;
  max-height: 340px;
  margin: 0;
  padding: 12px;
  overflow: auto;
  border-radius: 10px;
  border: 1px solid var(--dsw-alias-border-l1);
  background: var(--dsw-alias-bg-base);
  font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  font-size: 12px;
  line-height: 18px;
  white-space: pre-wrap;
  word-break: break-word;
}
.acb-line { display: flex; gap: 10px; }
.acb-line-at { color: var(--dsw-alias-label-secondary); flex: none; }
.acb-line[data-level="err"] .acb-line-text { color: var(--dsw-alias-state-error-primary); }
.acb-line[data-level="out"] .acb-line-text { color: var(--dsw-alias-label-secondary); }
.acb-notice {
  padding: 8px 12px;
  border-radius: 8px;
  border: 1px solid var(--dsw-alias-border-l1);
  background: var(--dsw-alias-bg-layer-2);
  color: var(--dsw-alias-label-secondary);
  font-size: 12px;
}
`

    function call(action, payload) {
      return fetch(API, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ action, ...(payload ?? {}) }),
      }).then(async (response) => {
        const data = await response.json().catch(() => ({}))
        if (!response.ok) throw new Error(data.error ?? `HTTP ${response.status}`)
        return data
      })
    }

    function formatUptime(ms) {
      if (!ms || ms < 1000) return '—'
      const total = Math.floor(ms / 1000)
      const hh = Math.floor(total / 3600)
      const mm = Math.floor((total % 3600) / 60)
      const ss = total % 60
      const pad = (n) => String(n).padStart(2, '0')
      return hh > 0 ? `${hh}:${pad(mm)}:${pad(ss)}` : `${pad(mm)}:${pad(ss)}`
    }

    function BridgePanel(props) {
      const t = props?.t
      const [snapshot, setSnapshot] = React.useState(null)
      const [lines, setLines] = React.useState([])
      const [draft, setDraft] = React.useState(null)
      const [pending, setPending] = React.useState(null)
      const [notice, setNotice] = React.useState(null)
      const [bag] = React.useState(() => ({ pre: null }))
      const tr = (key, params) => (typeof t === 'function' ? t(key, params) : String(key))

      React.useEffect(() => {
        let alive = true
        const tick = async () => {
          try {
            const data = await call('status')
            if (!alive) return
            setSnapshot(data)
            setDraft((current) => current ?? { ...data.config })
          } catch (error) {
            if (alive) setNotice(String(error.message ?? error))
          }
        }
        tick()
        const timer = window.setInterval(tick, 2500)
        return () => {
          alive = false
          window.clearInterval(timer)
        }
      }, [])

      React.useEffect(() => {
        let alive = true
        let cursor = 0
        const tick = async () => {
          try {
            const data = await call('logs', { since: cursor })
            if (!alive) return
            cursor = data.seq ?? cursor
            if (Array.isArray(data.lines) && data.lines.length > 0) {
              setLines((prev) => {
                const next = prev.concat(data.lines)
                return next.length > 600 ? next.slice(next.length - 600) : next
              })
            }
          } catch {
            /* status polling already surfaces transport problems */
          }
        }
        tick()
        const timer = window.setInterval(tick, 1200)
        return () => {
          alive = false
          window.clearInterval(timer)
        }
      }, [])

      React.useEffect(() => {
        const node = bag.pre
        if (node) node.scrollTop = node.scrollHeight
      }, [lines])

      const run = async (action, label) => {
        setPending(action)
        setNotice(null)
        try {
          const data = await call(action)
          if (data && data.ok === false) setNotice(data.error ?? 'failed')
          else if (data && data.config) setSnapshot((prev) => ({ ...(prev ?? {}), ...data }))
          else if (data) setSnapshot((prev) => ({ ...(prev ?? {}), ...data }))
        } catch (error) {
          setNotice(String(error.message ?? error))
        } finally {
          setPending(null)
        }
      }

      const save = async () => {
        if (!draft) return
        setPending('save')
        setNotice(null)
        try {
          const data = await call('saveConfig', { config: draft })
          setSnapshot((prev) => ({ ...(prev ?? {}), config: data.config }))
          setNotice(tr('saved'))
        } catch (error) {
          setNotice(String(error.message ?? error))
        } finally {
          setPending(null)
        }
      }

      const running = snapshot?.running === true
      const healthy = snapshot?.healthy === true
      const stateKey = !snapshot ? 'checking' : running && healthy ? 'running' : running ? 'degraded' : 'stopped'
      const stateText =
        stateKey === 'running'
          ? tr('stateRunning')
          : stateKey === 'degraded'
            ? tr('stateDegraded')
            : stateKey === 'stopped'
              ? tr('stateStopped')
              : tr('stateChecking')
      const dotState = stateKey === 'checking' ? 'stopped' : stateKey

      const field = (labelKey, key, options) =>
        h(
          'label',
          { className: 'acb-field' },
          h('span', { className: 'acb-label' }, tr(labelKey)),
          h('input', {
            className: 'acb-input',
            value: draft?.[key] ?? '',
            spellCheck: false,
            onChange: (event) => setDraft({ ...(draft ?? {}), [key]: event.target.value }),
            ...(options ?? {}),
          }),
        )

      return h(
        'div',
        { className: 'acb-root' },
        h(
          'header',
          { className: 'acb-head' },
          h(
            'div',
            { className: 'acb-head-main' },
            h(
              'div',
              { className: 'acb-title-row' },
              h('span', { className: 'acb-dot', 'data-state': dotState }),
              h('h1', { className: 'acb-title' }, tr('title')),
              h('span', { className: 'acb-badge' }, stateText),
            ),
            h('p', { className: 'acb-sub' }, tr('subtitle')),
          ),
          h(
            'div',
            { className: 'acb-actions' },
            h(
              'button',
              {
                className: 'acb-btn',
                'data-variant': 'primary',
                disabled: pending !== null || running,
                onClick: () => run('start'),
              },
              tr('start'),
            ),
            h(
              'button',
              { className: 'acb-btn', disabled: pending !== null, onClick: () => run('restart') },
              tr('restart'),
            ),
            h(
              'button',
              { className: 'acb-btn', disabled: pending !== null || !running, onClick: () => run('stop') },
              tr('stop'),
            ),
          ),
        ),
        notice ? h('div', { className: 'acb-notice' }, notice) : null,
        h(
          'div',
          { className: 'acb-grid' },
          h(
            'section',
            { className: 'acb-card' },
            h('h2', { className: 'acb-card-title' }, tr('runStatus')),
            h(
              'dl',
              { className: 'acb-kv' },
              h('dt', null, tr('labelProcess')),
              h(
                'dd',
                { className: running ? 'acb-tag-ok' : 'acb-tag-idle' },
                running ? tr('valueYes') : tr('valueNo'),
              ),
              h('dt', null, tr('labelPort')),
              h('dd', { className: 'acb-mono' }, String(snapshot?.config?.httpPort ?? '—')),
              h('dt', null, tr('labelEndpoint')),
              h('dd', { className: 'acb-mono' }, snapshot?.mcpUrl ?? '—'),
              h('dt', null, tr('labelTransport')),
              h('dd', { className: 'acb-mono' }, 'streamable-http'),
              h('dt', null, tr('labelUptime')),
              h('dd', { className: 'acb-mono' }, formatUptime(snapshot?.uptimeMs)),
              h('dt', null, tr('labelLastError')),
              h(
                'dd',
                { className: snapshot?.lastError ? 'acb-tag-bad acb-mono' : 'acb-tag-idle' },
                snapshot?.lastError ?? tr('none'),
              ),
              h('dt', null, tr('labelProgId')),
              h('dd', { className: 'acb-mono' }, snapshot?.config?.progId ?? '—'),
              h('dt', null, tr('labelWatchdog')),
              h(
                'dd',
                { className: snapshot?.pendingRestart ? 'acb-tag-warn acb-mono' : 'acb-mono' },
                snapshot?.pendingRestart
                  ? tr('watchdogWaiting', { n: snapshot?.restartAttempt ?? 0 })
                  : snapshot?.restartAttempt > 0
                    ? tr('watchdogRetried', { n: snapshot.restartAttempt })
                    : tr('watchdogIdle'),
              ),
              h('dt', null, tr('labelProgIds')),
              h(
                'dd',
                { className: 'acb-mono' },
                Array.isArray(snapshot?.progIds) && snapshot.progIds.length > 0
                  ? snapshot.progIds.join(', ')
                  : tr('progIdsEmpty'),
              ),
              h('dt', null, tr('labelBridgeDir')),
              h(
                'dd',
                { className: 'acb-mono' },
                `${snapshot?.found?.bridgeDir ?? '—'}（${snapshot?.found?.bundled ? tr('bundledYes') : tr('bundledNo')}）`,
              ),
              h('dt', null, tr('labelPythonSource')),
              h(
                'dd',
                { className: 'acb-mono' },
                snapshot?.found?.pythonSource === 'venv'
                  ? tr('sourceVenv')
                  : snapshot?.found?.pythonSource === 'env'
                    ? tr('sourceEnv')
                    : tr('sourcePath'),
              ),
              h('dt', null, tr('labelStateFile')),
              h('dd', { className: 'acb-mono' }, snapshot?.found?.stateFile ?? '—'),
            ),
            h('div', { className: 'acb-notice acb-mono' }, tr('agentHint')),
          ),
          h(
            'section',
            { className: 'acb-card' },
            h('h2', { className: 'acb-card-title' }, tr('setting')),
            field('labelPython', 'pythonPath'),
            field('labelScript', 'serverScript'),
            field('labelCwd', 'workingDirectory'),
            h(
              'div',
              { className: 'acb-grid', style: { gridTemplateColumns: '1fr 1fr', gap: '12px' } },
              field('labelPortField', 'httpPort', { inputMode: 'numeric' }),
              field('labelProgId', 'progId'),
            ),
            h(
              'label',
              { className: 'acb-check' },
              h('input', {
                type: 'checkbox',
                checked: draft?.autoStart === true,
                onChange: (event) => setDraft({ ...(draft ?? {}), autoStart: event.target.checked }),
              }),
              tr('labelAutoStart'),
            ),
            h(
              'div',
              { className: 'acb-actions' },
              h(
                'button',
                { className: 'acb-btn', disabled: pending !== null || !draft, onClick: save },
                pending === 'save' ? tr('saving') : tr('save'),
              ),
              h(
                'button',
                {
                  className: 'acb-btn',
                  'data-variant': 'primary',
                  disabled: pending !== null || snapshot?.installing === true,
                  onClick: () => run('bootstrap'),
                },
                snapshot?.installing === true ? tr('bootstrapping') : tr('bootstrap'),
              ),
              h(
                'button',
                { className: 'acb-btn', disabled: pending !== null, onClick: () => run('redetect') },
                tr('redetect'),
              ),
              h(
                'button',
                { className: 'acb-btn', disabled: pending !== null, onClick: () => run('progids') },
                tr('detectProgIds'),
              ),
            ),
            h('div', { className: 'acb-notice' }, tr('installHint')),
            h('div', { className: 'acb-notice acb-mono' }, tr('portHint')),
          ),
        ),
        h(
          'section',
          { className: 'acb-card acb-logs' },
          h(
            'div',
            { className: 'acb-logs-head' },
            h('h2', { className: 'acb-card-title' }, tr('log')),
            h('span', { className: 'acb-badge' }, tr('logCount', { n: lines.length })),
          ),
          h(
            'pre',
            {
              className: 'acb-pre',
              ref: (node) => {
                bag.pre = node
              },
            },
            lines.length === 0
              ? tr('empty')
              : lines.map((line) =>
                  h(
                    'div',
                    { key: line.n, className: 'acb-line', 'data-level': line.level },
                    h('span', { className: 'acb-line-at' }, line.at),
                    h('span', { className: 'acb-line-text' }, line.text),
                  ),
                ),
          ),
        ),
      )
    }

    function BridgeIcon(props) {
      const size = props?.size ?? 20
      const active = props?.active === true
      return h(
        'svg',
        {
          width: size,
          height: size,
          viewBox: '0 0 20 20',
          fill: 'none',
          'aria-hidden': true,
          style: { color: active ? 'var(--dsw-alias-brand-primary)' : 'var(--dsw-alias-label-secondary)' },
        },
        h('path', {
          d: 'M2.5 15.5V4.5h6.2l1.6 2.2h7.2v8.8z',
          stroke: 'currentColor',
          strokeWidth: 1.5,
          strokeLinejoin: 'round',
        }),
        h('path', {
          d: 'M7.4 9.2h5.2M7.4 12.1h3.4',
          stroke: 'currentColor',
          strokeWidth: 1.5,
          strokeLinecap: 'round',
        }),
        h('circle', { cx: 15.2, cy: 12.1, r: 1.9, stroke: 'currentColor', strokeWidth: 1.5 }),
      )
    }

    return {
      inject: ['slots', 'locale'],
      apply(ctx) {
        ctx.effect(
          () =>
            ctx.locale.register(NS, { zh, en }),
          'autocad-bridge: dictionaries',
        )

        const t = ctx.locale.bind(NS)

        ctx.effect(() => {
          const tag = document.createElement('style')
          tag.dataset.autocadBridge = 'panel'
          tag.textContent = CSS
          document.head.appendChild(tag)
          return () => tag.remove()
        }, 'autocad-bridge: stylesheet')

        ctx.slots.inject('main', () =>
          ctx.slots.register(
            {
              name: 'main',
              key: PANEL_ID,
              locale: NS,
              inject: () => ({ t }),
            },
            BridgePanel,
          ),
        )

        ctx.slots.inject('sidebar.panellist', () =>
          ctx.slots.register(
            {
              name: 'sidebar.panellist',
              id: PANEL_ID,
              order: 5,
              locale: NS,
              label: () => t('panel'),
            },
            BridgeIcon,
          ),
        )
      },
    }
  },
})

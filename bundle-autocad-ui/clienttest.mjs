/**
 * Logic smoke test for the Client half.
 *
 * Loads `client.js` through a stub `window.__ModuleLoader__` with a stub `react`
 * and a stub `document`, then asserts that `apply()` performs exactly the
 * registrations the panel needs. This is NOT visual verification — it catches
 * wiring mistakes (wrong slot name, wrong registration fields, a throw inside
 * apply) that would otherwise only surface as a crashed slot entry in the browser.
 *
 * Run with: node clienttest.mjs
 */
import fs from 'node:fs'
import path from 'node:path'
import vm from 'node:vm'
import { fileURLToPath } from 'node:url'

const here = path.dirname(fileURLToPath(import.meta.url))
const source = fs.readFileSync(path.join(here, 'client.js'), 'utf8')

const calls = []
let loaded

const stubReact = {
  createElement: (type, props, ...children) => ({ type, props, children }),
  useState: (initial) => [typeof initial === 'function' ? initial() : initial, () => {}],
  useEffect: () => {},
}

const styles = []
const sandbox = {
  console,
  window: {
    __ModuleLoader__: {
      load: (definition) => {
        loaded = definition
      },
    },
    setInterval: () => 0,
    clearInterval: () => {},
  },
  document: {
    createElement: (tag) => ({ tag, dataset: {}, textContent: '', remove() {
      const at = styles.indexOf(tag)
      if (at >= 0) styles.splice(at, 1)
    } }),
    head: { appendChild: (node) => styles.push(node) },
  },
  fetch: async () => ({ ok: true, json: async () => ({}) }),
  setTimeout,
  clearTimeout,
}

vm.createContext(sandbox)
vm.runInContext(source, sandbox, { filename: 'client.js' })

const assert = (condition, message) => {
  if (!condition) {
    console.error(`FAIL ${message}`)
    process.exitCode = 1
  } else {
    console.log(`ok   ${message}`)
  }
}

assert(loaded && typeof loaded.factory === 'function', 'loader received a factory')
assert(loaded.id === '@local/autocad-bridge', `module id is ${loaded.id}`)

const module = loaded.factory((name) => {
  calls.push(['require', name])
  if (name === 'react') return stubReact
  throw new Error(`unexpected require(${name})`)
})

assert(calls.length === 1 && calls[0][1] === 'react', 'factory only requires react')
assert(Array.isArray(module.inject), 'module declares inject')
assert(module.inject.includes('slots'), 'module injects slots')
assert(module.inject.includes('locale'), 'module injects locale')
assert(typeof module.apply === 'function', 'module exposes apply')

const registrations = []
const dictionaries = []
const effects = []
const ctx = {
  effect: (factory, label) => {
    effects.push(label)
    const dispose = factory()
    return typeof dispose === 'function' ? dispose : () => {}
  },
  locale: {
    register: (ns, dicts) => {
      dictionaries.push({ ns, dicts })
      return () => {}
    },
    bind: (ns) => (key, params) => `[[${ns}:${key}]]${params ? JSON.stringify(params) : ''}`,
  },
  slots: {
    inject: (owner, callback) => {
      registrations.push(['inject', owner])
      callback()
    },
    register: (definition, component) => {
      registrations.push(['register', definition, component])
      return () => {}
    },
  },
}

module.apply(ctx)

const injected = registrations.filter((r) => r[0] === 'inject').map((r) => r[1])
const registered = registrations.filter((r) => r[0] === 'register').map((r) => r[1])

assert(injected.includes('main'), 'injected into the main slot')
assert(injected.includes('sidebar.panellist'), 'injected into sidebar.panellist')
assert(dictionaries.length === 1, 'registered one dictionary namespace')
assert(dictionaries[0].ns === 'autocad-bridge', `dictionary namespace is ${dictionaries[0].ns}`)
assert(
  dictionaries[0].dicts.zh && dictionaries[0].dicts.en,
  'dictionaries cover zh and en',
)
const zhKeys = Object.keys(dictionaries[0].dicts.zh).sort()
const enKeys = Object.keys(dictionaries[0].dicts.en).sort()
assert(zhKeys.join(',') === enKeys.join(','), 'zh and en carry identical key sets')
const nonString = (dict) => Object.entries(dict).filter(([, value]) => typeof value !== 'string')
assert(nonString(dictionaries[0].dicts.zh).length === 0, 'every zh value is a string')
assert(nonString(dictionaries[0].dicts.en).length === 0, 'every en value is a string')

const main = registered.find((r) => r.name === 'main')
assert(Boolean(main), 'registered the main panel')
assert(main.key === 'autocad-bridge', `main key is ${main.key}`)
assert(typeof main.locale === 'string', 'main registration carries a locale namespace')
assert(typeof main.inject === 'function', 'main registration carries an inject()')

const side = registered.find((r) => r.name === 'sidebar.panellist')
assert(Boolean(side), 'registered the sidebar entry')
assert(side.id === 'autocad-bridge', `sidebar id is ${side.id}`)
assert(side.id === main.key, 'sidebar id matches the main panel key')
assert(side.order === 5, `sidebar order is ${side.order}`)
assert(typeof side.label === 'function', 'sidebar label is a thunk')
assert(side.label().startsWith('[[') , `sidebar label resolves through locale (${side.label()})`)

assert(effects.length === 2, `apply registered ${effects.length} effects`)
assert(styles.length === 1, 'inserted exactly one stylesheet')
assert(styles[0].dataset.autocadBridge === 'panel', 'stylesheet is tagged for this plugin')
assert(styles[0].textContent.includes('--dsw-alias-'), 'stylesheet uses theme tokens')
assert(!/#[0-9a-fA-F]{3,8}\b/.test(styles[0].textContent), 'stylesheet carries no literal hex colours')

const panel = registered.find((r) => r.name === 'main')
assert(typeof panel !== 'undefined', 'panel component is a function')
const PanelComponent = registrations.find((r) => r[0] === 'register' && r[1].name === 'main')?.[2]
assert(typeof PanelComponent === 'function', 'main slot received a component function')
const IconComponent = registrations.find((r) => r[0] === 'register' && r[1].name === 'sidebar.panellist')?.[2]
assert(typeof IconComponent === 'function', 'sidebar slot received an icon component')

// Render the icon once; it is pure and needs no state.
const iconTree = IconComponent({ size: 22, active: true })
assert(iconTree.type === 'svg', 'icon renders an svg root')
assert(iconTree.props.width === 22 && iconTree.props.height === 22, 'icon honours the owner size')

// Render the panel once through the injected translate function; stub hooks keep
// it on the first (loading) frame, which must still be a valid element tree.
const panelTree = PanelComponent({ t: (key, params) => `[[${key}]]${params ? JSON.stringify(params) : ''}` })
assert(panelTree && panelTree.type === 'div', 'panel renders a root div')
console.log(`     panel root className = ${panelTree.props.className}`)
console.log(`     panel children      = ${panelTree.children.filter(Boolean).length}`)

process.exit(process.exitCode ?? 0)

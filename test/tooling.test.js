'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')
const { XMLParser, XMLValidator } = require('fast-xml-parser')
const { createManifest } = require('../scripts/manifest.cjs')
const { hasSecret } = require('../scripts/check-secrets.cjs')
const packageConfig = require('../package.json')
const fs = require('node:fs')
const vm = require('node:vm')
const ts = require('typescript')
const { webcrypto } = require('node:crypto')

const authCode = ts.transpileModule(fs.readFileSync(require.resolve('../web-src/auth.ts'), 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } }).outputText
const sessionKey = 'frameio.app-session'
function authClient(persistent = new Map(), tab = new Map(), blocked = {}) {
  function storage(values, unavailable) {
    const check = () => { if (unavailable) throw new Error('Storage blocked') }
    return { getItem: key => { check(); return values.get(key) || null }, setItem: (key, value) => { check(); values.set(key, value) }, removeItem: key => { check(); values.delete(key) } }
  }
  const exports = {}
  vm.runInNewContext(authCode, { exports, require: () => ({}), crypto: webcrypto, localStorage: storage(persistent, blocked.persistent), sessionStorage: storage(tab, blocked.tab) })
  return exports
}

test('hosted OAuth uses the pinned same-origin Stage proxy and rejects other origins', async () => {
  const publicConfig = require('../web-src/public-config.json')
  const origin = 'https://27200-msofficetoframe-stage.adobeio-static.net'
  const location = { hostname: new URL(origin).hostname, origin, href: origin + '/index.html' }
  const exports = {}
  const calls = []
  vm.runInNewContext(authCode, {
    exports, require: () => ({ default: publicConfig }), location, URL, AbortSignal,
    localStorage: { getItem: () => 'ab'.repeat(32) }, sessionStorage: { removeItem: () => {} },
    fetch: async (url, options) => {
      calls.push({ url: url.href, credentials: options.credentials })
      return Response.json({ authenticated: false })
    }
  })
  await exports.oauthRequest('session')
  assert.equal(calls[0].url, origin + '/api/v1/web/msofficetoframe/oauth')
  assert.equal(calls[0].credentials, 'omit')
  location.origin = 'https://untrusted.invalid'
  await assert.rejects(exports.oauthRequest('session'), /Invalid authentication endpoint/)
  assert.equal(calls.length, 1)
})

test('hosted manifest accepts a cache version without allowing arbitrary query parameters', () => {
  const origin = 'https://27200-msofficetoframe-stage.adobeio-static.net'
  const url = origin + '/index.html?v=20261013-stage'
  const manifest = new XMLParser({ ignoreAttributes: false }).parse(createManifest(url)).OfficeApp
  assert.equal(manifest.DefaultSettings.SourceLocation['@_DefaultValue'], url)
  assert.equal(manifest.IconUrl['@_DefaultValue'], origin + '/assets/icon-32.png?v=frameio')
  assert.throws(() => createManifest(origin + '/index.html?redirect=elsewhere'))
  assert.throws(() => createManifest(url + '&redirect=elsewhere'))
})

test('app session persists across fresh panels and sign-out clears both stores', () => {
  const persistent = new Map()
  const first = authClient(persistent).appSession()
  assert.match(first, /^[a-f0-9]{64}$/)
  assert.equal(authClient(persistent, new Map()).appSession(), first)
  const legacy = new Map([[sessionKey, 'ab'.repeat(32)]])
  authClient(persistent, legacy).clearAppSession()
  assert.equal(persistent.has(sessionKey), false)
  assert.equal(legacy.has(sessionKey), false)
  assert.notEqual(authClient(persistent).appSession(), first)
})

test('existing tab session migrates without requiring another Adobe login', () => {
  const persistent = new Map()
  const legacy = new Map([[sessionKey, 'cd'.repeat(32)]])
  assert.equal(authClient(persistent, legacy).appSession(), 'cd'.repeat(32))
  assert.equal(persistent.get(sessionKey), 'cd'.repeat(32))
  assert.equal(legacy.has(sessionKey), false)
  assert.equal(authClient(persistent).appSession(), 'cd'.repeat(32))
})

test('persistent session is authoritative and malformed stored keys are replaced', () => {
  const persistent = new Map([[sessionKey, 'ef'.repeat(32)]])
  const legacy = new Map([[sessionKey, 'ab'.repeat(32)]])
  assert.equal(authClient(persistent, legacy).appSession(), 'ef'.repeat(32))
  assert.equal(legacy.has(sessionKey), false)
  persistent.set(sessionKey, 'invalid')
  legacy.set(sessionKey, 'invalid')
  const value = authClient(persistent, legacy).appSession()
  assert.match(value, /^[a-f0-9]{64}$/)
  assert.equal(persistent.get(sessionKey), value)
})

test('blocked persistent storage falls back to tab storage without generating changing keys', () => {
  const legacy = new Map()
  const client = authClient(new Map(), legacy, { persistent: true })
  const value = client.appSession()
  assert.equal(client.appSession(), value)
  client.clearAppSession()
  assert.equal(legacy.has(sessionKey), false)
  assert.throws(() => authClient(new Map(), new Map(), { persistent: true, tab: true }).appSession(), /session storage is unavailable/)
})

test('connection menu survives native null-target blur and activates the sign-in click', () => {
  const source = ts.createSourceFile('index.tsx', fs.readFileSync(require.resolve('../web-src/index.tsx'), 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX)
  const component = source.statements.find(node => ts.isFunctionDeclaration(node) && node.name?.text === 'ConnectionMenu')
  assert.ok(component)
  const module = { exports: {} }
  const changes = []
  let focused = 0
  let connects = 0
  const element = (type, props) => ({ type, props })
  const code = ts.transpileModule(component.getText(source) + '\nmodule.exports = ConnectionMenu;', { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022, jsx: ts.JsxEmit.ReactJSX } }).outputText
  vm.runInNewContext(code, {
    module, exports: module.exports, require: () => ({ jsx: element, jsxs: element }),
    useState: () => [true, value => changes.push(value)],
    useRef: () => ({ current: { focus: () => { focused++ } } }), useEffect: () => {},
    ChevronDown: 'icon', CircleX: 'icon', LogIn: 'icon'
  })
  const tree = module.exports({ account: 'signed-out', disabled: false, connect: () => { connects++ }, disconnect: () => {}, cancel: () => {} })
  tree.props.onBlur({ relatedTarget: null, currentTarget: { contains: () => false } })
  assert.equal(changes.length, 0)
  const option = tree.props.children[1].props.children
  let prevented = false
  option.props.onMouseDown({ preventDefault: () => { prevented = true } })
  assert.equal(prevented, true)
  option.props.onClick()
  assert.equal(connects, 1)
  assert.equal(focused, 1)
  assert.equal(changes[0], false)
  tree.props.onBlur({ relatedTarget: {}, currentTarget: { contains: () => false } })
  assert.equal(changes.at(-1), false)
})

test('frontend dependencies are bundled for a browser without environment inlining', () => {
  assert.equal(packageConfig.targets.default.context, 'browser')
  assert.equal(packageConfig.targets.default.includeNodeModules, true)
  assert.equal(packageConfig['@parcel/transformer-js'].inlineEnvironment, false)
})

test('manifest is PowerPoint-only, permits metadata writes and uses explicit HTTPS paths', () => {
  const xml = createManifest('https://localhost:3000/index.html')
  assert.equal(XMLValidator.validate(xml), true)
  const manifest = new XMLParser({ ignoreAttributes: false }).parse(xml).OfficeApp
  assert.equal(manifest.Permissions, 'ReadWriteDocument')
  assert.equal(manifest.DisplayName['@_DefaultValue'], 'Frame.io')
  assert.equal(manifest.Hosts.Host['@_Name'], 'Presentation')
  assert.equal(manifest.DefaultSettings.SourceLocation['@_DefaultValue'], 'https://localhost:3000/index.html')
  assert.equal(manifest.IconUrl['@_DefaultValue'], 'https://localhost:3000/assets/icon-32.png?v=frameio')
})

test('manifest rejects insecure, placeholder and credential-bearing URLs', () => {
  for (const url of ['http://localhost:3000/index.html', 'https://example.com/index.html', 'https://user:secret@localhost/index.html', 'https://localhost/index.html?token=secret']) {
    assert.throws(() => createManifest(url))
  }
})

test('custom-domain manifest keeps the task pane and assets on the chosen origin', () => {
  const xml = createManifest('https://www.localappbuilder.com:3000/index.html')
  const manifest = new XMLParser({ ignoreAttributes: false }).parse(xml).OfficeApp
  assert.equal(manifest.DefaultSettings.SourceLocation['@_DefaultValue'], 'https://www.localappbuilder.com:3000/index.html')
  assert.equal(manifest.IconUrl['@_DefaultValue'], 'https://www.localappbuilder.com:3000/assets/icon-32.png?v=frameio')
})

test('secret scanner detects raw and encoded values without printing them', () => {
  const secret = 'fake-secret:only-for-testing'
  for (const content of [secret, encodeURIComponent(secret), Buffer.from(secret).toString('base64')]) {
    assert.equal(hasSecret(Buffer.from(content), [secret]), true)
  }
  assert.equal(hasSecret('public health information', [secret]), false)
})
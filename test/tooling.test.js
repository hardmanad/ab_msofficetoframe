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
const { spawnSync } = require('node:child_process')

const authCode = ts.transpileModule(fs.readFileSync(require.resolve('../web-src/auth.ts'), 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } }).outputText
const sessionKey = 'frameio.app-session'
test('Spectrum panel uses the official font provider and isolated Parcel-bundled styles', () => {
  assert.ok(packageConfig.dependencies['@react-spectrum/s2'])
  assert.equal(packageConfig['@parcel/resolver-default'].packageExports, true)
  const bundle = packageConfig['@parcel/bundler-default'].manualSharedBundles.find(bundle => bundle.name === 's2-styles')
  assert.deepEqual(bundle.types, ['css'])
  assert.ok(bundle.assets.includes('**/@react-spectrum/s2/**'))
  assert.ok(bundle.assets.includes('web-src/**/*.{js,jsx,ts,tsx}'))
  const source = fs.readFileSync(require.resolve('../web-src/index.tsx'), 'utf8')
  assert.match(source, /import \{ Provider \} from '@react-spectrum\/s2\/Provider'/)
  assert.match(source, /<Provider colorScheme="light" background="base"/)
  assert.match(source, /fontFamily: 'sans'/)
  const stylesheet = fs.readFileSync(require.resolve('../web-src/styles.css'), 'utf8')
  assert.match(stylesheet, /^@layer panel, _;\s*@layer panel \{/)
  assert.match(stylesheet, /font-family: 'adobe-clean-spectrum-vf'/)
  assert.doesNotThrow(() => require('lightningcss').transform({ filename: 'styles.css', code: Buffer.from(stylesheet) }))
})

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

test('local OAuth uses only the HTTPS localhost port 3000 origin', async () => {
  const origin = 'https://localhost:3000'
  const location = { hostname: 'localhost', origin, href: origin + '/index.html' }
  const exports = {}
  const calls = []
  vm.runInNewContext(authCode, {
    exports, require: () => ({ default: require('../web-src/public-config.json') }), location, URL, AbortSignal,
    localStorage: { getItem: () => 'ab'.repeat(32) },
    fetch: async (url, options) => {
      calls.push({ url: url.href, credentials: options.credentials })
      return Response.json({ authenticated: false })
    }
  })
  await exports.oauthRequest('session')
  assert.equal(calls[0].url, origin + '/api/oauth')
  assert.equal(calls[0].credentials, 'omit')
  for (const origin of ['https://localhost:3002', 'http://localhost:3000', 'https://www.localappbuilder.com:3000', 'https://localappbuilder:3000']) {
    Object.assign(location, { hostname: new URL(origin).hostname, origin, href: origin + '/index.html' })
    await assert.rejects(exports.oauthRequest('session'), /Invalid authentication endpoint/)
  }
  assert.equal(calls.length, 1)
})

test('development server rejects alternate ports without suggesting a fallback', () => {
  for (const port of ['3001', '3002']) {
    const result = spawnSync(process.execPath, [require.resolve('../scripts/dev.cjs'), '--port', port], { encoding: 'utf8' })
    assert.equal(result.status, 1)
    assert.match(result.stderr, /Local development requires port 3000 on both the container and host/)
  }
  const settings = JSON.parse(fs.readFileSync(require.resolve('../.vscode/settings.json'), 'utf8'))
  assert.equal(settings['remote.portsAttributes']['3000'].requireLocalPort, true)
  assert.equal(settings['remote.portsAttributes']['3000'].protocol, 'https')
})

test('deployment and public OAuth endpoints remain pinned to hosted Stage', () => {
  const origin = 'https://27200-msofficetoframe-stage.adobeio-static.net'
  const config = fs.readFileSync(require.resolve('../app.config.yaml'), 'utf8')
  assert.match(config, /FRAMEIO_REDIRECT_URI: https:\/\/27200-msofficetoframe-stage\.adobeio-static\.net\/oauth-callback\.html/)
  assert.doesNotMatch(config, /localhost|localappbuilder|3002/)
  const publicConfig = require('../web-src/public-config.json')
  assert.equal(publicConfig.oauthUrl, origin + '/api/v1/web/msofficetoframe/oauth')
  assert.equal(publicConfig.statusUrl, origin + '/api/v1/web/msofficetoframe/status')
})

test('OAuth configuration defaults to Stage and preserves explicit local OAuth without trusting historical export redirects', () => {
  const script = fs.readFileSync(require.resolve('../scripts/configure-oauth.cjs'), 'utf8')
  const stage = 'https://27200-msofficetoframe-stage.adobeio-static.net/oauth-callback.html'
  for (const callback of [undefined, stage, 'https://localhost:3000/oauth-callback.html', 'https://www.localappbuilder.com:3000/oauth-callback.html']) {
    let written
    const process = {}
    const errors = []
    const files = {
      existsSync: () => true,
      readFileSync: name => name === '.env' ? (callback ? `FRAMEIO_REDIRECT_URI=${callback}\n` : '')
        : JSON.stringify({ project: { workspace: { details: { credentials: [{ oauth2: { client_id: 'test-client', client_secret: 'test-secret', defaultRedirectUri: 'https://www.localappbuilder.com:3000/oauth-callback.html' } }] } } } }),
      writeFileSync: (_name, content) => { written = content },
      chmodSync: () => {}
    }
    vm.runInNewContext(script, {
      require: name => name === 'node:fs' ? files : require(name), process,
      console: { log: () => {}, error: message => errors.push(message) }
    })
    if (callback?.includes('www.localappbuilder.com')) {
      assert.equal(process.exitCode, 1)
      assert.equal(written, undefined)
      assert.match(errors[0], /hosted Stage callback/)
    } else {
      assert.equal(process.exitCode, undefined)
      assert.equal(require('dotenv').parse(written).FRAMEIO_REDIRECT_URI, callback || stage)
    }
  }
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

test('authentication reports HTTP errors instead of native JSON parsing errors', async () => {
  const publicConfig = require('../web-src/public-config.json')
  const origin = 'https://27200-msofficetoframe-stage.adobeio-static.net'
  const exports = {}
  let response
  vm.runInNewContext(authCode, {
    exports, require: () => ({ default: publicConfig }), URL, AbortSignal,
    location: { hostname: new URL(origin).hostname, origin, href: origin + '/index.html' },
    localStorage: { getItem: () => 'ab'.repeat(32) }, sessionStorage: { removeItem: () => {} },
    fetch: async () => response
  })
  response = { ok: false, status: 503, json: async () => { throw new SyntaxError('The string did not match the expected pattern.') } }
  await assert.rejects(exports.oauthRequest('session'), { status: 503, message: 'Authentication service returned HTTP 503 instead of JSON.' })
  response = new Response('<!DOCTYPE html>', { status: 200 })
  await assert.rejects(exports.oauthRequest('session'), { status: 502, message: 'Authentication service returned an invalid response.' })
  response = Response.json({ error: 'Session expired.' }, { status: 401 })
  await assert.rejects(exports.oauthRequest('session'), { status: 401, message: 'Session expired.' })
  response = Response.json({ error: 'Asset missing.', relinkAllowed: true }, { status: 404 })
  await assert.rejects(exports.oauthRequest('presentation-lookup'), { status: 404, relinkAllowed: true })
  response = Response.json({ error: 'Asset missing.' }, { status: 404 })
  await assert.rejects(exports.oauthRequest('presentation-lookup'), { status: 404, relinkAllowed: false })
  response = Response.json({ error: 'Denied.', relinkAllowed: true }, { status: 403 })
  await assert.rejects(exports.oauthRequest('presentation-lookup'), { status: 403, relinkAllowed: false })
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

test('Word App uses Office dialogs even without export and renders no slide preview', async () => {
  const source = ts.createSourceFile('index.tsx', fs.readFileSync(require.resolve('../web-src/index.tsx'), 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX)
  const component = source.statements.find(node => ts.isFunctionDeclaration(node) && node.name?.text === 'App')
  const transformed = ts.transform(component, [context => node => ts.visitNode(node, function visit(child) {
    return ts.isMetaProperty(child) ? ts.factory.createIdentifier('moduleMeta') : ts.visitEachChild(child, visit, context)
  })])
  const code = ts.transpileModule(ts.createPrinter().printNode(ts.EmitHint.Unspecified, transformed.transformed[0], source) + '\nmodule.exports = App;', { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022, jsx: ts.JsxEmit.ReactJSX } }).outputText
  transformed.dispose()
  for (const available of [true, false]) {
    const module = { exports: {} }
    let stateIndex = 0
    const effects = [], dialogs = [], changes = []
    const adapter = { kind: 'word', noun: 'document', filename: () => 'Report.docx', ensureId: async () => 'aaaaaaaa-1234-1234-1234-123456789abc' }
    const host = { kind: 'word', adapter, inOffice: true, available, reason: available ? '' : 'DOCX uploads require Word on Mac or Windows desktop.' }
    const element = (type, props) => ({ type, props })
    vm.runInNewContext(code, {
      module, exports: module.exports, URL, moduleMeta: { url: 'https://localhost:3000/index.html' }, AbortController,
      require: () => ({ jsx: element, jsxs: element }),
      useState: initial => { const index = stateIndex++; return [index === 0 ? host : index === 1 ? 'signed-out' : initial, value => changes.push(value)] },
      useRef: () => ({ current: null }), useEffect: callback => effects.push(callback),
      signIn: async inOffice => dialogs.push(inOffice), clearAppSession: () => {},
      ConnectionMenu: 'ConnectionMenu', UploadPanel: 'UploadPanel', DestinationPicker: 'DestinationPicker', FileText: 'FileText', LoaderCircle: 'LoaderCircle', ImageOff: 'ImageOff'
    })
    const tree = module.exports()
    const elements = value => !value || typeof value !== 'object' ? [] : Array.isArray(value) ? value.flatMap(elements) : [value, ...elements(value.props?.children)]
    const nodes = elements(tree)
    assert.equal(nodes.some(node => node.props?.className === 'slide-preview'), false)
    assert.equal(nodes.some(node => node.type === 'ImageOff'), false)
    assert.equal(nodes.find(node => node.type === 'UploadPanel').props.adapter, adapter)
    assert.equal(nodes.find(node => node.type === 'UploadPanel').props.available, available)
    assert.equal(nodes.find(node => node.type === 'header').props.children[1].props.children[1].props.children.join(''), 'Word / Frame.io')
    effects[2]()
    assert.equal(changes.at(-1).name, 'Report.docx')
    assert.equal(changes.at(-1).loading, false)
    await nodes.find(node => node.type === 'ConnectionMenu').props.connect()
    assert.deepEqual(dialogs, [true])
  }
})

test('disposable Word documents and Office lock files remain ignored', () => {
  const patterns = fs.readFileSync(require.resolve('../.gitignore'), 'utf8').split(/\r?\n/)
  assert.ok(patterns.includes('*.docx'))
  assert.ok(patterns.includes('~$*'))
})

function saveReminderEffect(isSaved, initial = { id: 'document-id', created: true, checking: false, error: '' }) {
  const source = ts.createSourceFile('index.tsx', fs.readFileSync(require.resolve('../web-src/index.tsx'), 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX)
  const app = source.statements.find(node => ts.isFunctionDeclaration(node) && node.name?.text === 'App')
  const effect = app.body.statements.filter(ts.isExpressionStatement).map(node => node.expression)
    .find(node => ts.isCallExpression(node) && node.expression.getText(source) === 'useEffect' && node.arguments[0].getText(source).includes('checkSaved'))
  assert.ok(effect)
  const code = ts.transpileModule('const effect = ' + effect.arguments[0].getText(source) + '; module.exports = effect;', { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } }).outputText
  const module = { exports: {} }, timers = new Map()
  let identity = initial, updates = 0, nextTimer = 0
  vm.runInNewContext(code, {
    module, host: { adapter: { isSaved } }, identity: initial,
    setIdentity: update => { identity = update(identity); updates++ },
    window: { setTimeout: (callback, milliseconds) => { assert.equal(milliseconds, 2000); timers.set(++nextTimer, callback); return nextTimer }, clearTimeout: timer => timers.delete(timer) }
  })
  return { cleanup: module.exports(), timers, identity: () => identity, updates: () => updates,
    replaceIdentity: next => { identity = next }, settle: () => new Promise(setImmediate),
    tick: async () => { const [timer, callback] = timers.entries().next().value; timers.delete(timer); await callback() } }
}

test('Word save reminder clears after confirmed save and stops checking', async () => {
  let checks = 0
  const client = saveReminderEffect(async () => ++checks > 1)
  await client.settle()
  assert.equal(client.identity().created, true)
  assert.equal(client.timers.size, 1)
  await client.tick()
  assert.equal(client.identity().created, false)
  assert.equal(client.identity().id, 'document-id')
  assert.equal(client.updates(), 1)
  assert.equal(client.timers.size, 0)
  client.cleanup()
  assert.equal(checks, 2)
})

test('save reminder keeps failed or unknown saves visible and cleans up pending checks', async () => {
  let checks = 0
  const client = saveReminderEffect(async () => { if (++checks === 1) throw new Error('Save state unavailable'); return false })
  await client.settle()
  assert.equal(client.updates(), 0)
  await client.tick()
  assert.equal(client.identity().created, true)
  assert.equal(client.timers.size, 1)
  client.cleanup()
  assert.equal(client.timers.size, 0)
  for (const saved of [true, false]) {
    let finish
    const pending = saveReminderEffect(() => new Promise(resolve => { finish = resolve }))
    assert.equal(pending.timers.size, 0)
    pending.cleanup()
    finish(saved)
    await pending.settle()
    assert.equal(pending.updates(), 0)
    assert.equal(pending.timers.size, 0)
  }
})

test('save monitoring does not affect PowerPoint, existing IDs or a replaced identity', async () => {
  const powerpoint = saveReminderEffect(undefined)
  assert.equal(powerpoint.cleanup, undefined)
  assert.equal(powerpoint.timers.size, 0)
  const existing = saveReminderEffect(() => { throw new Error('Existing UUID must not be monitored') }, { id: 'existing', created: false })
  assert.equal(existing.cleanup, undefined)
  let finish
  const client = saveReminderEffect(() => new Promise(resolve => { finish = resolve }))
  client.replaceIdentity({ id: 'replacement', created: true })
  finish(true)
  await client.settle()
  assert.equal(client.identity().id, 'replacement')
  assert.equal(client.identity().created, true)
  client.cleanup()
})

test('local development upload allowlist accepts the additive document type field', () => {
  const source = ts.createSourceFile('dev.cjs', fs.readFileSync(require.resolve('../scripts/dev.cjs'), 'utf8'), ts.ScriptTarget.Latest, true)
  let fields
  function visit(node) {
    if (ts.isArrayLiteralExpression(node)) {
      const values = node.elements.filter(ts.isStringLiteral).map(element => element.text)
      if (values.includes('presentationId')) fields = values
    }
    ts.forEachChild(node, visit)
  }
  visit(source)
  assert.ok(fields.includes('documentType'))
  assert.equal(fields.includes('FRAMEIO_CLIENT_SECRET'), false)
})

test('frontend dependencies are bundled for a browser without environment inlining', () => {
  assert.equal(packageConfig.targets.default.context, 'browser')
  assert.equal(packageConfig.targets.default.includeNodeModules, true)
  assert.equal(packageConfig['@parcel/transformer-js'].inlineEnvironment, false)
})

test('manifest supports only PowerPoint and Word, permits metadata writes and uses explicit HTTPS paths', () => {
  const xml = createManifest('https://localhost:3000/index.html')
  assert.equal(XMLValidator.validate(xml), true)
  const manifest = new XMLParser({ ignoreAttributes: false }).parse(xml).OfficeApp
  assert.equal(manifest.Permissions, 'ReadWriteDocument')
  assert.equal(manifest.DisplayName['@_DefaultValue'], 'Frame.io')
  assert.deepEqual(manifest.Hosts.Host.map(host => host['@_Name']), ['Presentation', 'Document'])
  assert.equal(manifest.Id, '4f391fa8-6ea8-4a32-91d2-3b298c24104e')
  assert.equal(manifest.DefaultSettings.SourceLocation['@_DefaultValue'], 'https://localhost:3000/index.html')
  assert.equal(manifest.IconUrl['@_DefaultValue'], 'https://localhost:3000/assets/icon-32.png?v=frameio')
})

test('manifest rejects insecure, placeholder and credential-bearing URLs', () => {
  for (const url of ['http://localhost:3000/index.html', 'https://example.com/index.html', 'https://user:secret@localhost/index.html', 'https://localhost/index.html?token=secret']) {
    assert.throws(() => createManifest(url))
  }
})

test('local manifest command keeps the task pane and assets on localhost port 3000', () => {
  assert.equal(packageConfig.scripts['manifest:local'], 'node scripts/manifest.cjs --url https://localhost:3000/index.html --output office/manifest.local.xml')
  const xml = createManifest('https://localhost:3000/index.html')
  const manifest = new XMLParser({ ignoreAttributes: false }).parse(xml).OfficeApp
  assert.equal(manifest.DefaultSettings.SourceLocation['@_DefaultValue'], 'https://localhost:3000/index.html')
  assert.equal(manifest.IconUrl['@_DefaultValue'], 'https://localhost:3000/assets/icon-32.png?v=frameio')
})

test('secret scanner detects raw and encoded values without printing them', () => {
  const secret = 'fake-secret:only-for-testing'
  for (const content of [secret, encodeURIComponent(secret), Buffer.from(secret).toString('base64')]) {
    assert.equal(hasSecret(Buffer.from(content), [secret]), true)
  }
  assert.equal(hasSecret('public health information', [secret]), false)
})
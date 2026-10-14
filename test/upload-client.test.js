'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')
const ts = require('typescript')
const fs = require('node:fs')
const vm = require('node:vm')
const exportsObject = {}
const runtime = { exports: exportsObject, AbortController, AbortSignal, URL, Uint8Array, setTimeout, clearTimeout, Promise, console }
vm.runInNewContext(ts.transpileModule(fs.readFileSync(require.resolve('../web-src/upload-transfer.ts'), 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } }).outputText, runtime)
const { transferParts, TransferError, validateParts, openPresentation, putPart } = exportsObject
const identityExports = {}
const identityRuntime = { exports: identityExports, URL, Office: { context: { requirements: { isSetSupported: () => true } } } }
vm.runInNewContext(ts.transpileModule(fs.readFileSync(require.resolve('../web-src/presentation.ts'), 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } }).outputText, identityRuntime)
const parts = [3, 4, 2].map((size, index) => ({ size, url: `https://frameio-test.s3-accelerate.amazonaws.com/${index}` }))

function wordClient({ stored, supported = true, saved = false, failRead = false, failWrite = false } = {}) {
  const exports = {}
  let writes = 0
  let runs = 0
  let syncs = 0
  const runtime = { exports, URL, setTimeout, clearTimeout, Office: { AsyncResultStatus: { Succeeded: 'succeeded' }, context: { requirements: { isSetSupported: (name, version) => supported && name === 'WordApi' && ['1.1', '1.3'].includes(version) }, document: { url: '' } } }, Word: { run: async callback => {
    runs++
    return callback({ sync: async () => {
      syncs++
      if (failRead || (failWrite && writes)) throw Object.assign(new Error('Metadata denied'), { code: 'AccessDenied' })
    }, document: { saved, load: property => assert.equal(property, 'saved'), properties: { customProperties: {
      getItemOrNullObject: key => { assert.equal(key, 'MSOFFICE_TO_FRAME_UUID'); return { load: () => {}, isNullObject: stored === undefined, value: stored } },
      add: (key, value) => { assert.equal(key, 'MSOFFICE_TO_FRAME_UUID'); writes++; stored = value }
    } } } })
  } } }
  vm.runInNewContext(ts.transpileModule(fs.readFileSync(require.resolve('../web-src/word.ts'), 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } }).outputText, runtime)
  return { exports, runtime, counts: () => ({ writes, runs, syncs }) }
}

test('Word identity is persisted once and concurrent pane requests share initialization', async () => {
  const client = wordClient()
  const expected = 'AAAAAAAA-1234-1234-1234-123456789ABC'
  const first = client.exports.ensureWordDocumentId(() => expected)
  const second = client.exports.ensureWordDocumentId(() => { throw new Error('Duplicate generation') })
  assert.deepEqual(await Promise.all([first, second]), [expected.toLowerCase(), expected.toLowerCase()])
  assert.equal(await client.exports.ensureWordDocumentId(() => { throw new Error('Do not generate again') }), expected.toLowerCase())
  assert.equal(client.counts().writes, 1)
  assert.equal(await wordClient({ stored: expected }).exports.ensureWordDocumentId(), expected.toLowerCase())
})

test('Word identity rejects corrupted metadata, unsupported APIs and failed writes', async () => {
  for (const stored of ['invalid', 123, null]) {
    const client = wordClient({ stored })
    await assert.rejects(client.exports.ensureWordDocumentId(() => { throw new Error('Do not replace') }), /invalid tracking ID/)
    assert.equal(client.counts().writes, 0)
  }
  const unsupported = wordClient({ supported: false })
  await assert.rejects(unsupported.exports.ensureWordDocumentId(), /WordApi 1.3/)
  assert.equal(unsupported.counts().runs, 0)
  await assert.rejects(wordClient().exports.ensureWordDocumentId(() => 'invalid'), /could not be generated/)
  for (const options of [{ failRead: true }, { failWrite: true }]) await assert.rejects(wordClient(options).exports.ensureWordDocumentId(() => 'aaaaaaaa-1234-1234-1234-123456789abc'), { code: 'AccessDenied' })
})

test('Word filename reads decoded document URLs and handles unsaved documents', () => {
  const client = wordClient()
  assert.equal(client.exports.wordFilename(), 'Untitled document')
  client.runtime.Office.context.document.url = 'https://files.example/Report%20%C3%A9.docx?download=true'
  assert.equal(client.exports.wordFilename(), 'Report \u00e9.docx')
  for (const url of ['/Users/josh/Documents/My File.docx', 'C:\\Users\\josh\\Documents\\My File.docx', 'file:///Users/josh/Documents/My%20File.docx']) {
    client.runtime.Office.context.document.url = url
    assert.equal(client.exports.wordFilename(), 'My File.docx')
  }
  client.runtime.Office.context.document.url = 'not a URL'
  assert.equal(client.exports.wordFilename(), 'Untitled document')
})

test('Word saved-state checks load and sync without writing metadata or forcing a save', async () => {
  for (const saved of [false, true, undefined, 'true']) {
    const client = wordClient({ saved })
    assert.equal(await client.exports.wordDocumentSaved(), saved === true)
    assert.equal(client.counts().syncs, 1)
    assert.equal(client.counts().writes, 0)
  }
  const unsupported = wordClient({ supported: false })
  assert.equal(await unsupported.exports.wordDocumentSaved(), false)
  assert.equal(unsupported.counts().runs, 0)
  await assert.rejects(wordClient({ failRead: true }).exports.wordDocumentSaved(), { code: 'AccessDenied' })
})

test('Word loads the filename from Office file properties when the cached URL is empty', async () => {
  const client = wordClient()
  client.runtime.Office.context.document.getFilePropertiesAsync = callback => callback({ status: 'succeeded', value: { url: '/Users/josh/Documents/My File.docx' } })
  await client.exports.loadWordFileProperties()
  assert.equal(client.exports.wordFilename(), 'My File.docx')
  assert.equal(client.exports.wordFileUrl().pathname, '/Users/josh/Documents/My%20File.docx')
  assert.equal(client.counts().writes, 0)
  for (const result of [{ status: 'failed' }, { status: 'succeeded', value: { url: '' } }]) {
    const fallback = wordClient()
    fallback.runtime.Office.context.document.url = 'file:///Users/josh/Documents/Existing.docx'
    fallback.runtime.Office.context.document.getFilePropertiesAsync = callback => callback(result)
    await fallback.exports.loadWordFileProperties()
    assert.equal(fallback.exports.wordFilename(), 'Existing.docx')
  }
  const unsupported = wordClient()
  await unsupported.exports.loadWordFileProperties()
  assert.equal(unsupported.exports.wordFilename(), 'Untitled document')
})

function documentClient() {
  const exports = {}
  const expected = 'aaaaaaaa-1234-1234-1234-123456789abc'
  const calls = []
  const runtime = { exports, URL, Error, require: name => name === './word' ? {
    wordFilename: () => 'Report.docx', wordFileUrl: () => { try { return new URL(runtime.Office.context.document.url) } catch { return null } },
    loadWordFileProperties: async () => {}, ensureWordDocumentId: async () => { calls.push('word'); return expected }
  } : {
    presentationFilename: () => 'Deck.pptx', ensurePresentationId: async () => { calls.push('powerpoint'); return expected }, firstSlidePreview: async () => { calls.push('preview'); return null }
  }, Office: { HostType: { Word: 'Word', PowerPoint: 'PowerPoint', Excel: 'Excel' }, PlatformType: { PC: 'PC', Mac: 'Mac', OfficeOnline: 'OfficeOnline' }, context: { document: { url: '' }, requirements: { isSetSupported: () => true } } } }
  vm.runInNewContext(ts.transpileModule(fs.readFileSync(require.resolve('../web-src/office-document.ts'), 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } }).outputText, runtime)
  return { exports, runtime, calls, expected }
}

test('document adapters enable Word desktop only and keep Office dialog routing independent of export', async () => {
  const { exports, runtime, calls } = documentClient()
  for (const platform of ['Mac', 'PC']) {
    const host = exports.getOfficeHost({ host: 'Word', platform })
    assert.equal(host.available, true)
    assert.equal(host.inOffice, true)
    assert.equal(host.adapter.documentType, 'docx')
    assert.equal(host.adapter.preview, undefined)
    await host.adapter.ensureId()
  }
  for (const platform of ['OfficeOnline', 'iOS', null]) {
    const host = exports.getOfficeHost({ host: 'Word', platform })
    assert.equal(host.available, false)
    assert.equal(host.inOffice, true)
    assert.match(host.reason, /Mac or Windows/)
  }
  assert.deepEqual(calls, ['word', 'word'])
  assert.equal(exports.getOfficeHost({ host: 'Excel', platform: 'Mac' }).adapter, null)
  assert.equal(exports.getOfficeHost({ host: null, platform: null }).inOffice, false)
  runtime.Office.context.requirements.isSetSupported = name => name !== 'WordApi'
  assert.match(exports.getOfficeHost({ host: 'Word', platform: 'Mac' }).reason, /WordApi 1.3/)
  runtime.Office.context.requirements.isSetSupported = () => false
  assert.match(exports.getOfficeHost({ host: 'Word', platform: 'Mac' }).reason, /file export/)
  delete runtime.Office
  assert.equal(exports.getOfficeHost({ host: null, platform: null }).kind, 'browser')
})

test('adapters validate filenames and source formats without changing document identity', async () => {
  const { exports, runtime, calls, expected } = documentClient()
  const word = exports.getOfficeHost({ host: 'Word', platform: 'Mac' }).adapter
  const powerpoint = exports.getOfficeHost({ host: 'PowerPoint', platform: 'Mac' }).adapter
  assert.equal(exports.initialUploadName(word), 'Report.docx')
  assert.equal(exports.initialUploadName({ ...word, filename: () => 'Untitled document' }), 'Document.docx')
  for (const name of ['Report.DOCX', 'Report \u00e9.docx', 'a.docx']) assert.equal(exports.validUploadName(name, 'docx'), true)
  for (const name of ['.docx', '../Report.docx', 'Report.docm', 'Report.pptx', 'Report\u0000.docx', 'a'.repeat(176) + '.docx']) assert.equal(exports.validUploadName(name, 'docx'), false)
  await exports.verifyDocumentIdentity(word, expected)
  await exports.verifyDocumentIdentity(powerpoint, expected)
  assert.deepEqual(calls, ['word', 'powerpoint'])
  await assert.rejects(exports.verifyDocumentIdentity(word, 'bbbbbbbb-1234-1234-1234-123456789abc'), /tracking ID changed/)
  await assert.rejects(exports.verifyDocumentIdentity({ ...word, ensureId: async generate => generate() }, expected), /tracking ID was removed/)
  for (const extension of ['doc', 'docm', 'dotx', 'pptx']) {
    runtime.Office.context.document.url = 'https://files.example/Report.' + extension
    assert.equal(exports.getOfficeHost({ host: 'Word', platform: 'PC' }).available, false)
    await assert.rejects(exports.verifyDocumentIdentity(word, expected), /not supported/)
  }
})

test('presentation identity is created once in metadata and reused on reopen', async () => {
  let stored
  let writes = 0
  identityRuntime.PowerPoint = { run: async callback => callback({ sync: async () => {}, presentation: { tags: {
    getItemOrNullObject: key => { assert.equal(key, identityExports.presentationTag); return { load: () => {}, isNullObject: !stored, value: stored } },
    add: (key, value) => { writes++; stored = value }
  } } }) }
  const expected = '12345678-1234-1234-1234-123456789abc'
  assert.equal(await identityExports.ensurePresentationId(() => expected), expected)
  assert.equal(await identityExports.ensurePresentationId(() => { throw new Error('Do not generate again') }), expected)
  assert.equal(writes, 1)
  stored = 'invalid'
  await assert.rejects(identityExports.ensurePresentationId(() => expected), /invalid tracking ID/)
  assert.equal(writes, 1)
})

test('first-slide preview respects capability and empty presentations and requests only slide zero', async () => {
  const saved = identityRuntime.PowerPoint
  try {
    identityRuntime.Office.context.requirements.isSetSupported = (_, version) => version !== '1.8'
    identityRuntime.PowerPoint = { run: () => { throw new Error('Unsupported host must not render') } }
    assert.equal(await identityExports.firstSlidePreview(), null)
    identityRuntime.Office.context.requirements.isSetSupported = () => true
    let count = 0
    let renders = 0
    identityRuntime.PowerPoint = { run: async callback => callback({ sync: async () => {}, presentation: { slides: {
      getCount: () => ({ value: count }),
      getItemAt: index => { assert.equal(index, 0); return { getImageAsBase64: options => { renders++; assert.equal(options.width, 480); return { value: 'iVBORw0KGgo=' } } } }
    } } }) }
    assert.equal(await identityExports.firstSlidePreview(), null)
    assert.equal(renders, 0)
    count = 3
    assert.equal(await identityExports.firstSlidePreview(), 'data:image/png;base64,iVBORw0KGgo=')
    assert.equal(renders, 1)
    identityRuntime.Office.context.document = { url: 'https://files.example/Test%20Deck.pptx?download=true' }
    assert.equal(identityExports.presentationFilename(), 'Test Deck.pptx')
    identityRuntime.Office.context.document.url = ''
    assert.equal(identityExports.presentationFilename(), 'Untitled presentation')
  } finally { identityRuntime.PowerPoint = saved }
})

test('client uploads exact byte ranges, retries only failed parts and bounds concurrency', async () => {
  const bytes = Uint8Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8])
  const reads = []
  const source = { size: 9, read: async (offset, size) => { reads.push([offset, size]); return bytes.slice(offset, offset + size) } }
  const attempts = new Map()
  let active = 0
  let peak = 0
  let progress = 0
  await transferParts(source, parts, value => { progress = value }, new AbortController().signal, async (url, body, report) => {
    active++
    peak = Math.max(peak, active)
    await Promise.resolve()
    const index = Number(new URL(url).pathname.slice(1))
    assert.deepEqual([...body], [...bytes.slice([0, 3, 7][index], [3, 7, 9][index])])
    attempts.set(index, (attempts.get(index) || 0) + 1)
    active--
    if (index === 1 && attempts.get(index) === 1) throw new TransferError('Temporary', 503)
    report(body.length)
  }, async () => {})
  assert.equal(progress, 9)
  assert.equal(attempts.get(1), 2)
  assert.equal(attempts.get(0), 1)
  assert.equal(reads.length, 3)
  assert.ok(peak <= 2)
})

test('client rejects mismatched plans and does not retry expired signed URLs', async () => {
  assert.throws(() => validateParts(parts, 10))
  assert.throws(() => validateParts([parts[0], parts[0]], 6))
  assert.throws(() => validateParts([{ size: 9, url: 'https://attacker.example/upload' }], 9))
  let attempts = 0
  await assert.rejects(transferParts({ size: 9, read: async (_, size) => new Uint8Array(size) }, parts, () => {}, new AbortController().signal, async () => { attempts++; throw new TransferError('Expired', 403) }), { status: 403 })
  assert.ok(attempts <= 2)
})

test('client honors cancellation before any presentation bytes are read', async () => {
  const controller = new AbortController()
  controller.abort()
  let reads = 0
  await assert.rejects(transferParts({ size: 9, read: async () => { reads++; return new Uint8Array(9) } }, parts, () => {}, controller.signal))
  assert.equal(reads, 0)
})

test('Office reads cross-slice byte ranges and closes its file exactly once', async () => {
  const size = 1024 * 1024
  const bytes = Uint8Array.from({ length: size + 10 }, (_, index) => index % 251)
  let closed = 0
  const file = { size: bytes.length, getSliceAsync: (index, callback) => callback({ status: 'succeeded', value: { index, data: bytes.slice(index * size, (index + 1) * size) } }), closeAsync: callback => { closed++; callback?.({ status: 'succeeded' }) } }
  runtime.Office = { AsyncResultStatus: { Succeeded: 'succeeded' }, FileType: { Compressed: 'compressed' }, context: { document: { getFileAsync: (type, options, callback) => { assert.equal(type, 'compressed'); assert.equal(options.sliceSize, size); callback({ status: 'succeeded', value: file }) } } } }
  const signal = new AbortController().signal
  const source = await openPresentation(signal)
  assert.deepEqual([...await source.read(size - 3, 8, signal)], [...bytes.slice(size - 3, size + 5)])
  await source.close()
  await source.close()
  assert.equal(closed, 1)
})

test('a late Office export result is closed after cancellation', async () => {
  let callback
  let closed = 0
  runtime.Office.context.document.getFileAsync = (_, __, result) => { callback = result }
  const controller = new AbortController()
  const pending = openPresentation(controller.signal)
  controller.abort()
  await assert.rejects(pending)
  callback({ status: 'succeeded', value: { closeAsync: () => { closed++ } } })
  assert.equal(closed, 1)
})

for (const mimeType of [undefined, exportsObject.docxType]) test(`${mimeType ? 'DOCX' : 'PPTX'} signed storage receives binary PUTs without app or provider authorization`, async () => {
  let request
  runtime.XMLHttpRequest = class {
    constructor() { request = this; this.upload = {}; this.headers = {}; this.status = 200 }
    open(method, url) { this.method = method; this.url = url }
    setRequestHeader(name, value) { this.headers[name] = value }
    send(body) { this.body = body; this.onload() }
  }
  const bytes = Uint8Array.from([1, 2, 3])
  await putPart(parts[0].url, bytes, () => {}, new AbortController().signal, mimeType)
  assert.equal(request.method, 'PUT')
  assert.deepEqual([...request.body], [...bytes])
  assert.equal(request.withCredentials, false)
  assert.equal(request.headers['x-amz-acl'], 'private')
  assert.equal(request.headers.Authorization, undefined)
  assert.equal(request.headers['Content-Type'], mimeType || exportsObject.pptxType)
})

test('shared native export retains its legacy alias and rejects unsupported storage content types', async () => {
  assert.equal(exportsObject.openDocument, openPresentation)
  await assert.rejects(putPart(parts[0].url, Uint8Array.from([1]), () => {}, new AbortController().signal, 'application/zip'), /Unsupported document content type/)
})

function uploadPane(options = {}) {
  const document = documentClient()
  const adapter = { ...document.exports.getOfficeHost({ host: options.host || 'Word', platform: 'Mac' }).adapter,
    filename: () => options.filename || 'Untitled document', ensureId: async () => { calls.push({ operation: 'identity' }); return options.changedId ? 'bbbbbbbb-1234-1234-1234-123456789abc' : document.expected } }
  const calls = []
  let completed = false
  const registered = { presentationId: document.expected, documentType: options.historyType || adapter.documentType, state: 'complete', version: 1, accountId: 'account', projectId: 'project', folderId: 'folder', viewUrl: 'https://next.frame.io/file/file', name: adapter.defaultName }
  const states = [], refs = [], effects = []
  let stateIndex, refIndex, effectIndex, changed, pending, tree
  const props = { adapter: options.startup ? null : adapter, unavailableReason: '', available: true, signedIn: true, presentationId: document.expected,
    destination: { accountId: 'account', projectId: 'project', folderId: 'folder' }, onBusyChange: () => {}, onSessionExpired: () => {}, onAssetChange: () => {} }
  const element = (type, props) => ({ type, props })
  const module = { exports: {} }
  const react = {
    useState: initial => {
      const index = stateIndex++
      if (!(index in states)) states[index] = typeof initial === 'function' ? initial() : initial
      return [states[index], next => { states[index] = typeof next === 'function' ? next(states[index]) : next; changed = true }]
    },
    useRef: initial => { const index = refIndex++; return refs[index] ||= { current: initial } },
    useEffect: (callback, deps) => {
      const index = effectIndex++
      if (!effects[index] || deps?.some((value, position) => value !== effects[index].deps[position])) pending.push(() => {
        effects[index]?.cleanup?.()
        effects[index] = { deps, cleanup: callback() }
      })
    }
  }
  const auth = { ApiError: class extends Error { constructor(message, status, relinkAllowed = false) { super(message); this.status = status; this.relinkAllowed = relinkAllowed } }, oauthRequest: async (operation, params) => {
    calls.push({ operation, params })
    if (operation === 'presentation-lookup') {
      if (options.historyStatus && !completed) throw new auth.ApiError('This Frame.io selection is no longer available.', options.historyStatus, options.relinkAllowed === true)
      return { entry: completed || options.historyType ? registered : null }
    }
    if (operation === 'upload-start') {
      Object.assign(registered, { accountId: params.accountId, projectId: params.projectId, folderId: params.folderId })
      return { fileId: 'file', parts: [{ size: 8, url: parts[0].url }] }
    }
    completed = true
    return { complete: true, failed: false, entry: registered }
  } }
  const transfer = { maxFileSize: exportsObject.maxFileSize, delay: async () => {}, openDocument: async () => {
    calls.push({ operation: 'export' })
    return { size: 8, close: async () => { calls.push({ operation: 'close' }) } }
  }, transferParts: async (source, plan, progress, signal, put) => { await put(plan[0].url, new Uint8Array(8), progress, signal); progress(8) },
  putPart: async (url, bytes, progress, signal, mimeType) => { calls.push({ operation: 'put', mimeType }) } }
  vm.runInNewContext(ts.transpileModule(fs.readFileSync(require.resolve('../web-src/UploadPanel.tsx'), 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022, jsx: ts.JsxEmit.ReactJSX } }).outputText, {
    module, exports: module.exports, AbortController, URL, Error, crypto: require('node:crypto').webcrypto,
    require: name => name === 'react' ? react : name === 'react/jsx-runtime' ? { jsx: element, jsxs: element } : name === './auth' ? auth : name === './office-document' ? document.exports : name === './upload-transfer' ? transfer : {}
  })
  function render() {
    let attempts = 0
    do {
      changed = false; pending = []; stateIndex = 0; refIndex = 0; effectIndex = 0
      tree = module.exports.UploadPanel(props)
      pending.forEach(callback => callback())
      assert.ok(++attempts < 10, 'Effects must settle')
    } while (changed)
    return tree
  }
  function elements(value) {
    if (!value || typeof value !== 'object') return []
    if (Array.isArray(value)) return value.flatMap(child => elements(child))
    return [value, ...elements(value.props?.children)]
  }
  return { calls, props, adapter, render, elements: () => elements(tree), find: predicate => elements(tree).find(predicate), settle: async () => { await new Promise(setImmediate); render() } }
}

test('Word pane initializes after host detection and uploads DOCX without losing edited filenames', async () => {
  const pane = uploadPane({ startup: true })
  pane.render()
  const input = () => pane.find(node => node.props?.id === 'upload-name')
  assert.equal(input().props.value, '')
  assert.equal(input().props.disabled, true)
  pane.props.adapter = pane.adapter
  pane.render()
  await pane.settle()
  assert.equal(input().props.value, 'Document.docx')
  input().props.onChange({ target: { value: 'Edited.docx' } })
  pane.render()
  const upload = pane.find(node => node.type === 'button' && node.props.className === 'upload')
  assert.equal(upload.props.disabled, false)
  await upload.props.onClick()
  pane.render()
  assert.deepEqual(pane.calls.map(call => call.operation), ['presentation-lookup', 'identity', 'export', 'upload-start', 'put', 'upload-status', 'close'])
  assert.equal(pane.calls.find(call => call.operation === 'upload-start').params.documentType, 'docx')
  assert.equal(pane.calls.find(call => call.operation === 'upload-start').params.name, 'Edited.docx')
  assert.equal(pane.calls.find(call => call.operation === 'put').mimeType, exportsObject.docxType)
  const refresh = pane.find(node => node.type === 'button' && node.props.children?.includes('Refresh history'))
  refresh.props.onClick()
  pane.render()
  await pane.settle()
  assert.equal(input().props.value, 'Edited.docx')
})

test('Word pane stops before export if its identity changes or linked history is PPTX', async () => {
  for (const options of [{ changedId: true }, { historyType: 'pptx' }]) {
    const pane = uploadPane(options)
    pane.render()
    await pane.settle()
    const upload = pane.find(node => node.type === 'button' && node.props.className === 'upload')
    assert.equal(upload.props.disabled, !!options.historyType)
    if (!upload.props.disabled) { await upload.props.onClick(); pane.render() }
    assert.equal(pane.calls.some(call => call.operation === 'export' || call.operation === 'upload-start'), false)
    const text = pane.elements().flatMap(node => node.props?.children || []).filter(value => typeof value === 'string').join(' ')
    assert.match(text, options.historyType ? /different document type/ : /tracking ID changed/)
  }
})

test('PowerPoint pane still submits PPTX through the shared upload flow', async () => {
  const pane = uploadPane({ host: 'PowerPoint', filename: 'Deck.pptx' })
  pane.render()
  await pane.settle()
  await pane.find(node => node.type === 'button' && node.props.className === 'upload').props.onClick()
  assert.equal(pane.calls.find(call => call.operation === 'upload-start').params.documentType, 'pptx')
  assert.equal(pane.calls.find(call => call.operation === 'put').mimeType, exportsObject.pptxType)
  assert.equal(pane.calls.at(-1).operation, 'close')
})

for (const host of ['Word', 'PowerPoint']) test(`${host} permits a replacement after selecting a destination for a confirmed missing asset`, async () => {
  const pane = uploadPane({ host, filename: host === 'Word' ? 'Report.docx' : 'Deck.pptx', historyStatus: 404, relinkAllowed: true })
  pane.props.destination = null
  pane.render()
  await pane.settle()
  const upload = () => pane.find(node => node.type === 'button' && node.props.className === 'upload')
  assert.equal(upload().props.disabled, true)
  assert.ok(pane.find(node => node.props?.role === 'alert'))
  pane.props.destination = { accountId: 'new-account', projectId: 'new-project', folderId: 'new-folder' }
  pane.render()
  assert.equal(upload().props.disabled, false)
  assert.ok(pane.find(node => node.props?.role === 'alert'))
  await upload().props.onClick()
  pane.render()
  const start = pane.calls.find(call => call.operation === 'upload-start').params
  assert.equal(start.presentationId, pane.props.presentationId)
  assert.equal(start.accountId, 'new-account')
  assert.equal(start.projectId, 'new-project')
  assert.equal(start.folderId, 'new-folder')
  assert.equal(pane.find(node => node.props?.role === 'alert'), undefined)
})

test('a selected destination does not bypass unknown, inaccessible or moved document history', async () => {
  for (const [historyStatus, relinkAllowed] of [[404, false], [403, false], [409, false], [502, false], [403, true]]) {
    const pane = uploadPane({ host: 'PowerPoint', filename: 'Deck.pptx', historyStatus, relinkAllowed })
    pane.render()
    await pane.settle()
    const upload = pane.find(node => node.type === 'button' && node.props.className === 'upload')
    assert.equal(upload.props.disabled, true)
    await upload.props.onClick()
    assert.equal(pane.calls.some(call => call.operation === 'export' || call.operation === 'upload-start'), false)
  }
})
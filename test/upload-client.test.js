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

test('signed storage receives binary PUTs without app or provider authorization', async () => {
  let request
  runtime.XMLHttpRequest = class {
    constructor() { request = this; this.upload = {}; this.headers = {}; this.status = 200 }
    open(method, url) { this.method = method; this.url = url }
    setRequestHeader(name, value) { this.headers[name] = value }
    send(body) { this.body = body; this.onload() }
  }
  const bytes = Uint8Array.from([1, 2, 3])
  await putPart(parts[0].url, bytes, () => {}, new AbortController().signal)
  assert.equal(request.method, 'PUT')
  assert.deepEqual([...request.body], [...bytes])
  assert.equal(request.withCredentials, false)
  assert.equal(request.headers['x-amz-acl'], 'private')
  assert.equal(request.headers.Authorization, undefined)
  assert.equal(request.headers['Content-Type'], exportsObject.pptxType)
})
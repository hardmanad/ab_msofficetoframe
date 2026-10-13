'use strict'

const { test } = require('node:test')
const assert = require('node:assert/strict')
const { uploadOperation, maxFileSize } = require('../actions/oauth/uploads')
const { createRegistry } = require('../actions/oauth/registry')
const { localState } = require('../scripts/local-state.cjs')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')

class RequestError extends Error { constructor(status, message) { super(message); this.status = status } }
function fixture(overrides = {}) {
  const records = new Map()
  const registryFiles = new Map()
  const registry = createRegistry({ list: async name => registryFiles.has(name) ? [{ name }] : [], read: async name => Buffer.from(registryFiles.get(name)), write: async (name, value) => registryFiles.set(name, value) })
  const calls = []
  const params = { operation: 'upload-start', presentationId: 'aaaaaaaa-1234-1234-1234-123456789abc', uploadId: '12345678-1234-1234-1234-123456789abc', accountId: 'account', projectId: 'project', folderId: 'root', name: 'Deck.pptx', fileSize: 8, ...overrides }
  const options = { params, identity: 'owner', registry, read: async key => records.get(key), write: async (key, value) => records.set(key, value), RequestError,
    request: async (path, paginated, init) => {
      calls.push({ path, init })
      if (path.endsWith('/projects/project')) return { data: { root_folder_id: 'root' } }
      if (path.endsWith('/folders/other')) return { data: { project_id: 'another-project' } }
      if (path.endsWith('/status')) return { data: { upload_complete: true, upload_failed: false } }
      if (path.endsWith('/files/file')) return { data: { id: 'file', name: 'Deck.pptx', project_id: 'project', parent_id: 'root', view_url: 'https://next.frame.io/file/file' } }
      return { data: { id: 'file', upload_urls: [{ size: 8, url: 'https://frameio-test.s3-accelerate.amazonaws.com/part?signature=test' }] } }
    } }
  return { options, records, calls }
}

test('upload creation is fixed-path, replayable and status is scoped to the owning session', async () => {
  const { options, calls } = fixture()
  const first = await uploadOperation(options)
  assert.equal(first.parts[0].size, 8)
  assert.deepEqual(await uploadOperation(options), first)
  assert.equal(calls.filter(call => call.init?.method === 'POST').length, 1)
  assert.deepEqual(JSON.parse(calls[1].init.body), { data: { name: 'Deck.pptx', file_size: 8 } })
  options.params = { operation: 'upload-status', uploadId: options.params.uploadId }
  assert.equal((await uploadOperation(options)).complete, true)
  await assert.rejects(uploadOperation({ ...options, identity: 'another-owner' }), { status: 404 })
})

test('upload rejects unsafe names, oversize files and cross-project destinations before creating a file', async () => {
  for (const overrides of [{ name: '../Deck.pptx' }, { fileSize: maxFileSize + 1 }, { fileSize: 0 }, { folderId: 'other' }]) {
    const { options, calls } = fixture(overrides)
    await assert.rejects(uploadOperation(options))
    assert.equal(calls.filter(call => call.init?.method === 'POST').length, 0)
  }
})

test('provider failures remain in durable history without becoming the current asset', async () => {
  const { options } = fixture()
  const presentationId = options.params.presentationId
  const uploadId = options.params.uploadId
  await uploadOperation(options)
  const original = options.request
  options.request = async (path, ...rest) => path.endsWith('/status') ? { data: { upload_complete: false, upload_failed: true } } : original(path, ...rest)
  options.params = { operation: 'upload-status', uploadId }
  assert.equal((await uploadOperation(options)).failed, true)
  assert.equal((await options.registry.event(presentationId, uploadId)).state, 'failed')
  assert.equal(await options.registry.latest(presentationId), undefined)
})

test('upload does not repeat an uncertain create request or return arbitrary upload URLs', async () => {
  const { options, calls } = fixture()
  const original = options.request
  options.request = async (...args) => {
    const result = await original(...args)
    if (args[2]?.method === 'POST') result.data.upload_urls[0].url = 'https://attacker.example/upload'
    return result
  }
  await assert.rejects(uploadOperation(options), { status: 502 })
  await assert.rejects(uploadOperation(options), { status: 409 })
  assert.equal(calls.filter(call => call.init?.method === 'POST').length, 1)
})

test('real local storage supports session-bound upload keys without allowing path traversal', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'frameio-upload-state-'))
  try {
    const store = localState(directory)
    const { options } = fixture()
    options.identity = 'ab'.repeat(32)
    options.read = async key => {
      const record = await store.get(key)
      return record ? JSON.parse(record.value) : undefined
    }
    options.write = async (key, value, ttl) => store.put(key, JSON.stringify(value), { ttl })
    assert.equal((await uploadOperation(options)).fileId, 'file')
    options.params = { operation: 'upload-status', uploadId: options.params.uploadId }
    assert.equal((await uploadOperation(options)).complete, true)
    const key = 'upload-' + options.identity + '-' + options.params.uploadId
    assert.equal((await fs.stat(path.join(directory, key + '.json'))).mode & 0o777, 0o600)
    await store.delete(key)
    assert.equal(await store.get(key), undefined)
    for (const invalid of ['../session-' + options.identity, key + '/../escape', 'upload-owner-' + options.params.uploadId]) {
      await assert.rejects(store.get(invalid), /Invalid storage key/)
      await assert.rejects(store.put(invalid, 'fixture', { ttl: 60 }), /Invalid storage key/)
      await assert.rejects(store.delete(invalid), /Invalid storage key/)
    }
    for (const otherKey of ['session-' + options.identity, 'login-' + 'a'.repeat(43)]) {
      await store.put(otherKey, 'encrypted-fixture', { ttl: 60 })
      assert.equal((await store.get(otherKey)).value, 'encrypted-fixture')
    }
  } finally { await fs.rm(directory, { recursive: true, force: true }) }
})
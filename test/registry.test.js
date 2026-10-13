'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { localRegistry } = require('../scripts/local-registry.cjs')
const { finalizeUpload, verifyAsset } = require('../actions/oauth/registry')
class RequestError extends Error { constructor(status, message) { super(message); this.status = status } }
const presentationId = '12345678-1234-1234-1234-123456789abc'
const uploadIds = ['aaaaaaaa-1234-1234-1234-123456789abc', 'bbbbbbbb-1234-1234-1234-123456789abc', 'cccccccc-1234-1234-1234-123456789abc']

test('durable registry keeps first upload and later versions across adapter restarts', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'frameio-registry-'))
  try {
    let registry = localRegistry(directory)
    const files = new Map([['first', { id: 'first', project_id: 'project', parent_id: 'folder', view_url: 'https://next.frame.io/file/first' }]])
    let stacks = 0
    let moves = 0
    const request = async (url, paginated, init) => {
      if (init?.method === 'POST') {
        stacks++
        const ids = JSON.parse(init.body).data.file_ids
        assert.deepEqual(ids, ['first', 'second'])
        ids.forEach(id => { files.get(id).parent_id = 'stack' })
        return { data: { id: 'stack', project_id: 'project', parent_id: 'folder', view_url: 'https://next.frame.io/stack/stack' } }
      }
      if (init?.method === 'PATCH') { moves++; files.get('third').parent_id = JSON.parse(init.body).data.parent_id; return { data: files.get('third') } }
      if (url.endsWith('/version_stacks/stack')) return { data: { id: 'stack', project_id: 'project', parent_id: 'folder', view_url: 'https://next.frame.io/stack/stack' } }
      return { data: files.get(url.split('/').pop()) }
    }
    const record = { presentationId, uploadId: uploadIds[0], accountId: 'account', projectId: 'project', folderId: 'folder', name: 'Deck.pptx', fileSize: 8, fileId: 'first', createdAt: new Date().toISOString() }
    assert.equal(await registry.latest(presentationId), undefined)
    const first = await finalizeUpload(record, registry, request, RequestError)
    assert.equal(first.version, 1)
    registry = localRegistry(directory)
    assert.equal((await registry.latest(presentationId)).fileId, 'first')
    files.set('second', { id: 'second', project_id: 'project', parent_id: 'folder', view_url: 'https://next.frame.io/file/second' })
    const second = await finalizeUpload({ ...record, fileId: 'second', uploadId: uploadIds[1] }, registry, request, RequestError)
    assert.equal(second.stackId, 'stack')
    assert.equal(second.version, 2)
    await finalizeUpload({ ...record, fileId: 'second', uploadId: uploadIds[1] }, registry, request, RequestError)
    assert.equal(stacks, 1)
    files.set('third', { id: 'third', project_id: 'project', parent_id: 'folder', view_url: 'https://next.frame.io/file/third' })
    const third = await finalizeUpload({ ...record, fileId: 'third', uploadId: uploadIds[2] }, registry, request, RequestError)
    assert.equal(third.version, 3)
    assert.equal(moves, 1)
    assert.equal((await localRegistry(directory).latest(presentationId)).fileId, 'third')
    await finalizeUpload(record, registry, request, RequestError)
    assert.equal((await registry.latest(presentationId)).fileId, 'third')
    assert.equal((await registry.event(presentationId, uploadIds[0])).fileId, 'first')
    assert.equal((await registry.event(presentationId, uploadIds[1])).fileId, 'second')
    assert.equal((await fs.stat(path.join(directory, 'presentation-uploads', presentationId, 'current.json'))).mode & 0o777, 0o600)
  } finally { await fs.rm(directory, { recursive: true, force: true }) }
})

test('registry lookup verifies current Frame.io access and refuses moved assets or arbitrary links', async () => {
  const entry = { accountId: 'account', projectId: 'project', folderId: 'folder', fileId: 'file' }
  await assert.rejects(verifyAsset(entry, async () => { throw new RequestError(403, 'Denied') }, RequestError), { status: 403 })
  await assert.rejects(verifyAsset(entry, async () => ({ data: { id: 'file', project_id: 'other', parent_id: 'folder' } }), RequestError), { status: 409 })
  await assert.rejects(verifyAsset(entry, async () => ({ data: { id: 'file', project_id: 'project', parent_id: 'folder', view_url: 'https://attacker.example/file' } }), RequestError), { status: 502 })
})

test('a lost version-stack response is recovered without creating a duplicate stack', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'frameio-stack-recovery-'))
  try {
    const registry = localRegistry(directory)
    const old = { id: 'old', project_id: 'project', parent_id: 'folder', view_url: 'https://next.frame.io/file/old' }
    const fresh = { id: 'fresh', project_id: 'project', parent_id: 'folder', view_url: 'https://next.frame.io/file/fresh' }
    const first = { presentationId, uploadId: uploadIds[0], accountId: 'account', projectId: 'project', folderId: 'folder', fileId: 'old', name: 'Deck.pptx', fileSize: 8, state: 'complete', version: 1, viewUrl: old.view_url, completedAt: new Date().toISOString() }
    await registry.record(presentationId, first)
    let creates = 0
    const request = async (url, paginated, init) => {
      if (init?.method === 'POST') { creates++; old.parent_id = fresh.parent_id = 'stack'; throw new Error('Response lost') }
      if (url.endsWith('/files/old')) return { data: old }
      if (url.endsWith('/files/fresh')) return { data: fresh }
      return { data: { id: 'stack', project_id: 'project', parent_id: 'folder', view_url: 'https://next.frame.io/stack/stack' } }
    }
    const record = { ...first, fileId: 'fresh', uploadId: uploadIds[1], createdAt: new Date().toISOString() }
    await assert.rejects(finalizeUpload(record, registry, request, RequestError), /Response lost/)
    assert.equal((await registry.latest(presentationId)).version, 1)
    const recovered = await finalizeUpload(record, registry, request, RequestError)
    assert.equal(recovered.stackId, 'stack')
    assert.equal(recovered.version, 2)
    assert.equal(creates, 1)
  } finally { await fs.rm(directory, { recursive: true, force: true }) }
})
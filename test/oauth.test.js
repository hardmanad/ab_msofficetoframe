'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')
const { createHandler } = require('../actions/oauth')
const { createRegistry } = require('../actions/oauth/registry')

function makeFixture(frameFetch) {
  const values = new Map()
  const registryFiles = new Map()
  const registry = createRegistry({ list: async name => registryFiles.has(name) ? [{ name }] : [], read: async name => Buffer.from(registryFiles.get(name)), write: async (name, value) => registryFiles.set(name, value) })
  let time = Date.now()
  let requests = 0
  let nonce
  const config = {
    FRAMEIO_CLIENT_ID: 'test-client', FRAMEIO_CLIENT_SECRET: 'test-client-secret',
    FRAMEIO_REDIRECT_URI: 'https://www.localappbuilder.com:3000/oauth-callback.html',
    FRAMEIO_SCOPES: 'openid email profile offline_access additional_info.roles', FRAMEIO_TOKEN_KEY: 'ab'.repeat(32)
  }
  const store = { get: async key => values.has(key) ? { value: values.get(key) } : undefined, put: async (key, value) => values.set(key, value), delete: async key => values.delete(key) }
  const handler = createHandler({ getStore: async () => store, getRegistry: async () => registry, now: () => time, fetchImpl: async (_url, options) => {
    if (new URL(_url).origin === 'https://api.frame.io' && frameFetch) return frameFetch(new URL(_url), options)
    requests++
    const body = new URLSearchParams(options.body)
    assert.equal(body.get('client_secret'), config.FRAMEIO_CLIENT_SECRET)
    if (body.get('grant_type') === 'authorization_code') assert.ok(body.get('code_verifier'))
    const encode = value => Buffer.from(JSON.stringify(value)).toString('base64url')
    const idToken = encode({ alg: 'RS256' }) + '.' + encode({ iss: 'https://ims-na1.adobelogin.com', aud: config.FRAMEIO_CLIENT_ID, sub: 'test-user', nonce, iat: Math.floor(Date.now() / 1000), exp: Math.floor(Date.now() / 1000) + 3600 }) + '.dGVzdA'
    return new Response(JSON.stringify({ access_token: 'private-access-token', refresh_token: 'private-refresh-token', expires_in: 3600, token_type: 'bearer', ...(body.get('grant_type') === 'authorization_code' ? { id_token: idToken } : {}) }), { status: 200, headers: { 'Content-Type': 'application/json' } })
  } })
  const call = (operation, extra = {}) => handler({ ...config, __ow_method: 'post', __ow_headers: { origin: 'https://www.localappbuilder.com:3000', authorization: 'Bearer ' + 'cd'.repeat(32) }, operation, ...extra })
  return { call, values, registryFiles, setNonce: value => { nonce = value }, advance: milliseconds => { time += milliseconds }, requests: () => requests }
}

for (const documentType of [undefined, 'docx']) test(`${documentType || 'legacy PPTX'} upload routes require login, use the user token and store signed URLs encrypted`, async () => {
  let creates = 0
  const fixture = makeFixture(async (url, options) => {
    assert.equal(options.headers.Authorization, 'Bearer private-access-token')
    let data
    if (url.pathname.endsWith('/projects/project')) data = { root_folder_id: 'root' }
    else if (url.pathname.endsWith('/status')) data = { upload_complete: true, upload_failed: false }
    else if (url.pathname.endsWith('/files/file')) data = { id: 'file', name: 'Deck.pptx', project_id: 'project', parent_id: 'root', view_url: 'https://next.frame.io/file/file' }
    else {
      assert.equal(url.pathname, '/v4/accounts/account/folders/root/files/local_upload')
      assert.equal(options.method, 'POST')
      creates++
      data = { id: 'file', upload_urls: [{ size: 8, url: 'https://frameio-test.s3-accelerate.amazonaws.com/part?signature=private-signature' }] }
    }
    return new Response(JSON.stringify({ data }), { headers: { 'Content-Type': 'application/json' } })
  })
  const params = { presentationId: 'aaaaaaaa-1234-1234-1234-123456789abc', ...(documentType ? { documentType } : {}), uploadId: '12345678-1234-1234-1234-123456789abc', accountId: 'account', projectId: 'project', folderId: 'root', name: documentType === 'docx' ? 'Report.docx' : 'Deck.pptx', fileSize: 8 }
  assert.equal((await fixture.call('upload-start', params)).statusCode, 401)
  const start = await fixture.call('start')
  const url = new URL(start.body.authorizationUrl)
  fixture.setNonce(url.searchParams.get('nonce'))
  await fixture.call('callback', { state: url.searchParams.get('state'), code: 'code' })
  const result = await fixture.call('upload-start', params)
  assert.equal(result.statusCode, 200)
  assert.equal(result.body.parts[0].size, 8)
  assert.equal(JSON.stringify(result.body).includes('private-access-token'), false)
  assert.equal([...fixture.values.values()].join('').includes('private-signature'), false)
  assert.equal((await fixture.call('upload-start', params)).statusCode, 200)
  assert.equal(creates, 1)
  assert.equal((await fixture.call('upload-status', { uploadId: params.uploadId })).body.complete, true)
  const found = await fixture.call('presentation-lookup', { presentationId: params.presentationId })
  assert.equal(found.body.entry.fileId, 'file')
  assert.equal(found.body.entry.version, 1)
  assert.equal(found.body.entry.documentType, documentType || 'pptx')
  assert.equal([...fixture.registryFiles.values()].join('').includes('private-signature'), false)
  assert.equal([...fixture.registryFiles.values()].join('').includes('private-access-token'), false)
  fixture.advance(25 * 60 * 60 * 1000)
  assert.equal((await fixture.call('presentation-lookup', { presentationId: params.presentationId })).statusCode, 401)
  const freshHeaders = { origin: 'https://www.localappbuilder.com:3000', authorization: 'Bearer ' + 'ef'.repeat(32) }
  const freshStart = await fixture.call('start', { __ow_headers: freshHeaders })
  const freshUrl = new URL(freshStart.body.authorizationUrl)
  fixture.setNonce(freshUrl.searchParams.get('nonce'))
  await fixture.call('callback', { state: freshUrl.searchParams.get('state'), code: 'fresh-code' })
  const reopened = await fixture.call('presentation-lookup', { presentationId: params.presentationId, __ow_headers: freshHeaders })
  assert.equal(reopened.body.entry.fileId, 'file')
})

test('history lookup keeps the missing-asset warning and authorizes recovery only with original project access', async () => {
  let projectStatus = 200
  const fixture = makeFixture(async url => {
    if (url.pathname.endsWith('/projects/project')) return Response.json({ data: { id: 'project', root_folder_id: 'root' } }, { status: projectStatus })
    return Response.json({ error: 'Missing' }, { status: 404 })
  })
  const presentationId = 'aaaaaaaa-1234-1234-1234-123456789abc'
  const entry = { presentationId, uploadId: '12345678-1234-1234-1234-123456789abc', accountId: 'account', projectId: 'project', folderId: 'root', fileId: 'deleted', state: 'complete', version: 1 }
  const current = `presentation-uploads/${presentationId}/current.json`
  fixture.registryFiles.set(current, JSON.stringify(entry))
  const start = await fixture.call('start')
  const authorization = new URL(start.body.authorizationUrl)
  fixture.setNonce(authorization.searchParams.get('nonce'))
  await fixture.call('callback', { state: authorization.searchParams.get('state'), code: 'code' })
  const missing = await fixture.call('presentation-lookup', { presentationId })
  assert.equal(missing.statusCode, 404)
  assert.equal(missing.body.error, 'This Frame.io selection is no longer available.')
  assert.equal(missing.body.relinkAllowed, true)
  assert.equal(JSON.parse(fixture.registryFiles.get(current)).fileId, 'deleted')
  projectStatus = 403
  const denied = await fixture.call('presentation-lookup', { presentationId })
  assert.equal(denied.statusCode, 403)
  assert.equal(denied.body.relinkAllowed, undefined)
  assert.equal(JSON.parse(fixture.registryFiles.get(current)).fileId, 'deleted')
})

test('OAuth uses PKCE, encrypted storage and session-scoped output without returning tokens', async () => {
  const fixture = makeFixture()
  const start = await fixture.call('start')
  assert.equal(start.statusCode, 200)
  const url = new URL(start.body.authorizationUrl)
  assert.equal(url.origin, 'https://ims-na1.adobelogin.com')
  assert.equal(url.searchParams.get('code_challenge_method'), 'S256')
  const state = url.searchParams.get('state')
  fixture.setNonce(url.searchParams.get('nonce'))
  const result = await fixture.call('callback', { state, code: 'provider-code' })
  assert.deepEqual(result.body, { authenticated: true })
  assert.equal([...fixture.values.values()].some(value => value.includes('private-access-token') || value.includes('private-refresh-token')), false)
  assert.deepEqual((await fixture.call('session')).body, { authenticated: true })
  fixture.advance(3550000)
  assert.deepEqual((await fixture.call('session')).body, { authenticated: true })
  assert.equal(fixture.requests(), 2)
  assert.equal((await fixture.call('callback', { state, code: 'provider-code' })).statusCode, 400)
  assert.deepEqual((await fixture.call('logout')).body, { authenticated: false, revoked: true })
  assert.deepEqual((await fixture.call('session')).body, { authenticated: false })
})

test('OAuth rejects wrong origins, missing session keys, invalid state and methods', async () => {
  const fixture = makeFixture()
  assert.equal((await fixture.call('start', { __ow_headers: { origin: 'https://attacker.invalid' } })).statusCode, 403)
  assert.equal((await fixture.call('start', { __ow_headers: { origin: 'https://www.localappbuilder.com:3000' } })).statusCode, 401)
  assert.equal((await fixture.call('start', { __ow_method: 'get' })).statusCode, 405)
  assert.equal((await fixture.call('proxy')).statusCode, 400)
  assert.equal((await fixture.call('callback', { state: 'unknown' })).statusCode, 400)
  assert.equal(fixture.requests(), 0)
})

test('expired transactions and provider denial do not create sessions', async () => {
  const fixture = makeFixture()
  const start = await fixture.call('start')
  const state = new URL(start.body.authorizationUrl).searchParams.get('state')
  fixture.advance(601000)
  assert.equal((await fixture.call('callback', { state, code: 'expired-code' })).statusCode, 400)
  const next = await fixture.call('start')
  const nextState = new URL(next.body.authorizationUrl).searchParams.get('state')
  assert.equal((await fixture.call('callback', { state: nextState, error: 'access_denied' })).statusCode, 502)
  assert.equal((await fixture.call('session')).body.authenticated, false)
  assert.equal(fixture.requests(), 0)
})

test('a different app session cannot read another user login', async () => {
  const fixture = makeFixture()
  const start = await fixture.call('start')
  const url = new URL(start.body.authorizationUrl)
  fixture.setNonce(url.searchParams.get('nonce'))
  await fixture.call('callback', { state: url.searchParams.get('state'), code: 'provider-code' })
  const result = await fixture.call('session', { __ow_headers: { origin: 'https://www.localappbuilder.com:3000', authorization: 'Bearer ' + 'ef'.repeat(32) } })
  assert.deepEqual(result.body, { authenticated: false })
})

test('an incorrect ID-token nonce cannot create a session', async () => {
  const fixture = makeFixture()
  const start = await fixture.call('start')
  const url = new URL(start.body.authorizationUrl)
  fixture.setNonce('wrong-nonce')
  const result = await fixture.call('callback', { state: url.searchParams.get('state'), code: 'provider-code' })
  assert.equal(result.statusCode, 502)
  assert.equal(JSON.stringify(result).includes('test-client-secret'), false)
  assert.deepEqual((await fixture.call('session')).body, { authenticated: false })
})

async function authenticate(fixture) {
  const result = await fixture.call('start')
  const url = new URL(result.body.authorizationUrl)
  fixture.setNonce(url.searchParams.get('nonce'))
  await fixture.call('callback', { state: url.searchParams.get('state'), code: 'provider-code' })
}

function json(body, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } })
}

test('browsing uses only the signed-in user token and fixed account-scoped paths', async () => {
  const requests = []
  const fixture = makeFixture(async (url, options) => {
    requests.push(url)
    assert.equal(options.headers.Authorization, 'Bearer private-access-token')
    assert.equal(options.method, 'GET')
    assert.equal(options.redirect, 'error')
    return json({ data: [{ id: 'choice-1', ...(url.pathname === '/v4/accounts' ? { display_name: 'Choice' } : { name: 'Choice' }), private_detail: 'not-returned' }], links: { next: url.pathname + '?after=next-page' } })
  })
  assert.equal((await fixture.call('accounts')).statusCode, 401)
  assert.equal(requests.length, 0)
  await authenticate(fixture)
  assert.deepEqual((await fixture.call('accounts')).body, { items: [{ id: 'choice-1', name: 'Choice' }], nextCursor: 'next-page' })
  assert.equal(requests[0].searchParams.has('sort'), false)
  await fixture.call('projects', { accountId: 'account-1', cursor: 'opaque-cursor' })
  assert.equal(requests[1].pathname, '/v4/accounts/account-1/projects')
  assert.equal(requests[1].searchParams.get('after'), 'opaque-cursor')
  assert.equal(requests[1].searchParams.get('sort'), 'name_asc')
  assert.equal((await fixture.call('projects', { accountId: '../other' })).statusCode, 400)
  fixture.advance(3550000)
  assert.equal((await fixture.call('accounts')).statusCode, 200)
  assert.equal(fixture.requests(), 2)
})

test('folder browsing verifies the project root and rejects cross-project folders', async () => {
  const fixture = makeFixture(async url => {
    if (url.pathname.endsWith('/projects/project-1')) return json({ data: { id: 'project-1', name: 'Project', root_folder_id: 'root-1' } })
    if (url.pathname.endsWith('/folders/wrong-folder')) return json({ data: { id: 'wrong-folder', name: 'Wrong', project_id: 'project-2', parent_id: 'root-2' } })
    if (url.pathname.endsWith('/folders/child-1')) return json({ data: { id: 'child-1', name: 'Child', project_id: 'project-1', parent_id: 'root-1' } })
    return json({ data: [{ id: 'child-1', name: 'Child', project_id: 'project-1' }], links: { next: null } })
  })
  await authenticate(fixture)
  const root = await fixture.call('folders', { accountId: 'account-1', projectId: 'project-1' })
  assert.deepEqual(root.body.folder, { id: 'root-1', name: 'Project root', parentId: null })
  const child = await fixture.call('folders', { accountId: 'account-1', projectId: 'project-1', folderId: 'child-1' })
  assert.deepEqual(child.body.folder, { id: 'child-1', name: 'Child', parentId: 'root-1' })
  assert.equal((await fixture.call('folders', { accountId: 'account-1', projectId: 'project-1', folderId: 'wrong-folder' })).statusCode, 403)
})

test('browsing never follows arbitrary pagination URLs and sanitizes permission errors', async () => {
  let status = 200
  const fixture = makeFixture(async () => json(status === 200
    ? { data: [], links: { next: 'https://attacker.invalid/steal?after=cursor' } }
    : { errors: [{ detail: 'private-provider-error' }] }, status))
  await authenticate(fixture)
  assert.equal((await fixture.call('accounts')).statusCode, 502)
  for (const code of [401, 403, 404, 429]) {
    status = code
    const result = await fixture.call('accounts')
    assert.equal(result.statusCode, code)
    assert.equal(JSON.stringify(result).includes('private-provider-error'), false)
  }
})
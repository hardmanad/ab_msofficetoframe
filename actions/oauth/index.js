'use strict'

const crypto = require('node:crypto')
const stateLib = require('@adobe/aio-lib-state')
const { uploadOperation } = require('./uploads')
const { createRegistry, verifyAsset, uuid } = require('./registry')

const provider = {
  issuer: 'https://ims-na1.adobelogin.com',
  authorization_endpoint: 'https://ims-na1.adobelogin.com/ims/authorize/v2',
  token_endpoint: 'https://ims-na1.adobelogin.com/ims/token/v3',
  revocation_endpoint: 'https://ims-na1.adobelogin.com/ims/revoke',
  code_challenge_methods_supported: ['S256']
}
const transactionTTL = 600
const sessionTTL = 86400

class RequestError extends Error {
  constructor(status, message) { super(message); this.status = status }
}

function configuration(params) {
  const { FRAMEIO_CLIENT_ID: clientId, FRAMEIO_CLIENT_SECRET: clientSecret, FRAMEIO_REDIRECT_URI: redirectUri, FRAMEIO_SCOPES: scopes, FRAMEIO_TOKEN_KEY: key } = params
  if (typeof clientId !== 'string' || !clientId || typeof clientSecret !== 'string' || !clientSecret || typeof scopes !== 'string' || !scopes || !/^[a-f0-9]{64}$/.test(key || '')) {
    throw new RequestError(503, 'Sign-in is not configured.')
  }
  let redirect
  try { redirect = new URL(redirectUri) } catch { throw new RequestError(503, 'Sign-in redirect is not configured.') }
  if (redirect.protocol !== 'https:' || redirect.pathname.split('/').pop() !== 'oauth-callback.html' || redirect.username || redirect.password || redirect.search || redirect.hash) {
    throw new RequestError(503, 'Sign-in redirect is not configured.')
  }
  return { clientId, clientSecret, redirectUri, scopes, key: Buffer.from(key, 'hex'), origin: redirect.origin }
}

function seal(value, key, identity) {
  const iv = crypto.randomBytes(12)
  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv)
  cipher.setAAD(Buffer.from(identity))
  const encrypted = Buffer.concat([cipher.update(JSON.stringify(value), 'utf8'), cipher.final()])
  return Buffer.concat([iv, cipher.getAuthTag(), encrypted]).toString('base64url')
}

function open(value, key, identity) {
  const buffer = Buffer.from(value, 'base64url')
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, buffer.subarray(0, 12))
  decipher.setAAD(Buffer.from(identity))
  decipher.setAuthTag(buffer.subarray(12, 28))
  return JSON.parse(Buffer.concat([decipher.update(buffer.subarray(28)), decipher.final()]).toString('utf8'))
}

function sessionIdentity(headers) {
  const match = typeof headers.authorization === 'string' && headers.authorization.match(/^Bearer ([a-f0-9]{64})$/)
  if (!match) throw new RequestError(401, 'A valid app session is required.')
  return crypto.createHash('sha256').update(match[1]).digest('hex')
}

function createHandler({ getStore = () => stateLib.init(), getRegistry = async () => createRegistry(await require('@adobe/aio-lib-files').init()), fetchImpl = fetch, now = Date.now } = {}) {
  return async function handle(params = {}) {
    let config
    const response = (statusCode, body) => ({ statusCode, headers: {
      'Content-Type': 'application/json', 'Cache-Control': 'no-store', 'Vary': 'Origin',
      ...(config ? { 'Access-Control-Allow-Origin': config.origin, 'Access-Control-Allow-Methods': 'POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type, Authorization' } : {})
    }, body })
    try {
      if (!params || typeof params !== 'object' || Array.isArray(params)) throw new RequestError(400, 'Invalid request.')
      config = configuration(params)
      const headers = params.__ow_headers || {}
      if (headers.origin !== config.origin) throw new RequestError(403, 'Request origin is not allowed.')
      const method = typeof params.__ow_method === 'string' ? params.__ow_method.toLowerCase() : ''
      if (method === 'options') return response(204, '')
      if (method !== 'post') throw new RequestError(405, 'Only POST is supported.')
      const browsing = ['accounts', 'projects', 'folders', 'presentation-lookup', 'upload-start', 'upload-status'].includes(params.operation)
      if (!['start', 'callback', 'session', 'logout', 'accounts', 'projects', 'folders', 'presentation-lookup', 'upload-start', 'upload-status'].includes(params.operation)) throw new RequestError(400, 'Unknown operation.')
      const oauth = await import('oauth4webapi')
      const client = { client_id: config.clientId }
      const auth = oauth.ClientSecretPost(config.clientSecret)
      const options = { [oauth.customFetch]: fetchImpl, signal: () => AbortSignal.timeout(15000) }
      const store = await getStore()

      async function read(identity) {
        const result = await store.get(identity)
        if (!result?.value) return undefined
        const value = open(result.value, config.key, identity)
        return value.expiresAt > now() ? value : undefined
      }
      async function write(identity, value, ttl) {
        await store.put(identity, seal(value, config.key, identity), { ttl })
      }
      async function saveTokens(identity, tokens, oldRefresh) {
        if (!Number.isFinite(tokens.expires_in) || tokens.expires_in <= 0 || typeof tokens.access_token !== 'string') {
          throw new RequestError(502, 'The identity provider returned an invalid token response.')
        }
        await write('session-' + identity, {
          accessToken: tokens.access_token, refreshToken: tokens.refresh_token || oldRefresh,
          accessExpiresAt: now() + tokens.expires_in * 1000, expiresAt: now() + sessionTTL * 1000
        }, sessionTTL)
      }

      if (params.operation === 'callback') {
        if (typeof params.state !== 'string' || !/^[A-Za-z0-9_-]{43}$/.test(params.state)) throw new RequestError(400, 'Invalid login state.')
        const recordKey = 'login-' + params.state
        const transaction = await read(recordKey)
        if (!transaction) throw new RequestError(400, 'Login expired or was already used. Start sign-in again.')
        const callback = new URL(config.redirectUri)
        callback.searchParams.set('state', params.state)
        if (typeof params.error === 'string') callback.searchParams.set('error', params.error.slice(0, 100))
        else if (typeof params.code === 'string' && params.code.length > 0 && params.code.length <= 4096) callback.searchParams.set('code', params.code)
        else throw new RequestError(400, 'Authorization code is missing.')
        await store.delete(recordKey)
        const authorization = oauth.validateAuthResponse(provider, client, callback, params.state)
        const tokenResponse = await oauth.authorizationCodeGrantRequest(provider, client, auth, authorization, config.redirectUri, transaction.verifier, options)
        const tokens = await oauth.processAuthorizationCodeResponse(provider, client, tokenResponse, { expectedNonce: transaction.nonce, requireIdToken: true })
        await saveTokens(transaction.session, tokens)
        return response(200, { authenticated: true })
      }

      const identity = sessionIdentity(headers)
      const sessionKey = 'session-' + identity
      if (params.operation === 'start') {
        const state = oauth.generateRandomState()
        const verifier = oauth.generateRandomCodeVerifier()
        const nonce = oauth.generateRandomNonce()
        await write('login-' + state, { session: identity, verifier, nonce, expiresAt: now() + transactionTTL * 1000 }, transactionTTL)
        const url = new URL(provider.authorization_endpoint)
        for (const [key, value] of Object.entries({ client_id: config.clientId, redirect_uri: config.redirectUri, response_type: 'code', scope: config.scopes, state, nonce, code_challenge_method: 'S256', code_challenge: await oauth.calculatePKCECodeChallenge(verifier) })) {
          url.searchParams.set(key, value)
        }
        return response(200, { authorizationUrl: url.href })
      }

      const session = await read(sessionKey)
      if (params.operation === 'logout') {
        await store.delete(sessionKey)
        let revoked = !session
        if (session) {
          try {
            const result = await oauth.revocationRequest(provider, client, auth, session.refreshToken || session.accessToken, options)
            revoked = result.ok
          } catch { revoked = false }
        }
        return response(200, { authenticated: false, revoked })
      }
      if (!session) {
        if (browsing) throw new RequestError(401, 'Sign in to browse Frame.io.')
        return response(200, { authenticated: false })
      }
      if (session.accessExpiresAt <= now() + 60000) {
        if (!session.refreshToken) {
          await store.delete(sessionKey)
          if (browsing) throw new RequestError(401, 'Your session expired. Sign in again.')
          return response(200, { authenticated: false })
        }
        try {
          const tokenResponse = await oauth.refreshTokenGrantRequest(provider, client, auth, session.refreshToken, options)
          const tokens = await oauth.processRefreshTokenResponse(provider, client, tokenResponse)
          await saveTokens(identity, tokens, session.refreshToken)
          session.accessToken = tokens.access_token
        } catch (error) {
          if (error instanceof oauth.ResponseBodyError && error.error === 'invalid_grant') {
            await store.delete(sessionKey)
            if (browsing) throw new RequestError(401, 'Your session expired. Sign in again.')
            return response(200, { authenticated: false })
          }
          throw error
        }
      }
      if (browsing) {
        const id = value => {
          if (typeof value !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(value)) throw new RequestError(400, 'Invalid selection ID.')
          return value
        }
        const cursor = params.cursor
        if (cursor !== undefined && (typeof cursor !== 'string' || cursor.length === 0 || cursor.length > 4096)) throw new RequestError(400, 'Invalid pagination cursor.')
        async function request(path, paginated = false, init = {}) {
          const url = new URL(path, 'https://api.frame.io')
          if (paginated) {
            url.searchParams.set('page_size', '50')
            if (path !== '/v4/accounts') url.searchParams.set('sort', 'name_asc')
            if (cursor) url.searchParams.set('after', cursor)
          }
          const result = await fetchImpl(url, { ...init, method: init.method || 'GET', headers: { Authorization: 'Bearer ' + session.accessToken, Accept: 'application/json', ...(init.body ? { 'Content-Type': 'application/json' } : {}) }, redirect: 'error', signal: AbortSignal.timeout(15000) })
          if (!result.ok) {
            if (result.status === 401) throw new RequestError(401, 'Frame.io requires a new sign-in.')
            if (result.status === 403) throw new RequestError(403, 'You do not have access to this Frame.io selection.')
            if (result.status === 404) throw new RequestError(404, 'This Frame.io selection is no longer available.')
            if (result.status === 429) throw new RequestError(429, 'Frame.io is busy. Wait a moment and retry.')
            throw new RequestError(502, 'Frame.io could not complete this request.')
          }
          return result.json()
        }
        if (params.operation === 'presentation-lookup') {
          if (typeof params.presentationId !== 'string' || !uuid.test(params.presentationId)) throw new RequestError(400, 'Invalid presentation tracking ID.')
          const registry = await getRegistry()
          const entry = await registry.latest(params.presentationId)
          return response(200, { entry: entry ? await verifyAsset(entry, request, RequestError) : null })
        }
        if (params.operation === 'upload-start' || params.operation === 'upload-status') {
          return response(200, await uploadOperation({ params, identity, read, write, request, registry: await getRegistry(), RequestError, now }))
        }
        function page(body, path) {
          if (!Array.isArray(body.data)) throw new RequestError(502, 'Frame.io returned an invalid list.')
          let nextCursor = null
          if (body.links?.next) {
            const next = new URL(body.links.next, 'https://api.frame.io')
            if (next.origin !== 'https://api.frame.io' || next.pathname !== path || next.username || next.password || next.hash) throw new RequestError(502, 'Frame.io returned invalid pagination.')
            nextCursor = next.searchParams.get('after')
            if (!nextCursor || nextCursor.length > 4096 || nextCursor === cursor) throw new RequestError(502, 'Frame.io returned invalid pagination.')
          }
          const items = body.data.map(item => {
            const name = path === '/v4/accounts' ? item?.display_name : item?.name
            if (!item || typeof name !== 'string') throw new RequestError(502, 'Frame.io returned an invalid choice.')
            return { id: id(item.id), name }
          })
          return { items, nextCursor }
        }
        if (params.operation === 'accounts') {
          const path = '/v4/accounts'
          return response(200, page(await request(path, true), path))
        }
        const accountId = id(params.accountId)
        if (params.operation === 'projects') {
          const path = `/v4/accounts/${accountId}/projects`
          return response(200, page(await request(path, true), path))
        }
        const projectId = id(params.projectId)
        const project = (await request(`/v4/accounts/${accountId}/projects/${projectId}`)).data
        if (!project || typeof project.name !== 'string') throw new RequestError(502, 'Frame.io returned an invalid project.')
        const rootId = id(project.root_folder_id)
        const folderId = params.folderId === undefined ? rootId : id(params.folderId)
        let folder = { id: rootId, name: 'Project root', parentId: null }
        if (folderId !== rootId) {
          const detail = (await request(`/v4/accounts/${accountId}/folders/${folderId}`)).data
          if (!detail || detail.project_id !== projectId) throw new RequestError(403, 'This folder does not belong to the selected project.')
          if (typeof detail.name !== 'string') throw new RequestError(502, 'Frame.io returned an invalid folder.')
          folder = { id: folderId, name: detail.name, parentId: detail.parent_id ? id(detail.parent_id) : null }
        }
        const path = `/v4/accounts/${accountId}/folders/${folderId}/folders`
        const body = await request(path, true)
        if (!Array.isArray(body.data) || body.data.some(item => item.project_id !== projectId)) throw new RequestError(502, 'Frame.io returned folders outside the selected project.')
        return response(200, { ...page(body, path), folder, rootId })
      }
      return response(200, { authenticated: true })
    } catch (error) {
      if (error instanceof RequestError) return response(error.status, { error: error.message })
      if (params?.operation === 'presentation-lookup') return response(502, { error: 'Presentation history is unavailable. Retry before uploading.' })
      if (['upload-start', 'upload-status'].includes(params?.operation)) return response(502, { error: 'Frame.io upload or version registration could not be confirmed. Check the destination before starting again.' })
      return response(502, { error: ['accounts', 'projects', 'folders'].includes(params?.operation)
        ? 'Frame.io choices could not be loaded. Retry.' : 'Sign-in could not be completed. Try again or start a new login.' })
    }
  }
}

exports.main = createHandler()
exports.createHandler = createHandler
import publicConfig from './public-config.json'

const storageKey = 'frameio.app-session'

export type Choice = { id: string; name: string }
export type FolderChoice = Choice & { parentId: string | null }
export type BrowsingOperation = 'accounts' | 'projects' | 'folders'
export type ChoicePage = { items: Choice[]; nextCursor: string | null; folder?: FolderChoice; rootId?: string }
export type Destination = { accountId: string; projectId: string; folderId: string }
export type RegistryEntry = Destination & { presentationId: string; documentType?: 'pptx' | 'docx'; uploadId: string; name: string; fileSize: number; fileId: string; stackId?: string; viewUrl: string; state: 'complete'; version: number; createdAt: string; completedAt: string }

export class ApiError extends Error {
  constructor(message: string, public status: number, public relinkAllowed = false) { super(message) }
}

export function appSession() {
  let persistent: Storage | undefined
  let value: string | null = null
  try { persistent = localStorage; value = persistent.getItem(storageKey) } catch { persistent = undefined }
  if (value && /^[a-f0-9]{64}$/.test(value)) {
    try { sessionStorage.removeItem(storageKey) } catch {}
    return value
  }
  try { value = sessionStorage.getItem(storageKey) } catch { value = null }
  if (!value || !/^[a-f0-9]{64}$/.test(value)) {
    value = Array.from(crypto.getRandomValues(new Uint8Array(32)), byte => byte.toString(16).padStart(2, '0')).join('')
  }
  if (persistent) {
    try {
      persistent.setItem(storageKey, value)
      try { sessionStorage.removeItem(storageKey) } catch {}
      return value
    } catch {}
  }
  try { sessionStorage.setItem(storageKey, value) } catch { throw new Error('Browser session storage is unavailable. Sign-in cannot be retained.') }
  return value
}

export function clearAppSession() {
  try { localStorage.removeItem(storageKey) } catch {}
  try { sessionStorage.removeItem(storageKey) } catch {}
}

export async function oauthRequest(operation: 'start' | 'callback' | 'session' | 'logout' | BrowsingOperation | 'upload-start' | 'upload-status' | 'presentation-lookup', data: Record<string, string | number> = {}, signal?: AbortSignal) {
  const local = ['localhost', '127.0.0.1', 'www.localappbuilder.com'].includes(location.hostname)
  const url = new URL(local ? '/api/oauth' : publicConfig.oauthUrl, location.href)
  const allowed = local ? url.origin === location.origin && url.pathname === '/api/oauth'
    : url.origin === location.origin && url.origin === 'https://27200-msofficetoframe-stage.adobeio-static.net' && url.pathname === '/api/v1/web/msofficetoframe/oauth'
  if (!allowed || url.username || url.password || url.search || url.hash) throw new Error('Invalid authentication endpoint.')
  const response = await fetch(url, {
    method: 'POST', credentials: 'omit', referrerPolicy: 'no-referrer',
    headers: { 'Content-Type': 'application/json', ...(operation === 'callback' ? {} : { Authorization: 'Bearer ' + appSession() }) },
    body: JSON.stringify({ operation, ...data }),
    signal: signal ? AbortSignal.any([signal, AbortSignal.timeout(20000)]) : AbortSignal.timeout(20000)
  })
  const body = await response.json().catch(() => {
    throw new ApiError(response.ok ? 'Authentication service returned an invalid response.' : `Authentication service returned HTTP ${response.status} instead of JSON.`, response.ok ? 502 : response.status)
  })
  if (!response.ok) throw new ApiError(typeof body.error === 'string' ? body.error : 'Authentication request failed.', response.status, operation === 'presentation-lookup' && response.status === 404 && body.relinkAllowed === true)
  return body as { authenticated?: boolean; authorizationUrl?: string; revoked?: boolean; fileId?: string; parts?: { url: string; size: number }[]; complete?: boolean; failed?: boolean; entry?: RegistryEntry | null } & Partial<ChoicePage>
}

export async function loadChoices(operation: BrowsingOperation, data: Record<string, string>, progress: (items: Choice[]) => void, signal: AbortSignal): Promise<ChoicePage> {
  const items = new Map<string, Choice>()
  const cursors = new Set<string>()
  let cursor: string | null = null
  let first: Partial<ChoicePage> | undefined
  do {
    if (signal.aborted) throw new Error('Request cancelled.')
    const result = await oauthRequest(operation, { ...data, ...(cursor ? { cursor } : {}) }, signal)
    if (!Array.isArray(result.items) || result.items.some(item => typeof item.id !== 'string' || typeof item.name !== 'string') || (result.nextCursor !== null && typeof result.nextCursor !== 'string')) throw new Error('Frame.io returned invalid choices.')
    if (!first) first = result
    for (const item of result.items) items.set(item.id, item)
    if (signal.aborted) throw new Error('Request cancelled.')
    progress([...items.values()].sort((left, right) => left.name.localeCompare(right.name)))
    cursor = result.nextCursor
    if (cursor) {
      if (cursors.has(cursor) || cursors.size >= 1000) throw new Error('Frame.io returned too many result pages. Retry or choose a different account.')
      cursors.add(cursor)
    }
  } while (cursor)
  return { items: [...items.values()].sort((left, right) => left.name.localeCompare(right.name)), nextCursor: null, folder: first?.folder, rootId: first?.rootId }
}

export async function signIn(inOffice: boolean, signal: AbortSignal) {
  const startPage = new URL('./oauth-start.html', location.href)
  let popup: Window | null = null
  let dialog: Office.Dialog | undefined
  let dialogClosed = false
  let callbackFailed = false
  function receive(message: unknown) {
    if (typeof message !== 'string') return
    try {
      const data = JSON.parse(message)
      if (data.type === 'frameio-oauth-complete' && data.ok === false) callbackFailed = true
    } catch {}
  }
  const onMessage = (event: MessageEvent) => {
    if (event.origin === location.origin && event.source === popup) receive(event.data)
  }
  window.addEventListener('message', onMessage)
  if (!inOffice) {
    popup = window.open('about:blank', 'frameio-sign-in', 'popup,width=560,height=720')
    if (!popup) { window.removeEventListener('message', onMessage); throw new Error('The sign-in window was blocked.') }
  }
  try {
    const result = await oauthRequest('start', {}, signal)
    if (!result.authorizationUrl) throw new Error('Authorization URL is missing.')
    const authorization = new URL(result.authorizationUrl)
    if (authorization.origin !== 'https://ims-na1.adobelogin.com' || authorization.pathname !== '/ims/authorize/v2') throw new Error('Invalid authorization URL.')
    startPage.hash = new URLSearchParams({ authorizationUrl: authorization.href }).toString()
    if (inOffice) {
      await new Promise<void>((resolve, reject) => Office.context.ui.displayDialogAsync(startPage.href, { height: 65, width: 45, displayInIframe: false }, result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) { reject(new Error('Office could not open the sign-in dialog.')); return }
        dialog = result.value
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => { dialogClosed = true })
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, event => {
          if ('message' in event && (!('origin' in event) || event.origin === location.origin)) receive(event.message)
        })
        resolve()
      }))
    } else popup!.location.replace(startPage.href)
    const deadline = Date.now() + 240000
    while (Date.now() < deadline) {
      if (signal.aborted) throw new Error('Sign-in cancelled.')
      if (callbackFailed) throw new Error('Sign-in was cancelled or could not be completed.')
      const status = await oauthRequest('session', {}, signal)
      if (status.authenticated) return
      if (dialogClosed) throw new Error('The sign-in dialog was closed.')
      await new Promise(resolve => setTimeout(resolve, 1500))
    }
    throw new Error('Sign-in timed out. Start a new login.')
  } finally {
    window.removeEventListener('message', onMessage)
    dialog?.close()
    popup?.close()
  }
}
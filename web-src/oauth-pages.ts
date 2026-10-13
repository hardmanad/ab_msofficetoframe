import { oauthRequest } from './auth'

const status = document.querySelector<HTMLElement>('[role="status"]')!

function notify(ok: boolean) {
  const message = JSON.stringify({ type: 'frameio-oauth-complete', ok })
  if (window.opener) window.opener.postMessage(message, location.origin)
  const send = () => Office.onReady().then(info => {
    if (info.host) Office.context.ui.messageParent(message, { targetOrigin: location.origin })
  }).catch(() => {})
  if (typeof Office !== 'undefined') void send()
  else {
    const script = document.createElement('script')
    script.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js'
    script.referrerPolicy = 'no-referrer'
    script.onload = () => { void send() }
    document.head.appendChild(script)
  }
}

async function initialize() {
  if (location.pathname.endsWith('/oauth-start.html')) {
    const value = new URLSearchParams(location.hash.slice(1)).get('authorizationUrl')
    if (!value) return
    const url = new URL(value)
    if (url.origin !== 'https://ims-na1.adobelogin.com' || url.pathname !== '/ims/authorize/v2' || url.username || url.password) throw new Error('Invalid sign-in destination.')
    location.replace(url.href)
    return
  }
  const query = new URLSearchParams(location.search)
  const state = query.get('state')
  const code = query.get('code')
  const error = query.get('error')
  const clean = () => {
    if (typeof history.replaceState === 'function') history.replaceState(null, '', location.pathname)
  }
  clean()
  if (!state || (!code && !error) || query.getAll('state').length !== 1 || query.getAll('code').length > 1 || query.getAll('error').length > 1) throw new Error('Missing or invalid sign-in result. Start sign-in from the task pane.')
  const result = await oauthRequest('callback', { state, ...(error ? { error } : { code: code! }) })
  if (!result.authenticated) throw new Error('Sign-in was not completed.')
  status.textContent = 'Signed in. You can close this window.'
  notify(true)
}

initialize().catch(error => {
  status.textContent = error instanceof Error ? error.message : 'Sign-in failed. Start a new login.'
  if (location.pathname.endsWith('/oauth-callback.html')) notify(false)
})
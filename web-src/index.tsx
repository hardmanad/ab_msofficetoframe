import { useEffect, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { ChevronDown, CircleCheck, CircleX, FileText, ImageOff, LoaderCircle, LogIn, LogOut, X } from 'lucide-react'
import { clearAppSession, Destination, RegistryEntry, oauthRequest, signIn } from './auth'
import { DestinationPicker } from './DestinationPicker'
import { UploadPanel } from './UploadPanel'
import { getOfficeHost, OfficeHost } from './office-document'
import './styles.css'

type AccountState = 'checking' | 'signed-out' | 'signed-in' | 'signing-in'

function ConnectionMenu({ account, disabled, connect, disconnect, cancel }: { account: AccountState; disabled: boolean; connect: () => void; disconnect: () => void; cancel: () => void }) {
  const [open, setOpen] = useState(false)
  const root = useRef<HTMLDivElement>(null)
  const trigger = useRef<HTMLButtonElement>(null)
  const option = useRef<HTMLButtonElement>(null)
  const connected = account === 'signed-in'
  const waiting = account === 'checking' || account === 'signing-in'
  useEffect(() => {
    if (!open) return
    option.current?.focus()
    const close = (event: PointerEvent) => { if (!root.current?.contains(event.target as Node)) setOpen(false) }
    document.addEventListener('pointerdown', close)
    return () => document.removeEventListener('pointerdown', close)
  }, [open])
  useEffect(() => { if (disabled || account === 'checking') setOpen(false) }, [disabled, account])
  return <div className="connection-menu" ref={root} onBlur={event => { if (event.relatedTarget && !event.currentTarget.contains(event.relatedTarget)) setOpen(false) }} onKeyDown={event => {
    if (event.key === 'Escape') { event.preventDefault(); setOpen(false); trigger.current?.focus() }
    if (open && ['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) { event.preventDefault(); option.current?.focus() }
  }}>
    <button className={`connection-trigger ${connected ? 'is-connected' : ''}`} ref={trigger} disabled={disabled || account === 'checking'} aria-haspopup="menu" aria-expanded={open} aria-controls={open ? 'connection-options' : undefined} onClick={() => setOpen(value => !value)} onKeyDown={event => { if (event.key === 'ArrowDown' && !open) { event.preventDefault(); setOpen(true) } }}>
      {waiting ? <LoaderCircle size={16} className="spin" aria-hidden="true" /> : connected ? <CircleCheck size={16} aria-hidden="true" /> : <CircleX size={16} aria-hidden="true" />}
      <span>{account === 'checking' ? 'Checking account' : account === 'signing-in' ? 'Signing in' : connected ? 'Connected' : 'Not connected'}</span><ChevronDown size={14} aria-hidden="true" />
    </button>
    {open && <div className="connection-options" role="menu" id="connection-options" aria-label="Connection actions"><button ref={option} role="menuitem" onMouseDown={event => event.preventDefault()} onClick={() => {
      setOpen(false)
      trigger.current?.focus()
      if (account === 'signing-in') cancel()
      else if (connected) disconnect()
      else connect()
    }}>{account === 'signing-in' ? <X size={16} aria-hidden="true" /> : connected ? <LogOut size={16} aria-hidden="true" /> : <LogIn size={16} aria-hidden="true" />}{account === 'signing-in' ? 'Cancel sign-in' : connected ? 'Sign out' : 'Sign in to Frame.io'}</button></div>}
  </div>
}

function App() {
  const [host, setHost] = useState<OfficeHost>({ kind: 'loading', adapter: null, available: false, reason: 'Checking Office host...', inOffice: false })
  const [account, setAccount] = useState<AccountState>('checking')
  const [preview, setPreview] = useState<{ image: string | null; name: string; loading: boolean }>({ image: null, name: '', loading: true })
  const [authMessage, setAuthMessage] = useState('')
  const [destination, setDestination] = useState<Destination | null>(null)
  const [uploadBusy, setUploadBusy] = useState(false)
  const [linkedAsset, setLinkedAsset] = useState<RegistryEntry | null>(null)
  const [identity, setIdentity] = useState<{ id: string | null; checking: boolean; error: string; created: boolean }>({ id: null, checking: false, error: '', created: false })
  const loginAbort = useRef<AbortController | null>(null)

  useEffect(() => {
    let active = true
    oauthRequest('session').then(result => { if (active) setAccount(result.authenticated ? 'signed-in' : 'signed-out') }).catch(error => {
      if (active) { setAccount('signed-out'); setAuthMessage(error instanceof Error ? error.message : 'Authentication unavailable.') }
    })
    return () => { active = false; loginAbort.current?.abort() }
  }, [])

  async function connect() {
    const controller = new AbortController()
    loginAbort.current = controller
    setAccount('signing-in')
    setAuthMessage('')
    try {
      await signIn(host.inOffice, controller.signal)
      setAccount('signed-in')
    } catch (error) {
      setAccount('signed-out')
      clearAppSession()
      setAuthMessage(controller.signal.aborted ? 'Sign-in cancelled.' : error instanceof Error ? error.message : 'Sign-in failed.')
    } finally { loginAbort.current = null }
  }

  async function disconnect() {
    setDestination(null)
    setLinkedAsset(null)
    setAccount('checking')
    setAuthMessage('')
    try {
      const result = await oauthRequest('logout')
      clearAppSession()
      setAccount('signed-out')
      if (!result.revoked) setAuthMessage('Signed out locally. Provider revocation could not be confirmed.')
    } catch {
      setAccount('signed-in')
      setAuthMessage('Sign-out could not be completed. Try again.')
    }
  }

  useEffect(() => {
    let active = true
    const timeout = window.setTimeout(() => {
      if (active) setHost(getOfficeHost({ host: null, platform: null }))
    }, 4000)
    if (typeof Office !== 'undefined') {
      Office.onReady().then(async info => {
        if (!active) return
        window.clearTimeout(timeout)
        await getOfficeHost(info).adapter?.loadMetadata?.()
        if (active) setHost(getOfficeHost(info))
      }).catch(() => {
        if (active) setHost(getOfficeHost({ host: null, platform: null }))
      })
    }
    return () => { active = false; window.clearTimeout(timeout) }
  }, [])

  useEffect(() => {
    if (host.kind === 'loading') return
    if (!host.adapter?.preview) { setPreview({ image: null, name: host.adapter?.filename() || 'No document', loading: false }); return }
    let active = true
    const name = host.adapter.filename()
    setPreview({ image: null, name, loading: true })
    host.adapter.preview().then(image => { if (active) setPreview({ image, name, loading: false }) }).catch(() => {
      if (active) setPreview({ image: null, name, loading: false })
    })
    return () => { active = false }
  }, [host.adapter, host.kind])

  useEffect(() => {
    if (!host.adapter || !host.available) return
    let active = true
    let created = false
    setIdentity({ id: null, checking: true, error: '', created: false })
    host.adapter.ensureId(() => { created = true; return crypto.randomUUID() }).then(id => {
      if (active) setIdentity({ id, checking: false, error: '', created })
    }).catch(error => {
      if (active) setIdentity({ id: null, checking: false, error: error?.code === 'AccessDenied' ? 'Tracking metadata requires document editing and the updated add-in manifest with document-write permission.' : error instanceof Error ? error.message : 'Document tracking metadata could not be accessed.', created: false })
    })
    return () => { active = false }
  }, [host.adapter, host.available])

  useEffect(() => {
    const isSaved = host.adapter?.isSaved
    if (!identity.created || !isSaved) return
    let active = true
    let timeout: number | undefined
    async function checkSaved() {
      try {
        if (await isSaved!()) {
          if (active) setIdentity(value => value.id === identity.id ? { ...value, created: false } : value)
          return
        }
      } catch {}
      if (active) timeout = window.setTimeout(checkSaved, 2000)
    }
    void checkSaved()
    return () => { active = false; window.clearTimeout(timeout) }
  }, [host.adapter, identity.id, identity.created])

  const documentNotice = identity.error || (identity.checking ? `Checking ${host.adapter?.noun || 'document'} identity...` : identity.created ? `Save the ${host.adapter?.noun || 'document'} to retain its tracking ID.` : '')

  return <main>
    <header>
      <img src={new URL('./assets/frameio.png', import.meta.url).href} width="36" height="36" alt="" />
      <div><h1>MSOffice to Frame.io</h1><p>{host.kind === 'word' ? 'Word' : host.kind === 'powerpoint' ? 'PowerPoint' : 'Office'} / Frame.io</p></div>
    </header>
    <section className="document">
      {host.kind === 'powerpoint' ? <figure className="presentation-preview"><div className="slide-preview" aria-busy={preview.loading}>
        {preview.image ? <img src={preview.image} alt="First slide preview" onError={() => setPreview(value => ({ ...value, image: null }))} />
          : preview.loading ? <LoaderCircle size={24} className="spin" aria-label="Loading slide preview" /> : <ImageOff size={28} strokeWidth={1.5} aria-label="Slide preview unavailable" />}
      </div><figcaption>{preview.name || 'Loading presentation...'}</figcaption></figure>
        : <div className="document-name"><FileText size={20} aria-hidden="true" /><p>{preview.name || host.adapter?.filename() || 'Loading document...'}</p></div>}
    </section>
    <section className="destination">
      <div className="section-heading"><h2>Frame.io</h2><ConnectionMenu account={account} disabled={uploadBusy || host.kind === 'loading'} connect={connect} disconnect={disconnect} cancel={() => loginAbort.current?.abort()} /></div>
      {authMessage && <p className="auth-error" role="alert">{authMessage}</p>}
      {account === 'signing-in' && <p className="detail" role="status">Waiting for sign-in...</p>}
      {account === 'signed-in' ? <>{linkedAsset ? <dl><div><dt>Asset</dt><dd>{linkedAsset.name}</dd></div><div><dt>Version</dt><dd>{linkedAsset.version}</dd></div></dl>
        : <DestinationPicker disabled={uploadBusy} onSelection={setDestination} onSessionExpired={() => { clearAppSession(); setDestination(null); setLinkedAsset(null); setAccount('signed-out'); setAuthMessage('Your Frame.io session expired. Sign in again.') }} />}
        <UploadPanel adapter={host.adapter} unavailableReason={host.reason} documentNotice={documentNotice} documentError={!!identity.error} presentationId={identity.id} onAssetChange={setLinkedAsset} destination={destination} available={host.available} signedIn onBusyChange={setUploadBusy} onSessionExpired={() => { clearAppSession(); setDestination(null); setLinkedAsset(null); setAccount('signed-out'); setAuthMessage('Your Frame.io session expired. Sign in again.') }} /></>
        : <dl><div><dt>Account</dt><dd>Not connected</dd></div><div><dt>Project</dt><dd>Not selected</dd></div><div><dt>Folder</dt><dd>Not selected</dd></div></dl>}
      {account !== 'signed-in' && <UploadPanel adapter={host.adapter} unavailableReason={host.reason} documentNotice={documentNotice} documentError={!!identity.error} presentationId={identity.id} onAssetChange={setLinkedAsset} destination={null} available={host.available} signedIn={false} onBusyChange={setUploadBusy} onSessionExpired={() => {}} />}
    </section>
    <footer>MSOffice to Frame.io <span>1.0.0</span></footer>
  </main>
}

createRoot(document.getElementById('root')!).render(<App />)
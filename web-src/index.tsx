import { useEffect, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { ChevronDown, CircleCheck, CircleX, ImageOff, LoaderCircle, LogIn, LogOut, X } from 'lucide-react'
import { clearAppSession, Destination, RegistryEntry, oauthRequest, signIn } from './auth'
import { DestinationPicker } from './DestinationPicker'
import { UploadPanel } from './UploadPanel'
import { ensurePresentationId, firstSlidePreview, presentationFilename } from './presentation'
import './styles.css'

type HostState = { kind: 'loading' | 'browser' | 'powerpoint' | 'other'; compressed: boolean }
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
  const [host, setHost] = useState<HostState>({ kind: 'loading', compressed: false })
  const [account, setAccount] = useState<AccountState>('checking')
  const [preview, setPreview] = useState<{ image: string | null; name: string; loading: boolean }>({ image: null, name: '', loading: true })
  const [authMessage, setAuthMessage] = useState('')
  const [destination, setDestination] = useState<Destination | null>(null)
  const [uploadBusy, setUploadBusy] = useState(false)
  const [linkedAsset, setLinkedAsset] = useState<RegistryEntry | null>(null)
  const [presentation, setPresentation] = useState<{ id: string | null; checking: boolean; error: string; created: boolean }>({ id: null, checking: false, error: '', created: false })
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
      await signIn(host.kind === 'powerpoint', controller.signal)
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
      if (active) setHost({ kind: 'browser', compressed: false })
    }, 4000)
    if (typeof Office !== 'undefined') {
      Office.onReady().then(info => {
        if (!active) return
        window.clearTimeout(timeout)
        setHost({
          kind: info.host === Office.HostType.PowerPoint ? 'powerpoint' : info.host ? 'other' : 'browser',
          compressed: info.host === Office.HostType.PowerPoint && Office.context.requirements.isSetSupported('CompressedFile', '1.1')
        })
      }).catch(() => {
        if (active) setHost({ kind: 'browser', compressed: false })
      })
    }
    return () => { active = false; window.clearTimeout(timeout) }
  }, [])

  useEffect(() => {
    if (host.kind === 'loading') return
    if (host.kind !== 'powerpoint') { setPreview({ image: null, name: 'No presentation', loading: false }); return }
    let active = true
    const name = presentationFilename()
    setPreview({ image: null, name, loading: true })
    firstSlidePreview().then(image => { if (active) setPreview({ image, name, loading: false }) }).catch(() => {
      if (active) setPreview({ image: null, name, loading: false })
    })
    return () => { active = false }
  }, [host.kind])

  useEffect(() => {
    if (host.kind !== 'powerpoint' || !host.compressed) return
    let active = true
    let created = false
    setPresentation({ id: null, checking: true, error: '', created: false })
    ensurePresentationId(() => { created = true; return crypto.randomUUID() }).then(id => {
      if (active) setPresentation({ id, checking: false, error: '', created })
    }).catch(error => {
      if (active) setPresentation({ id: null, checking: false, error: error?.code === 'AccessDenied' ? 'Tracking metadata requires the updated add-in manifest with document-write permission.' : error instanceof Error ? error.message : 'Presentation tracking metadata could not be accessed.', created: false })
    })
    return () => { active = false }
  }, [host.kind, host.compressed])

  const documentNotice = presentation.error || (presentation.checking ? 'Checking presentation identity...' : presentation.created ? 'Save the presentation to retain its tracking ID.' : '')

  return <main>
    <header>
      <img src={new URL('./assets/frameio.png', import.meta.url).href} width="36" height="36" alt="" />
      <div><h1>MSOffice to Frame.io</h1><p>PowerPoint / Frame.io</p></div>
    </header>
    <section className="document">
      <figure className="presentation-preview"><div className="slide-preview" aria-busy={preview.loading}>
        {preview.image ? <img src={preview.image} alt="First slide preview" onError={() => setPreview(value => ({ ...value, image: null }))} />
          : preview.loading ? <LoaderCircle size={24} className="spin" aria-label="Loading slide preview" /> : <ImageOff size={28} strokeWidth={1.5} aria-label="Slide preview unavailable" />}
      </div><figcaption>{preview.name || 'Loading presentation...'}</figcaption></figure>
    </section>
    <section className="destination">
      <div className="section-heading"><h2>Frame.io</h2><ConnectionMenu account={account} disabled={uploadBusy || host.kind === 'loading'} connect={connect} disconnect={disconnect} cancel={() => loginAbort.current?.abort()} /></div>
      {authMessage && <p className="auth-error" role="alert">{authMessage}</p>}
      {account === 'signing-in' && <p className="detail" role="status">Waiting for sign-in...</p>}
      {account === 'signed-in' ? <>{linkedAsset ? <dl><div><dt>Asset</dt><dd>{linkedAsset.name}</dd></div><div><dt>Version</dt><dd>{linkedAsset.version}</dd></div></dl>
        : <DestinationPicker disabled={uploadBusy} onSelection={setDestination} onSessionExpired={() => { clearAppSession(); setDestination(null); setLinkedAsset(null); setAccount('signed-out'); setAuthMessage('Your Frame.io session expired. Sign in again.') }} />}
        <UploadPanel documentNotice={documentNotice} documentError={!!presentation.error} presentationId={presentation.id} onAssetChange={setLinkedAsset} destination={destination} available={host.kind === 'powerpoint' && host.compressed} signedIn onBusyChange={setUploadBusy} onSessionExpired={() => { clearAppSession(); setDestination(null); setLinkedAsset(null); setAccount('signed-out'); setAuthMessage('Your Frame.io session expired. Sign in again.') }} /></>
        : <dl><div><dt>Account</dt><dd>Not connected</dd></div><div><dt>Project</dt><dd>Not selected</dd></div><div><dt>Folder</dt><dd>Not selected</dd></div></dl>}
      {account !== 'signed-in' && <UploadPanel documentNotice={documentNotice} documentError={!!presentation.error} presentationId={presentation.id} onAssetChange={setLinkedAsset} destination={null} available={false} signedIn={false} onBusyChange={setUploadBusy} onSessionExpired={() => {}} />}
    </section>
    <footer>MSOffice to Frame.io <span>1.0.0</span></footer>
  </main>
}

createRoot(document.getElementById('root')!).render(<App />)
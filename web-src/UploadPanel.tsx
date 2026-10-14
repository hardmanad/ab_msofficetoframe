import { useEffect, useRef, useState } from 'react'
import { ArrowUpToLine, CircleCheck, ExternalLink, RefreshCw, X } from 'lucide-react'
import { ApiError, Destination, RegistryEntry, oauthRequest } from './auth'
import { ByteSource, delay, maxFileSize, openDocument, putPart, transferParts } from './upload-transfer'
import { DocumentType, initialUploadName, OfficeDocument, validUploadName, verifyDocumentIdentity } from './office-document'

type Phase = 'idle' | 'exporting' | 'starting' | 'transferring' | 'verifying' | 'complete' | 'error' | 'cancelled' | 'pending'
type UploadState = { phase: Phase; loaded: number; size: number; message: string }
const idle: UploadState = { phase: 'idle', loaded: 0, size: 0, message: '' }
const busyPhases: Phase[] = ['exporting', 'starting', 'transferring', 'verifying']
const labels: Record<Phase, string> = { idle: '', exporting: 'Preparing document', starting: 'Creating Frame.io file', transferring: 'Uploading', verifying: 'Confirming upload', complete: 'Upload complete', error: 'Upload failed', cancelled: 'Upload cancelled', pending: 'Confirmation pending' }
const formatBytes = (value: number) => `${(value / (1024 * 1024)).toFixed(1)} MiB`

function registryEntry(value: RegistryEntry | null | undefined, presentationId: string, documentType: DocumentType): RegistryEntry | null {
  if (value === null) return null
  if (!value || value.presentationId !== presentationId || value.state !== 'complete' || !Number.isSafeInteger(value.version) || value.version < 1) throw new Error('Invalid document history.')
  if ((value.documentType === undefined ? 'pptx' : value.documentType) !== documentType) throw new Error('The linked history belongs to a different document type.')
  const url = new URL(value.viewUrl)
  if (url.protocol !== 'https:' || (url.hostname !== 'frame.io' && !url.hostname.endsWith('.frame.io')) || url.username || url.password || url.hash || (url.port && url.port !== '443')) throw new Error('Invalid Frame.io link.')
  return value
}

export function UploadPanel({ destination, presentationId, adapter, unavailableReason, available, signedIn, onBusyChange, onSessionExpired, onAssetChange, documentNotice, documentError }: { destination: Destination | null; presentationId: string | null; adapter: OfficeDocument | null; unavailableReason?: string; available: boolean; signedIn: boolean; onBusyChange: (busy: boolean) => void; onSessionExpired: () => void; onAssetChange: (entry: RegistryEntry | null) => void; documentNotice?: string; documentError?: boolean }) {
  const [name, setName] = useState(() => adapter ? initialUploadName(adapter) : '')
  useEffect(() => { setName(adapter ? initialUploadName(adapter) : '') }, [adapter])
  const [state, setState] = useState<UploadState>(idle)
  const [allowNew, setAllowNew] = useState(false)
  const [entry, setEntry] = useState<RegistryEntry | null>(null)
  const [history, setHistory] = useState({ loading: false, checked: false, error: '' })
  const [historyRetry, setHistoryRetry] = useState(0)
  const expired = useRef(onSessionExpired)
  useEffect(() => { expired.current = onSessionExpired }, [onSessionExpired])
  const abort = useRef<AbortController | null>(null)
  const uploadId = useRef<string | null>(null)
  const remoteCreated = useRef(false)
  const uncertainStart = useRef(false)
  const allBytesSent = useRef(false)
  const busy = busyPhases.includes(state.phase)
  const mustConfirmNew = uncertainStart.current || (remoteCreated.current && state.phase !== 'complete')
  const validName = !!adapter && validUploadName(name, adapter.documentType)
  useEffect(() => () => { abort.current?.abort() }, [])
  useEffect(() => {
    setEntry(null)
    onAssetChange(null)
    if (!signedIn || !presentationId || !adapter) { setHistory({ loading: false, checked: false, error: '' }); return }
    const controller = new AbortController()
    setHistory({ loading: true, checked: false, error: '' })
    oauthRequest('presentation-lookup', { presentationId }, controller.signal).then(result => {
      if (controller.signal.aborted) return
      const found = registryEntry(result.entry, presentationId, adapter.documentType)
      setEntry(found)
      onAssetChange(found)
      setHistory({ loading: false, checked: true, error: '' })
    }).catch(error => {
      if (controller.signal.aborted) return
      if (error instanceof ApiError && error.status === 401) expired.current()
      setHistory({ loading: false, checked: false, error: error instanceof Error ? error.message : 'Document history is unavailable.' })
    })
    return () => controller.abort()
  }, [presentationId, signedIn, adapter, historyRetry, onAssetChange])

  async function confirm(signal: AbortSignal) {
    for (let attempt = 0; attempt < 24; attempt++) {
      signal.throwIfAborted()
      try {
        const result = await oauthRequest('upload-status', { uploadId: uploadId.current! }, signal)
        if (result.failed) throw new ApiError('Frame.io reported an upload failure.', 422)
        if (result.complete) {
          const found = registryEntry(result.entry, presentationId!, adapter!.documentType)
          if (!found) throw new Error('Upload finished but registration is incomplete.')
          setEntry(found)
          onAssetChange(found)
          setHistory({ loading: false, checked: true, error: '' })
          setState(value => ({ ...value, phase: 'complete', message: '' }))
          return
        }
        if (typeof result.complete !== 'boolean' || typeof result.failed !== 'boolean') throw new Error('Invalid upload confirmation.')
      } catch (error) {
        if (!(error instanceof ApiError) || (error.status !== 429 && error.status < 500)) throw error
      }
      await delay(attempt < 5 ? 2000 : 5000, signal)
    }
    setState(value => ({ ...value, phase: 'pending', message: allBytesSent.current ? 'File bytes were sent. Frame.io has not confirmed completion yet.' : 'Frame.io has not confirmed completion. Some file bytes may not have been sent.' }))
  }

  async function run(checkOnly = false) {
    const target = entry ? { accountId: entry.accountId, projectId: entry.projectId, folderId: entry.folderId } : destination
    if (abort.current || !adapter || (!checkOnly && (!target || !available || !signedIn || !validName || !presentationId || !history.checked))) return
    if (!checkOnly && mustConfirmNew && !allowNew) return
    const controller = new AbortController()
    abort.current = controller
    onBusyChange(true)
    let source: ByteSource | undefined
    let sent = allBytesSent.current
    let stage: Phase = checkOnly ? 'verifying' : 'exporting'
    try {
      if (!checkOnly) {
        setAllowNew(false)
        sent = false
        remoteCreated.current = false
        uncertainStart.current = false
        allBytesSent.current = false
        uploadId.current = crypto.randomUUID()
        setState({ ...idle, phase: 'exporting' })
        await verifyDocumentIdentity(adapter, presentationId!)
        controller.signal.throwIfAborted()
        source = await openDocument(controller.signal)
        if (!Number.isSafeInteger(source.size) || source.size < 1 || source.size > maxFileSize) throw new Error('The document must be between 1 byte and 1 GiB.')
        stage = 'starting'
        setState({ phase: 'starting', loaded: 0, size: source.size, message: '' })
        uncertainStart.current = true
        const plan = await oauthRequest('upload-start', { ...target!, presentationId: presentationId!, documentType: adapter.documentType, uploadId: uploadId.current, name, fileSize: source.size }, controller.signal)
        if (!plan.fileId || !plan.parts) throw new Error('Frame.io returned an invalid upload plan.')
        remoteCreated.current = true
        uncertainStart.current = false
        stage = 'transferring'
        setState(value => ({ ...value, phase: 'transferring' }))
        await transferParts(source, plan.parts, loaded => setState(value => ({ ...value, loaded })), controller.signal, (url, bytes, progress, signal) => putPart(url, bytes, progress, signal, adapter.mimeType))
        sent = true
        allBytesSent.current = true
      }
      stage = 'verifying'
      setState(value => ({ ...value, phase: 'verifying', message: '' }))
      await confirm(controller.signal)
    } catch (error) {
      if (error instanceof ApiError && error.status === 401) onSessionExpired()
      if (stage === 'starting' && error instanceof ApiError && [400, 401, 403, 404, 429].includes(error.status)) uncertainStart.current = false
      const terminal = error instanceof ApiError && [403, 404, 409, 422].includes(error.status)
      const message = sent && !terminal ? 'File bytes were sent, but upload/version registration could not be confirmed. Check upload status.' : controller.signal.aborted ? 'Any file already created in Frame.io remains in the destination.' : (error instanceof Error ? error.message : 'Upload could not be completed.') + (remoteCreated.current ? ' A file may remain in Frame.io without being registered as a version.' : uncertainStart.current ? ' Check the destination before starting again; the create request may have succeeded.' : '')
      setState(value => ({ ...value, phase: sent && !terminal ? 'pending' : controller.signal.aborted ? 'cancelled' : 'error', message }))
    } finally {
      if (source) {
        try { await source.close() } catch { setState(value => ({ ...value, message: `${value.message} Office file-handle cleanup could not be confirmed.`.trim() })) }
      }
      abort.current = null
      onBusyChange(false)
    }
  }

  const reason = !available ? unavailableReason || 'Word or PowerPoint document access required' : !signedIn ? 'Sign-in required' : !presentationId ? 'Document tracking ID required' : history.loading ? 'Checking document history' : !history.checked ? 'Document history unavailable' : !entry && !destination ? 'Destination required' : !validName ? `Valid .${adapter?.documentType} filename required` : ''
  return <div className="upload-panel">
    {signedIn && presentationId && (entry || history.error) && <div className="asset-controls">
      {entry && <a className="account-button asset-link" href={entry.viewUrl} target="_blank" rel="noopener noreferrer"><ExternalLink size={16} aria-hidden="true" />Open in Frame.io</a>}
      <button className="account-button" disabled={busy || history.loading} onClick={() => setHistoryRetry(value => value + 1)}><RefreshCw size={16} className={history.loading ? 'spin' : ''} aria-hidden="true" />{entry ? 'Refresh history' : 'Retry history lookup'}</button>
    </div>}
    {history.error && <p className="auth-error" role="alert">{history.error}</p>}
    <div className="picker-field"><label htmlFor="upload-name">Filename</label><div className="picker-control"><input id="upload-name" value={name} maxLength={180} disabled={busy || !adapter} onChange={event => setName(event.target.value)} spellCheck={false} /></div></div>
    {state.phase !== 'idle' && <div className="upload-status" role="status" aria-live="polite">
      <p className={state.phase === 'error' ? 'auth-error' : 'detail'}>{state.phase === 'complete' && <CircleCheck size={16} aria-hidden="true" />}{labels[state.phase]}</p>
      {state.size > 0 && <><progress max={state.size} value={state.loaded} aria-label="Document bytes transferred" /><p className="detail">{formatBytes(state.loaded)} / {formatBytes(state.size)}</p></>}
      {state.message && <p className="detail">{state.message}</p>}
    </div>}
    {reason && <p className="detail" id="upload-reason">{reason}</p>}
    {documentNotice && <p className={documentError ? 'auth-error upload-document-notice' : 'detail upload-document-notice'} role={documentError ? 'alert' : 'status'}>{documentNotice}</p>}
    {!busy && mustConfirmNew && <label className="upload-confirm"><input type="checkbox" checked={allowNew} onChange={event => setAllowNew(event.target.checked)} />Create another file; the previous file may remain.</label>}
    {busy ? <button className="upload cancel-upload" onClick={() => abort.current?.abort()}><X size={18} aria-hidden="true" />Cancel upload</button>
      : <><button className="upload" disabled={!!reason || (mustConfirmNew && !allowNew)} aria-describedby={reason ? 'upload-reason' : undefined} onClick={() => run()}><ArrowUpToLine size={18} aria-hidden="true" />{entry ? 'Add a new version' : 'Upload to Frame.io'}</button>
        {(remoteCreated.current || uncertainStart.current) && state.phase !== 'complete' && <button className="account-button check-upload" onClick={() => run(true)}><RefreshCw size={16} aria-hidden="true" />Check upload status</button>}</>}
  </div>
}
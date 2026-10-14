import { useEffect, useRef, useState } from 'react'
import { CornerLeftUp, RefreshCw, X } from 'lucide-react'
import { ApiError, Choice, Destination, FolderChoice, loadChoices } from './auth'

type LoadState = { loading: boolean; error: string }
const idle: LoadState = { loading: false, error: '' }

export function DestinationPicker({ onSessionExpired, onSelection, disabled = false }: { onSessionExpired: () => void; onSelection: (selection: Destination | null) => void; disabled?: boolean }) {
  const [accounts, setAccounts] = useState<Choice[]>([])
  const [accountId, setAccountId] = useState('')
  const [projects, setProjects] = useState<Choice[]>([])
  const [project, setProject] = useState<Choice | null>(null)
  const [query, setQuery] = useState('')
  const [open, setOpen] = useState(false)
  const [active, setActive] = useState(-1)
  const [folders, setFolders] = useState<Choice[]>([])
  const [folderId, setFolderId] = useState<string | undefined>()
  const [folder, setFolder] = useState<FolderChoice | null>(null)
  const [accountLoad, setAccountLoad] = useState<LoadState>(idle)
  const [projectLoad, setProjectLoad] = useState<LoadState>(idle)
  const [folderLoad, setFolderLoad] = useState<LoadState>(idle)
  const [accountRetry, setAccountRetry] = useState(0)
  const [projectRetry, setProjectRetry] = useState(0)
  const [folderRetry, setFolderRetry] = useState(0)
  const expired = useRef(onSessionExpired)
  useEffect(() => { expired.current = onSessionExpired }, [onSessionExpired])
  useEffect(() => {
    onSelection(accountId && project && folder && !folderLoad.loading && !folderLoad.error ? { accountId, projectId: project.id, folderId: folder.id } : null)
  }, [accountId, project, folder, folderLoad.loading, folderLoad.error, onSelection])
  useEffect(() => { if (disabled) setOpen(false) }, [disabled])
  useEffect(() => {
    if (open && active >= 0) document.getElementById(`project-option-${active}`)?.scrollIntoView({ block: 'nearest' })
  }, [open, active])

  function failure(error: unknown): LoadState {
    if (error instanceof ApiError && error.status === 401) expired.current()
    return { loading: false, error: error instanceof Error ? error.message : 'Choices could not be loaded.' }
  }

  useEffect(() => {
    const controller = new AbortController()
    setAccountLoad({ loading: true, error: '' })
    loadChoices('accounts', {}, setAccounts, controller.signal).then(result => {
      if (!controller.signal.aborted) {
        setAccounts(result.items)
        if (result.items.length === 1) setAccountId(result.items[0].id)
        setAccountLoad(idle)
      }
    }).catch(error => { if (!controller.signal.aborted) setAccountLoad(failure(error)) })
    return () => controller.abort()
  }, [accountRetry])

  useEffect(() => {
    setProjects([])
    setProjectLoad(idle)
    if (!accountId) return
    const controller = new AbortController()
    setProjectLoad({ loading: true, error: '' })
    loadChoices('projects', { accountId }, setProjects, controller.signal).then(result => {
      if (!controller.signal.aborted) { setProjects(result.items); setProjectLoad(idle) }
    }).catch(error => { if (!controller.signal.aborted) setProjectLoad(failure(error)) })
    return () => controller.abort()
  }, [accountId, projectRetry])

  useEffect(() => {
    setFolder(null)
    setFolders([])
    setFolderLoad(idle)
    if (!accountId || !project) return
    const controller = new AbortController()
    setFolderLoad({ loading: true, error: '' })
    loadChoices('folders', { accountId, projectId: project.id, ...(folderId ? { folderId } : {}) }, setFolders, controller.signal).then(result => {
      if (controller.signal.aborted) return
      if (!result.folder || typeof result.folder.id !== 'string' || typeof result.folder.name !== 'string') throw new Error('Frame.io returned an invalid folder.')
      setFolder(result.folder)
      setFolders(result.items)
      setFolderLoad(idle)
    }).catch(error => { if (!controller.signal.aborted) setFolderLoad(failure(error)) })
    return () => controller.abort()
  }, [accountId, project?.id, folderId, folderRetry])

  function clearProject() {
    setProject(null); setQuery(''); setOpen(false); setActive(-1)
    setFolderId(undefined); setFolder(null); setFolders([])
  }

  function chooseProject(choice: Choice) {
    setProject(choice); setQuery(choice.name); setOpen(false); setActive(-1)
    setFolderId(undefined); setFolder(null); setFolders([])
  }

  const matches = projects.filter(choice => choice.name.toLocaleLowerCase().includes(query.trim().toLocaleLowerCase()))
  const visible = matches.slice(0, 50)

  return <fieldset className="destination-picker" disabled={disabled} aria-label="Upload destination">
    <div className="picker-field">
      <label htmlFor="frameio-account">Account</label>
      <div className="picker-control">
        <select id="frameio-account" value={accountId} disabled={accountLoad.loading && accounts.length === 0} onChange={event => { setAccountId(event.target.value); setProjects([]); clearProject() }}>
          <option value="">{accountLoad.loading && accounts.length === 0 ? 'Loading accounts...' : accounts.length === 0 && !accountLoad.error ? 'No accounts available' : 'Select account'}</option>
          {accounts.map(choice => <option key={choice.id} value={choice.id}>{choice.name}</option>)}
        </select>
        {accountLoad.error && <button className="icon-button" title="Retry accounts" aria-label="Retry accounts" onClick={() => setAccountRetry(value => value + 1)}><RefreshCw size={16} /></button>}
      </div>
      {accountLoad.error && <p className="auth-error" role="alert">{accountLoad.error}</p>}
    </div>
    <div className="picker-field">
      <label htmlFor="frameio-project">Project</label>
      <div className="picker-control project-control">
        <input id="frameio-project" role="combobox" autoComplete="off" aria-autocomplete="list" aria-expanded={open} aria-controls="frameio-project-options" aria-activedescendant={open && active >= 0 && active < visible.length ? `project-option-${active}` : undefined} disabled={!accountId} value={query} placeholder={!accountId ? 'Select account first' : projectLoad.loading && projects.length === 0 ? 'Loading projects...' : 'Search projects'}
          onFocus={() => setOpen(true)} onBlur={() => setOpen(false)} onChange={event => { setQuery(event.target.value); setProject(null); setOpen(true); setActive(-1); setFolderId(undefined); setFolder(null); setFolders([]) }}
          onKeyDown={event => {
            if (event.key === 'ArrowDown') { event.preventDefault(); setOpen(true); setActive(value => Math.min(value + 1, visible.length - 1)) }
            else if (event.key === 'ArrowUp') { event.preventDefault(); setOpen(true); setActive(value => Math.max(value - 1, 0)) }
            else if (event.key === 'Enter' && open && active >= 0 && visible[active]) { event.preventDefault(); chooseProject(visible[active]) }
            else if (event.key === 'Escape') setOpen(false)
          }} />
        {(query || project) && <button className="icon-button" title="Clear project" aria-label="Clear project" onClick={clearProject}><X size={16} /></button>}
        {projectLoad.error && <button className="icon-button" title="Retry projects" aria-label="Retry projects" onClick={() => setProjectRetry(value => value + 1)}><RefreshCw size={16} /></button>}
        {open && accountId && <div className="project-suggestions">
          <ul id="frameio-project-options" role="listbox" aria-label="Projects">
            {visible.map((choice, index) => <li id={`project-option-${index}`} key={choice.id} role="option" aria-selected={project?.id === choice.id} className={index === active ? 'active' : ''} onMouseDown={event => event.preventDefault()} onClick={() => chooseProject(choice)}>{choice.name}</li>)}
          </ul>
          {visible.length === 0 && <p>{projectLoad.loading ? 'Searching available projects...' : projectLoad.error ? 'Projects unavailable' : query ? 'No matching projects' : 'No projects available'}</p>}
          {matches.length > 50 && <p>{matches.length} matches. Refine search.</p>}
          {projectLoad.loading && visible.length > 0 && <p>Loading more projects...</p>}
        </div>}
      </div>
      {projectLoad.loading && !open && <p className="detail" role="status">Loading projects...</p>}
      {projectLoad.error && <p className="auth-error" role="alert">{projectLoad.error}</p>}
    </div>
    <div className="picker-field">
      <label htmlFor="frameio-folder">Folder</label>
      <div className="picker-control">
        <select id="frameio-folder" disabled={!project || folderLoad.loading || !folder} value={folder?.id || ''} onChange={event => { setFolderId(event.target.value); setFolder(null) }}>
          {folder ? <option value={folder.id}>{folder.name}</option> : <option value="">{!project ? 'Select project first' : folderLoad.loading ? 'Loading folders...' : 'Folders unavailable'}</option>}
          {folders.filter(choice => choice.id !== folder?.id).map(choice => <option key={choice.id} value={choice.id}>{choice.name}</option>)}
        </select>
        <button className="icon-button" title="Parent folder" aria-label="Parent folder" disabled={!folder?.parentId || folderLoad.loading} onClick={() => { if (folder?.parentId) { setFolderId(folder.parentId); setFolder(null) } }}><CornerLeftUp size={16} /></button>
        {folderLoad.error && <button className="icon-button" title="Retry folders" aria-label="Retry folders" onClick={() => setFolderRetry(value => value + 1)}><RefreshCw size={16} /></button>}
      </div>
      {folderLoad.error && <p className="auth-error" role="alert">{folderLoad.error}</p>}
    </div>
  </fieldset>
}
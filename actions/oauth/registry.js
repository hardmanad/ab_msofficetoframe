'use strict'

const uuid = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/

function createRegistry(files) {
  function name(presentationId, uploadId = 'current') {
    if (!uuid.test(presentationId) || (uploadId !== 'current' && !uuid.test(uploadId))) throw new Error('Invalid registry key.')
    return `presentation-uploads/${presentationId}/${uploadId}.json`
  }
  async function read(path) {
    const entries = await files.list(path)
    if (!entries.length) return undefined
    return JSON.parse((await files.read(path)).toString('utf8'))
  }
  return {
    async latest(presentationId) {
      const entry = await read(name(presentationId))
      return entry?.state === 'reset' ? undefined : entry
    },
    event: (presentationId, uploadId) => read(name(presentationId, uploadId)),
    async reset(presentationId, uploadId, now = Date.now) {
      const current = await read(name(presentationId))
      if (!current || current.uploadId !== uploadId) return false
      await files.write(name(presentationId), JSON.stringify({ state: 'reset', resetAt: new Date(now()).toISOString() }))
      return true
    },
    async record(presentationId, entry) {
      const current = await read(name(presentationId))
      if (entry.state === 'complete' && current?.resetAt && (typeof entry.createdAt !== 'string' || entry.createdAt < current.resetAt)) return false
      if (current?.resetAt) entry = { ...entry, resetAt: current.resetAt }
      await files.write(name(presentationId, entry.uploadId), JSON.stringify(entry))
      if (entry.state === 'complete') {
        if (!current || current.state === 'reset' || entry.version > current.version || (entry.version === current.version && entry.completedAt >= current.completedAt)) await files.write(name(presentationId), JSON.stringify(entry))
      }
    }
  }
}

function validId(value, RequestError) {
  if (typeof value !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(value)) throw new RequestError(502, 'Frame.io returned an invalid asset ID.')
  return value
}

function viewUrl(value, RequestError) {
  let url
  try { url = new URL(value) } catch { throw new RequestError(502, 'Frame.io returned an invalid asset link.') }
  if (url.protocol !== 'https:' || (url.hostname !== 'frame.io' && !url.hostname.endsWith('.frame.io')) || url.username || url.password || url.hash || (url.port && url.port !== '443')) throw new RequestError(502, 'Frame.io returned an unsupported asset link.')
  return url.href
}

async function verifyAsset(entry, request, RequestError) {
  const assetId = validId(entry.stackId || entry.fileId, RequestError)
  const kind = entry.stackId ? 'version_stacks' : 'files'
  const base = `/v4/accounts/${validId(entry.accountId, RequestError)}`
  let asset
  try { asset = (await request(`${base}/${kind}/${assetId}`)).data }
  catch (error) {
    if (!(error instanceof RequestError) || error.status !== 404) throw error
    const project = (await request(`${base}/projects/${validId(entry.projectId, RequestError)}`)).data
    if (project?.id !== entry.projectId || typeof project.root_folder_id !== 'string' || !project.root_folder_id) throw new RequestError(403, 'The original project is unavailable. The existing mapping was preserved.')
    error.relinkAllowed = true
    throw error
  }
  if (asset?.id !== assetId || asset.project_id !== entry.projectId || asset.parent_id !== entry.folderId) throw new RequestError(409, 'The tracked asset moved or changed in Frame.io. Its mapping must be reviewed before uploading another version.')
  return { ...entry, name: typeof asset.name === 'string' ? asset.name : entry.name, viewUrl: viewUrl(asset.view_url, RequestError) }
}

async function finalizeUpload(record, registry, request, RequestError, now = Date.now) {
  const documentType = record.documentType === undefined ? 'pptx' : record.documentType
  if (!['pptx', 'docx'].includes(documentType)) throw new RequestError(409, 'Invalid tracked document type.')
  const priorEvent = await registry.event(record.presentationId, record.uploadId)
  if (priorEvent?.state === 'complete') {
    if ((priorEvent.documentType === undefined ? 'pptx' : priorEvent.documentType) !== documentType) throw new RequestError(409, 'The registered upload has a different document type.')
    if (await registry.record(record.presentationId, priorEvent) === false) throw new RequestError(409, 'This upload belongs to a replaced document mapping. Refresh history.')
    return priorEvent
  }
  const current = await registry.latest(record.presentationId)
  const previous = current || record.previous
  if (previous && (previous.documentType === undefined ? 'pptx' : previous.documentType) !== documentType) throw new RequestError(409, 'New versions must use the existing document type. No new asset was registered.')
  if (previous && (previous.accountId !== record.accountId || previous.projectId !== record.projectId || previous.folderId !== record.folderId)) throw new RequestError(409, 'This document is already linked to another destination. No new asset was registered.')
  const base = `/v4/accounts/${record.accountId}`
  const file = (await request(`${base}/files/${record.fileId}`)).data
  if (file?.id !== record.fileId || file.project_id !== record.projectId) throw new RequestError(502, 'Frame.io returned an inconsistent uploaded file.')
  let stackId
  let asset = file
  if (previous) {
    stackId = previous.stackId
    if (!stackId) {
      const original = (await request(`${base}/files/${validId(previous.fileId, RequestError)}`)).data
      if (original?.project_id !== record.projectId) throw new RequestError(409, 'The existing Frame.io file is no longer in this project.')
      if (original.parent_id !== record.folderId) stackId = validId(original.parent_id, RequestError)
    }
    if (stackId) {
      const stack = (await request(`${base}/version_stacks/${validId(stackId, RequestError)}`)).data
      if (stack?.id !== stackId || stack.project_id !== record.projectId || stack.parent_id !== record.folderId) throw new RequestError(409, 'The version stack is no longer in the selected destination.')
      if (file.parent_id !== stackId) {
        if (file.parent_id !== record.folderId) throw new RequestError(409, 'The uploaded file was moved before version registration.')
        await request(`${base}/files/${record.fileId}/move`, false, { method: 'PATCH', body: JSON.stringify({ data: { parent_id: stackId } }) })
      }
      asset = stack
    } else {
      if (file.parent_id !== record.folderId) throw new RequestError(409, 'The uploaded file was moved before version registration.')
      asset = (await request(`${base}/folders/${record.folderId}/version_stacks`, false, {
        method: 'POST', body: JSON.stringify({ data: { file_ids: [previous.fileId, record.fileId] } })
      })).data
      stackId = validId(asset?.id, RequestError)
      if (asset.project_id !== record.projectId || asset.parent_id !== record.folderId) throw new RequestError(502, 'Frame.io returned an inconsistent version stack.')
    }
  } else if (file.parent_id !== record.folderId) throw new RequestError(409, 'The uploaded file is no longer in the selected destination.')
  const entry = {
    presentationId: record.presentationId, documentType, uploadId: record.uploadId, accountId: record.accountId,
    projectId: record.projectId, folderId: record.folderId, name: record.name, fileSize: record.fileSize,
    fileId: record.fileId, ...(stackId ? { stackId } : {}), viewUrl: viewUrl(asset.view_url, RequestError),
    state: 'complete', version: (previous?.version || 0) + 1,
    createdAt: record.createdAt, completedAt: new Date(now()).toISOString()
  }
  if (await registry.record(record.presentationId, entry) === false) throw new RequestError(409, 'This upload belongs to a replaced document mapping. Refresh history.')
  return entry
}

module.exports = { createRegistry, verifyAsset, finalizeUpload, uuid }
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
    latest: presentationId => read(name(presentationId)),
    event: (presentationId, uploadId) => read(name(presentationId, uploadId)),
    async record(presentationId, entry) {
      await files.write(name(presentationId, entry.uploadId), JSON.stringify(entry))
      if (entry.state === 'complete') {
        const current = await read(name(presentationId))
        if (!current || entry.version > current.version || (entry.version === current.version && entry.completedAt >= current.completedAt)) await files.write(name(presentationId), JSON.stringify(entry))
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
  const asset = (await request(`/v4/accounts/${validId(entry.accountId, RequestError)}/${kind}/${assetId}`)).data
  if (asset?.id !== assetId || asset.project_id !== entry.projectId || asset.parent_id !== entry.folderId) throw new RequestError(409, 'The tracked asset moved or changed in Frame.io. Its mapping must be reviewed before uploading another version.')
  return { ...entry, name: typeof asset.name === 'string' ? asset.name : entry.name, viewUrl: viewUrl(asset.view_url, RequestError) }
}

async function finalizeUpload(record, registry, request, RequestError, now = Date.now) {
  const priorEvent = await registry.event(record.presentationId, record.uploadId)
  if (priorEvent?.state === 'complete') {
    await registry.record(record.presentationId, priorEvent)
    return priorEvent
  }
  const current = await registry.latest(record.presentationId)
  const previous = current || record.previous
  if (previous && (previous.accountId !== record.accountId || previous.projectId !== record.projectId || previous.folderId !== record.folderId)) throw new RequestError(409, 'This presentation is already linked to another destination. No new asset was registered.')
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
    presentationId: record.presentationId, uploadId: record.uploadId, accountId: record.accountId,
    projectId: record.projectId, folderId: record.folderId, name: record.name, fileSize: record.fileSize,
    fileId: record.fileId, ...(stackId ? { stackId } : {}), viewUrl: viewUrl(asset.view_url, RequestError),
    state: 'complete', version: (previous?.version || 0) + 1,
    createdAt: record.createdAt, completedAt: new Date(now()).toISOString()
  }
  await registry.record(record.presentationId, entry)
  return entry
}

module.exports = { createRegistry, verifyAsset, finalizeUpload, uuid }
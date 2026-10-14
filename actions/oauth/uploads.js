'use strict'

const { finalizeUpload, verifyAsset, uuid } = require('./registry')
const maxFileSize = 1024 * 1024 * 1024
const uploadTTL = 86400

async function uploadOperation({ params, identity, read, write, request, registry, RequestError, now = Date.now }) {
  const fail = (status, message) => { throw new RequestError(status, message) }
  const id = value => {
    if (typeof value !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(value)) fail(400, 'Invalid selection ID.')
    return value
  }
  if (typeof params.uploadId !== 'string' || !/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/.test(params.uploadId)) fail(400, 'Invalid upload ID.')
  const key = 'upload-' + identity + '-' + params.uploadId
  const existing = await read(key)
  if (params.operation === 'upload-status') {
    if (!existing?.fileId) fail(404, 'This upload is unavailable. Start a new upload.')
    const status = (await request(`/v4/accounts/${existing.accountId}/files/${existing.fileId}/status`)).data
    if (typeof status?.upload_complete !== 'boolean' || typeof status?.upload_failed !== 'boolean') fail(502, 'Frame.io returned an invalid upload status.')
    if (status.upload_failed && existing.presentationId) {
      const event = await registry.event(existing.presentationId, existing.uploadId)
      if (event && event.state !== 'complete') await registry.record(existing.presentationId, { ...event, state: 'failed', checkedAt: new Date(now()).toISOString() })
    }
    const entry = status.upload_complete && !status.upload_failed && existing.presentationId ? await finalizeUpload(existing, registry, request, RequestError, now) : undefined
    return { complete: status.upload_complete, failed: status.upload_failed, fileId: existing.fileId, ...(entry ? { entry } : {}) }
  }
  if (typeof params.presentationId !== 'string' || !uuid.test(params.presentationId)) fail(400, 'A document tracking ID is required. Reload the updated add-in.')
  const accountId = id(params.accountId)
  const projectId = id(params.projectId)
  const folderId = id(params.folderId)
  const { name, fileSize } = params
  const documentType = params.documentType === undefined ? 'pptx' : params.documentType
  if (!['pptx', 'docx'].includes(documentType)) fail(400, 'Unsupported document type.')
  if (typeof name !== 'string' || name.length <= 5 || name.length > 180 || !name.toLowerCase().endsWith('.' + documentType) || /[\x00-\x1f\x7f/\\]/.test(name)) fail(400, `Provide a valid ${documentType.toUpperCase()} filename of at most 180 characters.`)
  if (!Number.isSafeInteger(fileSize) || fileSize < 1 || fileSize > maxFileSize) fail(400, 'Documents must be between 1 byte and 1 GiB.')
  if (existing) {
    if ((existing.documentType === undefined ? 'pptx' : existing.documentType) !== documentType || existing.presentationId !== params.presentationId || existing.accountId !== accountId || existing.projectId !== projectId || existing.folderId !== folderId || existing.name !== name || existing.fileSize !== fileSize) fail(409, 'The upload ID is already in use.')
    if (!existing.parts) fail(409, 'The upload start outcome is uncertain. Do not retry with a new ID until you check the destination for a duplicate.')
    return { fileId: existing.fileId, parts: existing.parts }
  }
  const previous = await registry.latest(params.presentationId)
  if (previous) {
    if ((previous.documentType === undefined ? 'pptx' : previous.documentType) !== documentType) fail(409, 'New versions must use the existing document type.')
    await verifyAsset(previous, request, RequestError)
    if (previous.accountId !== accountId || previous.projectId !== projectId || previous.folderId !== folderId) fail(409, 'New versions must use the existing Frame.io destination.')
  }
  const project = (await request(`/v4/accounts/${accountId}/projects/${projectId}`)).data
  if (!project?.root_folder_id) fail(502, 'Frame.io returned an invalid project.')
  if (folderId !== project.root_folder_id) {
    const folder = (await request(`/v4/accounts/${accountId}/folders/${folderId}`)).data
    if (folder?.project_id !== projectId) fail(403, 'The folder does not belong to this project.')
  }
  const event = { presentationId: params.presentationId, documentType, uploadId: params.uploadId, accountId, projectId, folderId, name, fileSize, createdAt: new Date(now()).toISOString() }
  const record = { ...event, previous, expiresAt: now() + uploadTTL * 1000 }
  await write(key, record, uploadTTL)
  await registry.record(params.presentationId, { ...event, state: 'creating' })
  let file
  try {
    file = (await request(`/v4/accounts/${accountId}/folders/${folderId}/files/local_upload`, false, {
      method: 'POST', body: JSON.stringify({ data: { name, file_size: fileSize } })
    })).data
  } catch (error) {
    if (error instanceof RequestError && error.status >= 400 && error.status < 500) await registry.record(params.presentationId, { ...event, state: 'rejected', statusCode: error.status })
    throw error
  }
  const fileId = id(file?.id)
  await write(key, { ...record, fileId }, uploadTTL)
  await registry.record(params.presentationId, { ...event, fileId, state: 'uploading' })
  if (!Array.isArray(file.upload_urls) || !file.upload_urls.length || file.upload_urls.length > 1000) fail(502, 'Frame.io returned invalid upload parts.')
  const parts = file.upload_urls.map(part => {
    let url
    try { url = new URL(part.url) } catch { fail(502, 'Frame.io returned an invalid upload URL.') }
    if (url.protocol !== 'https:' || !url.hostname.endsWith('.amazonaws.com') || url.username || url.password || url.hash || (url.port && url.port !== '443')) fail(502, 'Frame.io returned an unsupported upload URL.')
    if (!Number.isSafeInteger(part.size) || part.size < 1 || part.size > 64 * 1024 * 1024) fail(502, 'Frame.io returned an unsupported upload part size.')
    return { url: url.href, size: part.size }
  })
  if (new Set(parts.map(part => part.url)).size !== parts.length || parts.reduce((total, part) => total + part.size, 0) !== fileSize || JSON.stringify(parts).length > 500000) fail(502, 'Frame.io returned inconsistent upload parts.')
  await write(key, { ...record, fileId, parts }, uploadTTL)
  return { fileId, parts }
}

module.exports = { uploadOperation, maxFileSize }
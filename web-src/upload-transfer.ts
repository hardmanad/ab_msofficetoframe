export type UploadPart = { url: string; size: number }
export type ByteSource = { size: number; read: (offset: number, size: number, signal: AbortSignal) => Promise<Uint8Array>; close: () => Promise<void> }
export const maxFileSize = 1024 * 1024 * 1024
export const pptxType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

export class TransferError extends Error {
  constructor(message: string, public status = 0) { super(message) }
}

export function delay(milliseconds: number, signal: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    signal.throwIfAborted()
    const abort = () => { clearTimeout(timer); reject(signal.reason) }
    const timer = setTimeout(() => { signal.removeEventListener('abort', abort); resolve() }, milliseconds)
    signal.addEventListener('abort', abort, { once: true })
  })
}

export function validateParts(parts: UploadPart[], size: number) {
  if (!Array.isArray(parts) || !parts.length || parts.length > 1000 || !Number.isSafeInteger(size) || size < 1 || size > maxFileSize) throw new Error('Invalid upload plan.')
  for (const part of parts) {
    const url = new URL(part.url)
    if (url.protocol !== 'https:' || !url.hostname.endsWith('.amazonaws.com') || url.username || url.password || url.hash || (url.port && url.port !== '443') || !Number.isSafeInteger(part.size) || part.size < 1 || part.size > 64 * 1024 * 1024) throw new Error('Unsupported upload plan.')
  }
  if (new Set(parts.map(part => part.url)).size !== parts.length) throw new Error('The upload plan contains duplicate part URLs.')
  if (parts.reduce((total, part) => total + part.size, 0) !== size) throw new Error('The upload plan does not match the presentation size.')
}

export function putPart(url: string, bytes: Uint8Array, progress: (loaded: number) => void, signal: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    signal.throwIfAborted()
    const xhr = new XMLHttpRequest()
    const finish = (error?: Error) => {
      signal.removeEventListener('abort', abort)
      if (error) reject(error)
      else resolve()
    }
    const abort = () => xhr.abort()
    xhr.open('PUT', url)
    xhr.timeout = 600000
    xhr.withCredentials = false
    xhr.setRequestHeader('Content-Type', pptxType)
    xhr.setRequestHeader('x-amz-acl', 'private')
    xhr.upload.onprogress = event => progress(Math.min(event.loaded, bytes.length))
    xhr.onload = () => finish(xhr.status >= 200 && xhr.status < 300 ? undefined : new TransferError(xhr.status === 403 ? 'Upload permission or URL expiry prevented transfer. Check Frame.io before starting again.' : 'A file part could not be transferred.', xhr.status))
    xhr.onerror = () => finish(new TransferError('Transfer failed. Check your connection and storage CORS configuration.'))
    xhr.ontimeout = () => finish(new TransferError('A file part timed out.'))
    xhr.onabort = () => finish(new TransferError('Upload cancelled.'))
    signal.addEventListener('abort', abort, { once: true })
    try { xhr.send(bytes as Uint8Array<ArrayBuffer>) } catch { finish(new TransferError('Transfer could not start.')) }
  })
}

export async function transferParts(source: ByteSource, parts: UploadPart[], progress: (loaded: number) => void, signal: AbortSignal, put = putPart, wait = delay) {
  validateParts(parts, source.size)
  const controller = new AbortController()
  const combined = AbortSignal.any([signal, controller.signal])
  const offsets: number[] = []
  let offset = 0
  for (const part of parts) { offsets.push(offset); offset += part.size }
  const loaded = parts.map(() => 0)
  const report = () => progress(loaded.reduce((total, value) => total + value, 0))
  let next = 0
  let failure: unknown
  async function worker() {
    try {
      while (next < parts.length) {
        combined.throwIfAborted()
        const index = next++
        const part = parts[index]
        const bytes = await source.read(offsets[index], part.size, combined)
        if (bytes.length !== part.size) throw new Error('PowerPoint returned an incomplete file part.')
        for (let attempt = 0; ; attempt++) {
          combined.throwIfAborted()
          try {
            await put(part.url, bytes, value => { loaded[index] = Math.max(0, Math.min(value, part.size)); report() }, combined)
            loaded[index] = part.size
            report()
            break
          } catch (error) {
            combined.throwIfAborted()
            if (!(error instanceof TransferError) || (error.status !== 0 && error.status !== 408 && error.status !== 429 && error.status < 500) || attempt >= 3) throw error
            loaded[index] = 0
            report()
            await wait(Math.min(1000 * 2 ** attempt + Math.random() * 500, 10000), combined)
          }
        }
      }
    } catch (error) {
      if (!failure) failure = error
      controller.abort()
    }
  }
  await Promise.all(Array.from({ length: Math.min(2, parts.length) }, worker))
  if (failure) throw failure
}

export async function openPresentation(signal: AbortSignal): Promise<ByteSource> {
  const file = await officeCall<Office.File>(callback => Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 1024 * 1024 }, callback), signal, result => result.closeAsync())
  let closed = false
  let queue: Promise<unknown> = Promise.resolve()
  return {
    size: file.size,
    read(offset, size, readSignal) {
      const task = queue.then(async () => {
        readSignal.throwIfAborted()
        if (closed || offset < 0 || size < 1 || offset + size > file.size) throw new Error('Invalid presentation read.')
        const bytes = new Uint8Array(size)
        let copied = 0
        while (copied < size) {
          const position = offset + copied
          const index = Math.floor(position / (1024 * 1024))
          const slice = await officeCall<Office.Slice>(callback => file.getSliceAsync(index, callback), readSignal)
          const data = new Uint8Array(slice.data)
          const start = position % (1024 * 1024)
          const count = Math.min(data.length - start, size - copied)
          if (slice.index !== index || count <= 0 || data.length !== Math.min(1024 * 1024, file.size - index * 1024 * 1024)) throw new Error('PowerPoint returned an invalid presentation slice.')
          bytes.set(data.subarray(start, start + count), copied)
          copied += count
        }
        return bytes
      })
      queue = task.catch(() => {})
      return task
    },
    async close() {
      if (closed) return
      closed = true
      await officeCall<void>(callback => file.closeAsync(callback), new AbortController().signal)
    }
  }
}

function officeCall<Value>(invoke: (callback: (result: Office.AsyncResult<Value>) => void) => void, signal: AbortSignal, lateCleanup?: (value: Value) => void): Promise<Value> {
  return new Promise((resolve, reject) => {
    signal.throwIfAborted()
    let settled = false
    const finish = (error?: Error, value?: Value) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      signal.removeEventListener('abort', abort)
      if (error) reject(error)
      else resolve(value as Value)
    }
    const abort = () => finish(new Error('Upload cancelled.'))
    const timer = setTimeout(() => finish(new Error('PowerPoint file access timed out.')), 30000)
    signal.addEventListener('abort', abort, { once: true })
    try {
      invoke(result => {
        if (settled) { if (result.status === Office.AsyncResultStatus.Succeeded) lateCleanup?.(result.value); return }
        if (result.status !== Office.AsyncResultStatus.Succeeded) finish(new Error('PowerPoint could not access the presentation.'))
        else finish(undefined, result.value)
      })
    } catch { finish(new Error('PowerPoint file access is unavailable.')) }
  })
}
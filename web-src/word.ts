export const wordTrackingProperty = 'MSOFFICE_TO_FRAME_UUID'
const uuid = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i
let identityPending: Promise<string> | undefined
let fileUrl: string | undefined

export function wordFileUrl(): URL | null {
  try {
    const source = fileUrl || Office.context.document.url
    const path = /^[a-z]:[\\/]/i.test(source) ? 'file:///' + source.replace(/\\/g, '/') : source
    return new URL(path, source.startsWith('/') ? 'file:///' : undefined)
  } catch { return null }
}

export function wordFilename(): string {
  try {
    const name = decodeURIComponent(wordFileUrl()?.pathname.split('/').pop() || '')
    if (name) return name
  } catch {}
  return 'Untitled document'
}

export async function loadWordFileProperties(): Promise<void> {
  if (typeof Office.context.document.getFilePropertiesAsync !== 'function') return
  await new Promise<void>(resolve => {
    let active = true
    const finish = () => { active = false; clearTimeout(timeout); resolve() }
    const timeout = setTimeout(finish, 2000)
    try {
      Office.context.document.getFilePropertiesAsync(result => {
        if (!active) return
        if (result.status === Office.AsyncResultStatus.Succeeded && typeof result.value?.url === 'string' && result.value.url) fileUrl = result.value.url
        finish()
      })
    } catch { finish() }
  })
}

export async function wordDocumentSaved(): Promise<boolean> {
  if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) return false
  return Word.run(async context => {
    context.document.load('saved')
    await context.sync()
    return context.document.saved === true
  })
}

export async function ensureWordDocumentId(generate = () => crypto.randomUUID()): Promise<string> {
  if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) throw new Error('Document identity requires WordApi 1.3.')
  if (identityPending) return identityPending
  identityPending = Word.run(async context => {
    const property = context.document.properties.customProperties.getItemOrNullObject(wordTrackingProperty)
    property.load('value')
    await context.sync()
    if (!property.isNullObject) {
      if (typeof property.value !== 'string' || !uuid.test(property.value)) throw new Error('The document contains an invalid tracking ID.')
      return property.value.toLowerCase()
    }
    const value = generate()
    if (!uuid.test(value)) throw new Error('A document tracking ID could not be generated.')
    context.document.properties.customProperties.add(wordTrackingProperty, value)
    await context.sync()
    return value.toLowerCase()
  })
  try { return await identityPending } finally { identityPending = undefined }
}
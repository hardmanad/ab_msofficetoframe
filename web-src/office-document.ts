import { ensurePresentationId, firstSlidePreview, presentationFilename } from './presentation'
import { ensureWordDocumentId, loadWordFileProperties, wordDocumentSaved, wordFilename, wordFileUrl } from './word'

export type DocumentType = 'pptx' | 'docx'
export type OfficeDocument = {
  kind: 'powerpoint' | 'word'; documentType: DocumentType; mimeType: string
  noun: string; defaultName: string; filename: () => string
  ensureId: (generate?: () => ReturnType<Crypto['randomUUID']>) => Promise<string>; preview?: () => Promise<string | null>
  isSaved?: () => Promise<boolean>
  loadMetadata?: () => Promise<void>
}
export type OfficeHost = {
  kind: 'loading' | 'browser' | 'powerpoint' | 'word' | 'other'
  adapter: OfficeDocument | null; available: boolean; reason: string; inOffice: boolean
}

const powerpoint: OfficeDocument = {
  kind: 'powerpoint', documentType: 'pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  noun: 'presentation', defaultName: 'Presentation.pptx', filename: presentationFilename,
  ensureId: ensurePresentationId, preview: firstSlidePreview
}
const word: OfficeDocument = {
  kind: 'word', documentType: 'docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  noun: 'document', defaultName: 'Document.docx', filename: wordFilename, ensureId: ensureWordDocumentId, isSaved: wordDocumentSaved,
  loadMetadata: loadWordFileProperties
}

export function validUploadName(name: string, documentType: DocumentType): boolean {
  return name.length > 5 && name.length <= 180 && name.toLowerCase().endsWith('.' + documentType) && !/[\x00-\x1f\x7f/\\]/.test(name)
}

export function initialUploadName(adapter: OfficeDocument): string {
  const name = adapter.filename()
  return validUploadName(name, adapter.documentType) ? name : adapter.defaultName
}

export function assertSourceFormat(adapter: OfficeDocument) {
  let name = ''
  try {
    const url = adapter.kind === 'word' ? wordFileUrl() : new URL(Office.context.document.url)
    name = decodeURIComponent(url?.pathname.split('/').pop() || '')
  } catch {}
  if (name && !name.toLowerCase().endsWith('.' + adapter.documentType)) throw new Error(`Open a .${adapter.documentType} ${adapter.noun}. Legacy, macro-enabled and template files are not supported.`)
}

export async function verifyDocumentIdentity(adapter: OfficeDocument, expected: string) {
  assertSourceFormat(adapter)
  const current = await adapter.ensureId(() => { throw new Error('The document tracking ID was removed. Reload the add-in before uploading.') })
  if (current !== expected) throw new Error('The document tracking ID changed. Reload the add-in before uploading.')
}

export function getOfficeHost(info: { host: Office.HostType | null; platform: Office.PlatformType | null }): OfficeHost {
  if (!info.host) return { kind: 'browser', adapter: null, available: false, reason: 'Open this add-in in Word or PowerPoint.', inOffice: false }
  const adapter = info.host === Office.HostType.PowerPoint ? powerpoint : info.host === Office.HostType.Word ? word : null
  if (!adapter) return { kind: info.host ? 'other' : 'browser', adapter: null, available: false, reason: 'Open this add-in in Word or PowerPoint.', inOffice: false }
  let reason = ''
  if (adapter.kind === 'word' && info.platform !== Office.PlatformType.Mac && info.platform !== Office.PlatformType.PC) reason = 'DOCX uploads require Word on Mac or Windows desktop.'
  else if (!Office.context.requirements.isSetSupported('CompressedFile', '1.1')) reason = `${adapter.kind === 'word' ? 'Word DOCX' : 'PowerPoint PPTX'} file export is unavailable in this Office version.`
  else if (!Office.context.requirements.isSetSupported(adapter.kind === 'word' ? 'WordApi' : 'PowerPointApi', '1.3')) reason = `${adapter.kind === 'word' ? 'WordApi' : 'PowerPointApi'} 1.3 is required for document tracking.`
  if (!reason) {
    try { assertSourceFormat(adapter) } catch (error) { reason = error instanceof Error ? error.message : 'Unsupported document format.' }
  }
  return { kind: adapter.kind, adapter, available: !reason, reason, inOffice: true }
}
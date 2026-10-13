export const presentationTag = 'MSOFFICE_TO_FRAME_UUID'
const uuid = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i

export function presentationFilename(): string {
  try {
    const location = Office.context.document.url
    const name = decodeURIComponent(new URL(location).pathname.split('/').pop() || '')
    if (name) return name
  } catch {}
  return 'Untitled presentation'
}

export async function firstSlidePreview(): Promise<string | null> {
  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.8')) return null
  return PowerPoint.run(async context => {
    const count = context.presentation.slides.getCount()
    await context.sync()
    if (!count.value) return null
    const image = context.presentation.slides.getItemAt(0).getImageAsBase64({ width: 480 })
    await context.sync()
    if (typeof image.value !== 'string' || !image.value.startsWith('iVBORw0KGgo') || !/^[A-Za-z0-9+/=]+$/.test(image.value)) throw new Error('PowerPoint returned an invalid slide preview.')
    return 'data:image/png;base64,' + image.value
  })
}

export async function ensurePresentationId(generate = () => crypto.randomUUID()): Promise<string> {
  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.3')) throw new Error('Presentation identity requires PowerPointApi 1.3.')
  return PowerPoint.run(async context => {
    const tag = context.presentation.tags.getItemOrNullObject(presentationTag)
    tag.load('value')
    await context.sync()
    if (!tag.isNullObject) {
      if (!uuid.test(tag.value)) throw new Error('The presentation contains an invalid tracking ID.')
      return tag.value.toLowerCase()
    }
    const value = generate()
    if (!uuid.test(value)) throw new Error('A presentation tracking ID could not be generated.')
    context.presentation.tags.add(presentationTag, value)
    await context.sync()
    return value.toLowerCase()
  })
}
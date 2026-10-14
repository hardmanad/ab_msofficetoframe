'use strict'

const fs = require('node:fs')
const path = require('node:path')
const { parseArgs } = require('node:util')
const { XMLBuilder } = require('fast-xml-parser')

function createManifest(taskpaneUrl) {
  const url = new URL(taskpaneUrl)
  if (url.protocol !== 'https:' || url.username || url.password || (url.search && !/^\?v=[A-Za-z0-9_-]+$/.test(url.search)) || url.hash ||
      /(^|\.)(example\.(com|org|net)|invalid|test)$/.test(url.hostname) || /[<>]|YOUR_|PLACEHOLDER/i.test(taskpaneUrl)) {
    throw new Error('Use an actual HTTPS task-pane URL without credentials or fragments; only a ?v= cache version query is allowed.')
  }
  const icon = size => new URL(`./assets/icon-${size}.png?v=frameio`, url).href
  const manifest = {
    '?xml': { '@_version': '1.0', '@_encoding': 'UTF-8' },
    OfficeApp: {
      '@_xmlns': 'http://schemas.microsoft.com/office/appforoffice/1.1',
      '@_xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
      '@_xsi:type': 'TaskPaneApp',
      Id: '4f391fa8-6ea8-4a32-91d2-3b298c24104e',
      Version: '1.0.0.0',
      ProviderName: 'MSOffice to Frame',
      DefaultLocale: 'en-US',
      DisplayName: { '@_DefaultValue': 'Frame.io' },
      Description: { '@_DefaultValue': 'Upload PowerPoint presentations and Word documents to Frame.io with persistent document tracking.' },
      IconUrl: { '@_DefaultValue': icon(32) },
      HighResolutionIconUrl: { '@_DefaultValue': icon(64) },
      SupportUrl: { '@_DefaultValue': url.href },
      AppDomains: { AppDomain: url.origin },
      Hosts: { Host: [{ '@_Name': 'Presentation' }, { '@_Name': 'Document' }] },
      DefaultSettings: { SourceLocation: { '@_DefaultValue': url.href } },
      Permissions: 'ReadWriteDocument'
    }
  }
  return new XMLBuilder({ ignoreAttributes: false, format: true }).build(manifest)
}

function main() {
  const { values } = parseArgs({ options: { url: { type: 'string' }, output: { type: 'string' } } })
  if (!values.url || !values.output) throw new Error('Required: --url <actual HTTPS task-pane URL> --output office/manifest.local.xml or office/manifest.deployed.xml')
  const output = path.resolve(values.output)
  const office = path.resolve(__dirname, '../office')
  if (path.dirname(output) !== office || !/^manifest\.(local|deployed)\.xml$/.test(path.basename(output))) {
    throw new Error('Manifest output must be office/manifest.local.xml or office/manifest.deployed.xml.')
  }
  if (output.endsWith('manifest.deployed.xml') && ['localhost', '127.0.0.1', 'www.localappbuilder.com'].includes(new URL(values.url).hostname)) {
    throw new Error('A deployed manifest cannot use localhost.')
  }
  const xml = createManifest(values.url)
  fs.mkdirSync(office, { recursive: true })
  fs.writeFileSync(output, xml)
  console.log(`Generated ${path.relative(process.cwd(), output)}. Validate and verify all URLs before sideloading.`)
}

if (require.main === module) {
  try { main() } catch (error) { console.error(error.message); process.exitCode = 1 }
}
module.exports = { createManifest }
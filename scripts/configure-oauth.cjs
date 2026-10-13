'use strict'

const fs = require('node:fs')
const crypto = require('node:crypto')
const dotenv = require('dotenv')

function configure() {
  const exported = JSON.parse(fs.readFileSync('msofficetoframe-27200-Stage.json', 'utf8'))
  const credentials = exported.project.workspace.details.credentials.filter(credential => credential.oauth2)
  if (credentials.length !== 1) throw new Error('Expected exactly one OAuth Web App credential in the Stage export.')
  const credential = credentials[0].oauth2
  if (!credential.client_id || !credential.client_secret || credential.defaultRedirectUri !== 'https://www.localappbuilder.com:3000/oauth-callback.html') {
    throw new Error('Web App credential or registered custom-domain redirect is missing.')
  }
  const current = fs.existsSync('.env') ? fs.readFileSync('.env', 'utf8') : ''
  const env = dotenv.parse(current)
  const values = {
    FRAMEIO_CLIENT_ID: credential.client_id,
    FRAMEIO_CLIENT_SECRET: credential.client_secret,
    FRAMEIO_REDIRECT_URI: env.FRAMEIO_REDIRECT_URI || credential.defaultRedirectUri,
    FRAMEIO_SCOPES: !env.FRAMEIO_SCOPES || env.FRAMEIO_SCOPES === 'frameio.read,frameio.write,offline_access'
      ? 'openid email profile offline_access additional_info.roles' : env.FRAMEIO_SCOPES,
    FRAMEIO_TOKEN_KEY: env.FRAMEIO_TOKEN_KEY || crypto.randomBytes(32).toString('hex')
  }
  if (!/^[a-f0-9]{64}$/.test(values.FRAMEIO_TOKEN_KEY)) throw new Error('FRAMEIO_TOKEN_KEY must be a 32-byte hex key; existing value was preserved.')
  const lines = current.split(/\r?\n/).filter(line => !Object.keys(values).some(key => line.startsWith(key + '=')))
  while (lines.length && lines[lines.length - 1] === '') lines.pop()
  for (const [key, value] of Object.entries(values)) lines.push(`${key}=${JSON.stringify(value)}`)
  fs.writeFileSync('.env', lines.join('\n') + '\n', { mode: 0o600 })
  fs.chmodSync('.env', 0o600)
  console.log('Configured server-only FRAMEIO_* variables from the Web App credential. No values displayed; unrelated variables preserved.')
}

try { configure() } catch { console.error('OAuth configuration failed. Check the private export and existing FRAMEIO_* settings; no values displayed.'); process.exitCode = 1 }
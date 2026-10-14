'use strict'

const fs = require('node:fs')
const path = require('node:path')

function hasSecret(content, secrets) {
  const buffer = Buffer.isBuffer(content) ? content : Buffer.from(content)
  return secrets.some(secret => secret.length >= 8 && [secret, encodeURIComponent(secret), Buffer.from(secret).toString('base64')].some(value => buffer.includes(Buffer.from(value))))
}

function collectSecrets() {
  const values = new Set()
  function add(value) {
    if (typeof value === 'string' && value.length >= 8) {
      values.add(value)
      for (const part of value.split(':')) if (part.length >= 8) values.add(part)
    }
  }
  function visit(object, sensitive = false) {
    if (!object || typeof object !== 'object') return
    for (const [key, value] of Object.entries(object)) {
      if (key === 'redirect_uri' || key === 'defaultRedirectUri') continue
      const privateValue = sensitive || /auth|secret|token|password|credential|api.?key/i.test(key)
      if (typeof value === 'object') visit(value, privateValue)
      else if (privateValue) add(value)
    }
  }
  for (const file of ['.aio', 'msofficetoframe-27200-Stage.json']) {
    if (fs.existsSync(file)) visit(JSON.parse(fs.readFileSync(file, 'utf8')))
  }
  if (fs.existsSync('.env')) {
    for (const line of fs.readFileSync('.env', 'utf8').split(/\r?\n/)) {
      const match = line.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/)
      if (match && /auth|secret|token|password|key/i.test(match[1])) add(match[2].replace(/^['"]|['"]$/g, ''))
    }
  }
  return [...values]
}

function main() {
  const directory = path.resolve('dist/application/web-prod')
  if (!fs.existsSync(path.join(directory, 'index.html'))) throw new Error('Build web assets before checking secrets.')
  for (const size of [32, 64, 80, 128]) {
    const icon = path.join(directory, 'assets', `icon-${size}.png`)
    if (!fs.existsSync(icon) || fs.statSync(icon).size === 0) throw new Error('Manifest icons are missing from the web build.')
  }
  const secrets = collectSecrets()
  let count = 0
  function scan(folder) {
    for (const entry of fs.readdirSync(folder, { withFileTypes: true })) {
      const file = path.join(folder, entry.name)
      if (entry.isDirectory()) scan(file)
      else {
        count++
        if (hasSecret(fs.readFileSync(file), secrets)) throw new Error(`Secret detected in ${path.relative(process.cwd(), file)}. Do not deploy.`)
      }
    }
  }
  scan(directory)
  for (const entry of fs.readdirSync('web-src')) {
    if (/\.(tsx?|jsx?)$/.test(entry) && /process\s*\.\s*env|import\s*\.\s*meta\s*\.\s*env/.test(fs.readFileSync(path.join('web-src', entry), 'utf8'))) {
      throw new Error('Frontend environment access is prohibited.')
    }
  }
  console.log(`Secret scan passed for ${count} generated web files. Known local secret values were not found.`)
}

if (require.main === module) {
  try { main() } catch (error) { console.error(error.message); process.exitCode = 1 }
}
module.exports = { hasSecret, collectSecrets }
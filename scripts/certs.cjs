'use strict'

const fs = require('node:fs')
const path = require('node:path')
const { spawnSync } = require('node:child_process')

const directory = path.resolve('.certs')
const key = path.join(directory, 'localappbuilder.key')
const cert = path.join(directory, 'localappbuilder.crt')
if (fs.existsSync(key) || fs.existsSync(cert)) {
  console.error('Custom-domain certificate files already exist. Preserve them or remove them explicitly before regeneration.')
  process.exit(1)
}
fs.mkdirSync(directory, { recursive: true })
const result = spawnSync('openssl', ['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-keyout', key, '-out', cert,
  '-days', '30', '-subj', '/CN=www.localappbuilder.com', '-addext', 'subjectAltName=DNS:www.localappbuilder.com,DNS:localhost,IP:127.0.0.1',
  '-addext', 'basicConstraints=critical,CA:FALSE', '-addext', 'extendedKeyUsage=serverAuth'], { stdio: 'ignore' })
if (result.error || result.status !== 0) {
  console.error('Certificate generation failed. Install OpenSSL or provide your own hostname-valid cert/key.')
  process.exit(1)
}
fs.chmodSync(key, 0o600)
console.log('Generated .certs/localappbuilder.crt and its private key. No host trust was installed.')
console.log('Trust ONLY the public certificate on the Mac, or replace these files with a certificate from your trusted development CA.')
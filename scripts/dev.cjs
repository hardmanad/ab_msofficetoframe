'use strict'

const fs = require('node:fs')
const path = require('node:path')
const { X509Certificate } = require('node:crypto')
const https = require('node:https')
const { spawn } = require('node:child_process')
const { parseArgs } = require('node:util')
const serve = require('serve-handler')
const { main: status } = require('../actions/status')
const { createHandler } = require('../actions/oauth')
const { localState } = require('./local-state.cjs')
const { localRegistry } = require('./local-registry.cjs')
const dotenv = require('dotenv')

const privateConfig = fs.existsSync('.env') ? dotenv.parse(fs.readFileSync('.env')) : {}
const oauthConfig = Object.fromEntries(Object.entries(privateConfig).filter(([key]) => key.startsWith('FRAMEIO_')))
const oauth = createHandler({ getStore: async () => localState(), getRegistry: async () => localRegistry() })

const { values } = parseArgs({ options: { cert: { type: 'string' }, key: { type: 'string' }, port: { type: 'string', default: '3000' } } })
const port = Number(values.port)
if (port !== 3000) {
  console.error('Local development requires port 3000 on both the container and host.')
  process.exit(1)
}
const hostname = 'localhost'
const cert = values.cert || path.resolve('.certs/localappbuilder.crt')
const key = values.key || path.resolve('.certs/localappbuilder.key')
if (!fs.existsSync(cert) || !fs.existsSync(key)) {
  console.error('HTTPS certificate missing. Run npm run certs, or pass --cert and --key paths. Trust the public certificate on the Office host.')
  process.exit(1)
}
if (!new X509Certificate(fs.readFileSync(cert)).checkHost(hostname)) {
  console.error('HTTPS certificate does not cover the configured OAuth hostname.')
  process.exit(1)
}

let watcher
const server = https.createServer({ cert: fs.readFileSync(cert), key: fs.readFileSync(key) }, async (request, response) => {
  try {
    if (new URL(request.url, 'https://localhost').pathname === '/api/status') {
      const result = await status({ __ow_method: request.method.toLowerCase() })
      response.writeHead(result.statusCode, result.headers)
      response.end(JSON.stringify(result.body))
    } else if (new URL(request.url, 'https://localhost').pathname === '/api/oauth') {
      let body = {}
      if (request.method === 'POST') {
        if (request.headers['content-type']?.split(';')[0] !== 'application/json') {
          response.writeHead(415, { 'Content-Type': 'application/json' })
          response.end(JSON.stringify({ error: 'JSON request required.' }))
          return
        }
        let length = 0
        const chunks = []
        for await (const chunk of request) {
          length += chunk.length
          if (length > 8192) {
            response.writeHead(413, { 'Content-Type': 'application/json' })
            response.end(JSON.stringify({ error: 'Request too large.' }))
            return
          }
          chunks.push(chunk)
        }
        try { body = JSON.parse(Buffer.concat(chunks).toString('utf8')) } catch {
          response.writeHead(400, { 'Content-Type': 'application/json' })
          response.end(JSON.stringify({ error: 'Invalid JSON request.' }))
          return
        }
        if (!body || Array.isArray(body) || typeof body !== 'object' || Object.keys(body).some(key => !['operation', 'state', 'code', 'error', 'accountId', 'projectId', 'folderId', 'cursor', 'uploadId', 'name', 'fileSize', 'presentationId', 'documentType'].includes(key))) {
          response.writeHead(400, { 'Content-Type': 'application/json' })
          response.end(JSON.stringify({ error: 'Invalid request fields.' }))
          return
        }
      }
      const result = await oauth({ ...body, ...oauthConfig, __ow_method: request.method.toLowerCase(), __ow_headers: request.headers })
      response.writeHead(result.statusCode, result.headers)
      response.end(result.statusCode === 204 ? '' : JSON.stringify(result.body))
    } else if (/^\/assets\/icon-(32|64|80|128)\.png$/.test(new URL(request.url, 'https://localhost').pathname)) {
      await serve(request, response, { public: 'web-src', cleanUrls: false, directoryListing: false })
    } else {
      await serve(request, response, { public: 'dist/dev', cleanUrls: false, directoryListing: false })
    }
  } catch {
    response.writeHead(500, { 'Content-Type': 'application/json' })
    response.end(JSON.stringify({ error: 'Local request failed.' }))
  }
})

server.on('error', error => {
  console.error(`HTTPS server failed (${error.code || 'unknown error'}). Ensure port 3000 is available; local development must use localhost:3000.`)
  if (watcher) watcher.kill()
  process.exitCode = 1
})
server.listen(port, '0.0.0.0', () => {
  console.log(`Task pane: https://${hostname}:${port}/index.html (local actions; no deployment)`)
  watcher = spawn(process.execPath, [path.resolve('node_modules/parcel/lib/bin.js'), 'watch', 'web-src/index.html', 'web-src/oauth-start.html', 'web-src/oauth-callback.html', '--dist-dir', 'dist/dev', '--public-url', './', '--no-hmr'], { stdio: 'inherit' })
  watcher.on('exit', code => { server.close(); if (code) process.exitCode = code })
})
for (const signal of ['SIGINT', 'SIGTERM']) {
  process.on(signal, () => { if (watcher) watcher.kill(); server.close() })
}
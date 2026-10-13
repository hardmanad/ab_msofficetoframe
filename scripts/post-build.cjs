'use strict'

const fs = require('node:fs')
const path = require('node:path')

const source = path.resolve(__dirname, '../web-src/assets')
const output = path.resolve(__dirname, '../dist/application/web-prod')
if (fs.existsSync(path.join(output, 'index.html'))) {
  fs.mkdirSync(path.join(output, 'assets'), { recursive: true })
  for (const size of [32, 64, 80, 128]) {
    const name = `icon-${size}.png`
    fs.writeFileSync(path.join(output, 'assets', name), fs.readFileSync(path.join(source, name)))
  }
  console.log('Copied stable manifest icons to web assets.')
}
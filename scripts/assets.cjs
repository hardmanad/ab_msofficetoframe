'use strict'

const fs = require('node:fs/promises')
const path = require('node:path')
const sharp = require('sharp')

async function main() {
  const directory = path.resolve(__dirname, '../web-src/assets')
  await fs.mkdir(directory, { recursive: true })
  const artwork = await fs.readFile(path.join(directory, 'frameio.png'))
  for (const size of [32, 64, 80, 128]) {
    await sharp(artwork).resize(size, size).png().toFile(path.join(directory, `icon-${size}.png`))
  }
  console.log('Generated task-pane PNG icons (32, 64, 80, 128).')
}

main().catch(() => { console.error('Icon generation failed.'); process.exitCode = 1 })
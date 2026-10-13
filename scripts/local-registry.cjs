'use strict'

const fs = require('node:fs/promises')
const path = require('node:path')
const crypto = require('node:crypto')
const { createRegistry } = require('../actions/oauth/registry')

function localRegistry(directory = path.resolve('.local-state/registry')) {
  function filename(name) {
    if (!/^presentation-uploads\/[a-f0-9-]{36}\/(current|[a-f0-9-]{36})\.json$/.test(name)) throw new Error('Invalid registry path.')
    return path.join(directory, name)
  }
  return createRegistry({
    async list(name) {
      try { await fs.stat(filename(name)); return [{ name }] }
      catch (error) { if (error.code === 'ENOENT') return []; throw error }
    },
    read: name => fs.readFile(filename(name)),
    async write(name, value) {
      const output = filename(name)
      await fs.mkdir(path.dirname(output), { recursive: true, mode: 0o700 })
      const temporary = output + '.tmp-' + crypto.randomBytes(8).toString('hex')
      await fs.writeFile(temporary, value, { mode: 0o600 })
      await fs.rename(temporary, output)
    }
  })
}

module.exports = { localRegistry }
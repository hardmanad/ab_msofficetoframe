'use strict'

const fs = require('node:fs/promises')
const path = require('node:path')

function localState(directory = path.resolve('.local-state')) {
  function filename(key) {
    if (!/^(login|session)-[A-Za-z0-9_-]{43,64}$/.test(key) && !/^upload-[a-f0-9]{64}-[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/.test(key)) throw new Error('Invalid storage key.')
    return path.join(directory, key + '.json')
  }
  return {
    async get(key) {
      try {
        const record = JSON.parse(await fs.readFile(filename(key), 'utf8'))
        if (record.expiration <= Date.now()) { await this.delete(key); return undefined }
        return record
      } catch (error) { if (error.code === 'ENOENT') return undefined; throw error }
    },
    async put(key, value, { ttl }) {
      await fs.mkdir(directory, { recursive: true, mode: 0o700 })
      const output = filename(key)
      const temporary = output + '.tmp-' + require('node:crypto').randomBytes(8).toString('hex')
      await fs.writeFile(temporary, JSON.stringify({ value, expiration: Date.now() + ttl * 1000 }), { mode: 0o600 })
      await fs.rename(temporary, output)
    },
    async delete(key) {
      try { await fs.unlink(filename(key)) } catch (error) { if (error.code !== 'ENOENT') throw error }
    }
  }
}

module.exports = { localState }
'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')
const { main } = require('../actions/status')

test('status is fixed and does not echo caller data or credentials', async () => {
  const result = await main({ __ow_method: 'get', token: 'do-not-echo', url: 'https://example.invalid' })
  assert.equal(result.statusCode, 200)
  assert.deepEqual(result.body, {
    status: 'ok',
    service: 'msofficetoframe',
    frameio: { configured: false, uploadEnabled: false }
  })
  assert.equal(JSON.stringify(result).includes('do-not-echo'), false)
  assert.equal(result.headers['Cache-Control'], 'no-store')
})

test('only GET is accepted over HTTP', async () => {
  for (const method of ['post', 'put', 'delete', 'options', 7]) {
    const result = await main({ __ow_method: method })
    assert.equal(result.statusCode, 405)
    assert.equal(result.headers.Allow, 'GET')
  }
})

test('direct Runtime invocation works and malformed input fails', async () => {
  assert.equal((await main()).statusCode, 200)
  for (const params of [null, 'bad', []]) {
    assert.equal((await main(params)).statusCode, 400)
  }
})
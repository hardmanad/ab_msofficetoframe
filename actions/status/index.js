'use strict'

function response(statusCode, body, extraHeaders = {}) {
  return {
    statusCode,
    headers: {
      'Content-Type': 'application/json',
      'Cache-Control': 'no-store',
      ...extraHeaders
    },
    body
  }
}

async function main(params = {}) {
  if (!params || typeof params !== 'object' || Array.isArray(params)) {
    return response(400, { error: 'Invalid request.' })
  }

  const method = params.__ow_method
  if (method !== undefined && (typeof method !== 'string' || method.toLowerCase() !== 'get')) {
    return response(405, { error: 'Only GET is supported.' }, { Allow: 'GET' })
  }

  return response(200, {
    status: 'ok',
    service: 'msofficetoframe',
    frameio: { configured: false, uploadEnabled: false }
  })
}

exports.main = main
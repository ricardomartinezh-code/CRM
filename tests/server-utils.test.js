const test = require('node:test');
const assert = require('node:assert/strict');

const {
  buildForwardUrl,
  pickForwardHeaders,
  selectForwardBody,
  DEFAULT_APPS_SCRIPT_URL,
  DEFAULT_ACTION
} = require('../server');

test('buildForwardUrl agrega action por defecto cuando falta', () => {
  const result = buildForwardUrl(DEFAULT_APPS_SCRIPT_URL, '/webhook?hub.mode=subscribe&hub.challenge=abc', DEFAULT_ACTION);
  assert.equal(result.searchParams.get('hub.mode'), 'subscribe');
  assert.equal(result.searchParams.get('hub.challenge'), 'abc');
  assert.equal(result.searchParams.get('action'), DEFAULT_ACTION);
});

test('buildForwardUrl respeta action existente', () => {
  const source = '/webhook?action=custom&foo=bar';
  const result = buildForwardUrl(DEFAULT_APPS_SCRIPT_URL + '?lang=es', source, DEFAULT_ACTION);
  assert.equal(result.searchParams.get('lang'), 'es');
  assert.equal(result.searchParams.get('foo'), 'bar');
  assert.equal(result.searchParams.get('action'), 'custom');
});

test('pickForwardHeaders normaliza encabezados relevantes', () => {
  const headers = {
    'Content-Type': 'application/json',
    'x-hub-signature-256': 'sha256=abc',
    'X-Hub-Signature': 'sha1=legacy',
    other: 'ignored'
  };
  const result = pickForwardHeaders(headers);
  assert.deepEqual(result, {
    'Content-Type': 'application/json',
    'X-Hub-Signature-256': 'sha256=abc',
    'X-Hub-Signature': 'sha1=legacy'
  });
});

test('selectForwardBody prioriza rawBody', () => {
  const req = { rawBody: Buffer.from('raw-data'), body: { hello: 'world' } };
  assert.equal(selectForwardBody(req).toString(), 'raw-data');
});

test('selectForwardBody serializa objetos si falta rawBody', () => {
  const req = { body: { hello: 'world' } };
  const body = selectForwardBody(req);
  assert.ok(body instanceof Buffer);
  assert.equal(body.toString(), JSON.stringify({ hello: 'world' }));
});

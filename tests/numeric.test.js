const test = require('node:test');
const assert = require('node:assert/strict');
const { parseNumericValue, formatLosgNumber, formatLosgList } = require('../js/utils/data-format.js');

test('parseNumericValue extracts numbers from nested totals', () => {
  const payload = { total: { value: 'Total: 128' } };
  assert.equal(parseNumericValue(payload), 128);
});

test('parseNumericValue counts array length when no numbers', () => {
  const payload = ['one', 'two', 'three'];
  assert.equal(parseNumericValue(payload), 3);
});

test('parseNumericValue returns NaN for invalid content', () => {
  const payload = { description: 'sin datos' };
  assert.ok(Number.isNaN(parseNumericValue(payload)));
});

test('formatLosgNumber reuses parseNumericValue for objects', () => {
  const payload = { count: '42 accesos' };
  assert.equal(formatLosgNumber(payload), '42');
});

test('formatLosgNumber falls back to textual list when not numeric', () => {
  const payload = { descripcion: 'Sin actividad', usuario: 'tester@example.com' };
  assert.equal(formatLosgNumber(payload), formatLosgList(payload));
});

'use strict';

process.env.TZ = 'Asia/Tokyo';
const assert = require('node:assert/strict');
const { serialize } = require('node:v8');
const { test } = require('node:test');
const metadata = require('../src/photo-metadata');
// Preserve the original extraction baseline except for the three raw XMP title
// values: ExifReader 4.46+ intentionally decodes UTF-8 values instead of mojibake.
// https://github.com/mattiasw/ExifReader/blob/v4.52.0/CHANGELOG.md#4460---2026-09-26
const baseline = require('./fixtures/photo-metadata-baseline.json');
const { createMetadataCases, normalizeResult } = require('./fixtures/photo-metadata-cases');

const cases = createMetadataCases();
test('metadata cases match the fixed pre-extraction baseline inventory', () => {
  assert.equal(baseline.timezone, process.env.TZ);
  assert.deepEqual(cases.map(entry => entry.name), baseline.metadata.map(entry => entry.name));
});

for (const [index, spec] of cases.entries()) {
  test(`metadata compatibility: ${spec.name}`, () => {
    const before = serialize(spec.args);
    let outcome;
    try { outcome = { name: spec.name, value: normalizeResult(metadata[spec.fn](...spec.args)) }; }
    catch (error) { outcome = { name: spec.name, error: { name: error.name, message: error.message } }; }
    assert.deepEqual(outcome, baseline.metadata[index]);
    assert.deepEqual(serialize(spec.args), before, 'Metadata interpretation must not mutate its input');
  });
}

'use strict';

process.env.TZ = 'Asia/Tokyo';
const assert = require('node:assert/strict');
const { test } = require('node:test');
const { createPhotoRecordBuilder } = require('../src/photo-record');
const baseline = require('./fixtures/photo-metadata-baseline.json');
const { createRecordCases, runRecordCase, readFixture, FIXED_NOW, STAT_DATE } = require('./fixtures/photo-metadata-cases');

const cases = createRecordCases();
test('record cases match the fixed pre-extraction baseline inventory', () => {
  assert.deepEqual(cases.map(entry => entry.name), baseline.records.map(entry => entry.name));
});

for (const [index, spec] of cases.entries()) {
  test(`record compatibility: ${spec.name}`, async () => {
    const result = await runRecordCase(createPhotoRecordBuilder, spec);
    assert.deepEqual({ name: spec.name, ...result }, baseline.records[index]);
  });
}

function deferred() {
  let resolve;
  const promise = new Promise(onResolve => { resolve = onResolve; });
  return { promise, resolve };
}

test('record completion awaits an in-flight thumbnail and keeps injected time', async () => {
  const thumbnail = deferred();
  const clockReached = deferred();
  const bytes = readFixture('plain.png');
  const originalBytes = Buffer.from(bytes);
  let finished = false;
  const builder = createPhotoRecordBuilder({
    readFile: async () => bytes,
    stat: async () => { throw new Error('Filename capture time must avoid stat'); },
    decodeImage: () => ({ image: null, imageWidth: 320, imageHeight: 180,
      resolutionTier: null, orientationTier: 'landscape' }),
    createThumbnail: () => thumbnail.promise,
    now: () => { clockReached.resolve(); return new Date(FIXED_NOW); },
  });
  const pending = builder('C:/fixture/VRChat_2026-06-07_22-20-00.png')
    .then(record => { finished = true; return record; });
  await clockReached.promise;
  assert.equal(finished, false);
  thumbnail.resolve('C:/fixture/completed-thumbnail.jpg');
  const record = await pending;
  assert.equal(record.thumbnailPath, 'C:/fixture/completed-thumbnail.jpg');
  assert.equal(record.createdAt, FIXED_NOW);
  assert.equal(record.updatedAt, FIXED_NOW);
  assert.deepEqual(bytes, originalBytes);
});

test('timestamp stat starts while thumbnail generation is still in flight', async () => {
  const thumbnail = deferred();
  const statStarted = deferred();
  const statResult = deferred();
  const events = [];
  const builder = createPhotoRecordBuilder({
    readFile: async () => readFixture('plain.png'),
    decodeImage: () => ({ image: null, imageWidth: null, imageHeight: null }),
    createThumbnail: () => { events.push('thumbnail:start'); return thumbnail.promise; },
    stat: () => { events.push('stat:start'); statStarted.resolve(); return statResult.promise; },
    now: () => new Date(FIXED_NOW),
  });
  const pending = builder('C:/fixture/generic.png');
  await statStarted.promise;
  assert.deepEqual(events, ['thumbnail:start', 'stat:start']);
  statResult.resolve({ mtime: new Date(STAT_DATE) });
  thumbnail.resolve(null);
  const record = await pending;
  assert.equal(record.takenAtTimestamp, new Date(STAT_DATE).getTime());
  assert.equal(record.thumbnailPath, null);
});

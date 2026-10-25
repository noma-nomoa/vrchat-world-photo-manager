'use strict';

const assert = require('node:assert/strict');
const { test } = require('node:test');
// The same assertions passed against the original main-process workflow before extraction.
const { createPhotoImportService } = require('../src/photo-import-service');

const record = (filePath, overrides = {}) => ({ filePath, fileHash: filePath,
  thumbnailPath: `${filePath}.thumb`, takenAtTimestamp: 100, year: 2026, month: 6,
  worldId: 'embedded-world', worldUrl: 'embedded-url', ...overrides });
const deferred = () => {
  let resolve;
  let reject;
  const promise = new Promise((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
};
const tick = () => new Promise(resolve => setImmediate(resolve));

function fixture(options = {}) {
  const rows = new Map((options.existing || []).map(row => [row.file_hash, { ...row }]));
  const calls = [];
  const events = [];
  let lookups = 0;
  const store = value => rows.set(value.fileHash, { file_hash: value.fileHash,
    file_path: value.filePath, world_id: value.worldId, world_url: value.worldUrl,
    ...options.persisted?.[value.fileHash] });
  const photoDb = {
    getPhotosByHashes(hashes) {
      calls.push(['lookup', [...hashes]]);
      if (options.lookupError) throw options.lookupError;
      lookups += 1;
      return [...new Set(hashes)].map(hash => rows.get(hash)).filter(Boolean);
    },
    insertOrUpdatePhotos(records) {
      calls.push(['bulk', records.map(value => value.filePath)]);
      if (options.bulkError) throw options.bulkError;
      records.forEach(store);
    },
    insertOrUpdatePhoto(value) {
      calls.push(['single', value.filePath]);
      if (options.singleError) throw options.singleError;
      store(value);
    },
    getPhotoByHash(hash) {
      calls.push(['saved', hash]);
      return options.missingSavedRow ? null : rows.get(hash);
    },
  };
  const service = createPhotoImportService({ photoDb,
    isSupportedImageFile: value => /\.(png|jpe?g|webp)$/i.test(value),
    normalizeKnownFilePath: value => String(value || '').replaceAll('\\', '/').toLowerCase(),
    buildPhotoRecord: async file => {
      calls.push(['build', file]);
      return options.build ? options.build(file) : record(file);
    },
    ensureManagedThumbnailForRow: async row => {
      calls.push(['recover', row.file_hash]);
      if (options.recoveryError) throw options.recoveryError;
      if (options.recover) await options.recover(row);
    },
    buildWorldMetadataSyncTargetsFromPhotoRecords: records => {
      calls.push(['targets', records]);
      assert(lookups >= 2, 'Sync targets must be built after reading saved identities');
      return records;
    },
    concurrencyLimit: options.concurrencyLimit || 2,
  });
  return { service, calls, events, rows, run: files => service.importFiles(files, event => events.push(event)) };
}

test('empty imports keep the established summary and progress contract', async () => {
  const h = fixture();
  assert.deepEqual(await h.run([]), { canceled: false, totalSelected: 0, importedCount: 0,
    newCount: 0, updatedCount: 0, relocatedCount: 0, worldMetadataTargets: [],
    failedCount: 0, failedFiles: [], selectedMonth: null });
  assert.deepEqual(h.events, [
    { phase: 'process', current: 0, total: 0, message: '画像を取り込み中...' },
    { phase: 'save', current: 0, total: 0, message: '取り込み結果を保存中...' },
    { phase: 'complete', current: 0, total: 0, message: '取り込みが完了しました' },
  ]);
});

test('unsupported files and duplicate exact paths are excluded without mutating input', async () => {
  const h = fixture();
  const input = Object.freeze(['A.PNG', 'notes.txt', 'A.PNG', 'b.webp', 'c.jpeg']);
  const result = await h.run(input);
  assert.equal(result.totalSelected, 3);
  assert.equal(result.newCount, 3);
  assert.deepEqual(h.calls.filter(call => call[0] === 'build').map(call => call[1]), ['A.PNG', 'b.webp', 'c.jpeg']);
  assert.deepEqual(h.calls.find(call => call[0] === 'bulk'), ['bulk', ['A.PNG', 'b.webp', 'c.jpeg']]);
});

test('all unsupported selections behave like an empty import', async () => {
  const h = fixture();
  assert.equal((await h.run(['one.txt', 'directory'])).totalSelected, 0);
  assert(!h.calls.some(call => call[0] === 'build'));
});

test('partial preprocessing failures retain input order and permit successful records', async () => {
  const gates = { 'a.png': deferred(), 'b.png': deferred(), 'c.png': deferred() };
  const h = fixture({ concurrencyLimit: 3, build: file => gates[file].promise });
  const pending = h.run(Object.keys(gates));
  gates['c.png'].resolve(record('c.png'));
  await tick();
  gates['b.png'].reject(new Error('bad metadata'));
  await tick();
  gates['a.png'].resolve(record('a.png'));
  const result = await pending;
  assert.equal(result.importedCount, 2);
  assert.deepEqual(result.failedFiles, [{ filePath: 'b.png', message: 'bad metadata' }]);
  assert.equal(result.failedCount, 1);
  assert.deepEqual(h.calls.find(call => call[0] === 'bulk'), ['bulk', ['a.png', 'c.png']]);
  assert.deepEqual(h.events.filter(event => event.phase === 'process').map(event => event.current), [0, 1, 2, 3]);
  assert.equal(h.events.find(event => event.phase === 'save').current, 2);
  assert.equal(h.events.at(-1).current, 3);
});

test('preprocessing concurrency is bounded and finished slots accept the next input', async () => {
  const gates = Object.fromEntries(['a.png', 'b.png', 'c.png', 'd.png'].map(file => [file, deferred()]));
  const h = fixture({ concurrencyLimit: 2, build: file => gates[file].promise });
  const builds = () => h.calls.filter(call => call[0] === 'build').map(call => call[1]);
  const pending = h.run(Object.keys(gates));
  assert.deepEqual(builds(), ['a.png', 'b.png']);
  gates['b.png'].resolve(record('b.png'));
  await tick();
  assert.deepEqual(builds(), ['a.png', 'b.png', 'c.png']);
  gates['c.png'].resolve(record('c.png'));
  await tick();
  assert.deepEqual(builds(), ['a.png', 'b.png', 'c.png', 'd.png']);
  assert(!h.calls.some(call => call[0] === 'bulk'));
  gates['a.png'].resolve(record('a.png'));
  gates['d.png'].resolve(record('d.png'));
  await pending;
  assert.deepEqual(h.calls.find(call => call[0] === 'bulk')[1], Object.keys(gates));
});

test('missing preprocessing error text keeps the established unknown-error message', async () => {
  const h = fixture({ build: async () => { throw {}; } });
  const result = await h.run(['bad.png']);
  assert.deepEqual(result.failedFiles, [{ filePath: 'bad.png', message: '不明なエラー' }]);
  assert.equal(result.importedCount, 0);
});

test('missing thumbnails are recovered from persisted rows before reporting completion', async () => {
  const gate = deferred();
  const h = fixture({ build: file => record(file, { thumbnailPath: '' }), recover: () => gate.promise });
  const pending = h.run(['a.png']);
  await tick();
  assert.deepEqual(h.calls.slice(-2), [['saved', 'a.png'], ['recover', 'a.png']]);
  assert(!h.events.some(event => event.phase === 'complete'));
  gate.resolve();
  await pending;
  assert.deepEqual(h.events.at(-2), { phase: 'thumbnail-recovery', current: 1, total: 1, message: '不足サムネイルを補完中...' });
  assert.equal(h.events.at(-1).phase, 'complete');
});

test('missing saved rows are skipped during thumbnail recovery', async () => {
  const h = fixture({ build: file => record(file, { thumbnailPath: null }), missingSavedRow: true });
  assert.equal((await h.run(['a.png'])).importedCount, 1);
  assert(!h.calls.some(call => call[0] === 'recover'));
  assert(!h.events.some(event => event.phase === 'thumbnail-recovery'));
});

test('bulk insertion failure preserves the existing ordered per-record retry', async () => {
  const h = fixture({ bulkError: new Error('batch unavailable') });
  const result = await h.run(['a.png', 'b.png']);
  assert.deepEqual(h.calls.filter(call => ['bulk', 'single'].includes(call[0])), [
    ['bulk', ['a.png', 'b.png']], ['single', 'a.png'], ['single', 'b.png'],
  ]);
  assert.equal(result.importedCount, 2);
});

test('per-record insertion failure propagates and never reports completion', async () => {
  const error = new Error('database unavailable');
  const h = fixture({ bulkError: error, singleError: error });
  await assert.rejects(h.run(['a.png', 'b.png']), value => value === error);
  assert.deepEqual(h.calls.filter(call => call[0] === 'single'), [['single', 'a.png']]);
  assert(!h.events.some(event => event.phase === 'complete'));
});

test('initial database lookup failure prevents writes', async () => {
  const error = new Error('lookup failed');
  const h = fixture({ lookupError: error });
  await assert.rejects(h.run(['a.png']), value => value === error);
  assert(!h.calls.some(call => call[0] === 'bulk'));
});

test('thumbnail recovery failure propagates without a successful completion event', async () => {
  const error = new Error('thumbnail recovery failed');
  const h = fixture({ build: file => record(file, { thumbnailPath: '' }), recoveryError: error });
  await assert.rejects(h.run(['a.png']), value => value === error);
  assert(!h.events.some(event => event.phase === 'complete'));
});

test('world synchronization uses the saved identity instead of embedded source metadata', async () => {
  const h = fixture({ persisted: { 'a.png': { world_id: 'manual-world', world_url: 'manual-url' } } });
  const result = await h.run(['a.png']);
  assert.deepEqual(result.worldMetadataTargets, [{ worldId: 'manual-world', worldUrl: 'manual-url' }]);
});

test('new, updated, and relocated counts use the pre-insert database snapshot', async () => {
  const h = fixture({ existing: [
    { file_hash: 'same', file_path: 'C:\\Pictures\\SAME.PNG' },
    { file_hash: 'moved', file_path: 'old/moved.png' },
  ], build: file => record(file, { fileHash: file.includes('SAME') ? 'same' : file.includes('moved') ? 'moved' : 'new' }) });
  const result = await h.run(['C:/Pictures/SAME.PNG', 'new/moved.png', 'brand-new.png']);
  assert.equal(result.newCount, 1);
  assert.equal(result.updatedCount, 2);
  assert.equal(result.relocatedCount, 1);
});

test('different selected paths with identical hashes retain per-path counts and insertion order', async () => {
  const h = fixture({ build: file => record(file, { fileHash: 'same-content' }) });
  const result = await h.run(['first.png', 'second.png']);
  assert.equal(result.importedCount, 2);
  assert.equal(result.newCount, 2);
  assert.equal(result.worldMetadataTargets.length, 1);
  assert.deepEqual(h.calls.find(call => call[0] === 'bulk')[1], ['first.png', 'second.png']);
});

test('latest capture time selects the month, with the first record winning a timestamp tie', async () => {
  const h = fixture({ build: file => record(file, file === 'old.png'
    ? { takenAtTimestamp: 1, month: 1 } : { takenAtTimestamp: 200, month: file === 'first.png' ? 2 : 3 }) });
  assert.deepEqual((await h.run(['old.png', 'first.png', 'tied.png'])).selectedMonth, { year: 2026, month: 2 });
});

test('a record without integer year/month retains the null selected-month fallback', async () => {
  const h = fixture({ build: file => record(file, { month: '6' }) });
  assert.equal((await h.run(['a.png'])).selectedMonth, null);
});

test('optional progress reporting does not alter import results', async () => {
  const h = fixture();
  assert.equal((await h.service.importFiles(['a.png'])).importedCount, 1);
});

test('progress callback failure before preprocessing prevents work', async () => {
  const h = fixture();
  const error = new Error('observer failed');
  await assert.rejects(h.service.importFiles(['a.png'], () => { throw error; }), value => value === error);
  assert.deepEqual(h.calls, []);
});

test('progress callback failure at the save boundary prevents database mutation', async () => {
  const h = fixture();
  const error = new Error('save observer failed');
  await assert.rejects(h.service.importFiles(['a.png'], event => { if (event.phase === 'save') throw error; }), value => value === error);
  assert(!h.calls.some(call => call[0] === 'bulk'));
});

test('separate imports receive independent mutable summary collections', async () => {
  const h = fixture();
  const first = await h.run([]);
  first.failedFiles.push({ filePath: 'unrelated' });
  first.worldMetadataTargets.push({ worldId: 'unrelated' });
  const second = await h.run([]);
  assert.deepEqual(second.failedFiles, []);
  assert.deepEqual(second.worldMetadataTargets, []);
});

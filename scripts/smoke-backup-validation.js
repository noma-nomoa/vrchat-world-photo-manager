// Run with Electron so better-sqlite3 uses the application's native ABI.
// Only in-memory DBs are opened; no app startup, user files, or network is used.
const assert = require('node:assert/strict');
const { initDatabase } = require('../src/db');

const TABLE_NAMES = ['photos', 'tracked_folders', 'world_metadata_cache', 'tags', 'photo_tags'];
const COUNT_KEYS = {
  photos: 'photoCount', tracked_folders: 'trackedFolderCount',
  world_metadata_cache: 'worldCacheCount', tags: 'tagCount',
};
const FIXED_TIME = '2026-09-09T00:00:00.000Z';
const baselineOnly = process.argv.includes('--baseline-only');
const results = [];

function clone(value) {
  // Preserve undefined, NaN, and infinities for direct DB API validation cases.
  return structuredClone(value);
}

function seed(db, name) {
  const photo = db.insertOrUpdatePhoto({
    filePath: `backup-fixture/${name}.png`, fileName: `${name}.png`, fileHash: `${name}-hash`,
    takenAt: '2026/09/09 09:00:00', takenAtTimestamp: 1788912000000,
    groupDate: '2026-09-09', year: 2026, month: 9, day: 9,
    worldId: `wrld_${name}`, worldName: `${name} world`, worldNameManual: `${name} manual`,
    worldUrl: `https://example.invalid/world/${name}`, thumbnailPath: `backup-fixture/${name}-thumb.png`,
    imageWidth: 320, imageHeight: 180, resolutionTier: 'other', orientationTier: 'landscape',
    printNoteText: `${name} print note`, memoText: `${name} memo`,
    createdAt: FIXED_TIME, updatedAt: FIXED_TIME,
  });
  assert.ok(photo?.id);
  db.updateFavoriteStatus(photo.id, true);
  db.replacePhotoTags(photo.id, [
    { name: `${name} tag one`, colorHex: '#12AB34' },
    { name: `${name} tag two`, colorHex: '#AB1234' },
  ]);
  db.upsertTrackedFolder(`backup-fixture/${name}/tracked`);
  db.upsertWorldMetadata({
    worldId: `wrld_${name}`, sourceUrl: `https://example.invalid/world/${name}`,
    worldNameOfficial: `${name} official`, worldDescription: `${name} description`,
    worldTags: [name, 'fixture'], authorId: `usr_${name}`, authorName: `${name} author`,
    releaseStatus: 'public', imageUrl: null, thumbnailImageUrl: null,
    fetchStatus: 'success', fetchedAt: FIXED_TIME, lastAttemptedAt: FIXED_TIME,
  });
  const snapshot = db.createBackupSnapshot();
  assert.equal(snapshot.schemaVersion, 1);
  assert.deepEqual(Object.keys(snapshot.tables).sort(), [...TABLE_NAMES].sort());
  for (const table of TABLE_NAMES) assert.ok(snapshot.tables[table].length > 0, table);
  return snapshot;
}

function withSeededDb(run) {
  const db = initDatabase(':memory:');
  try { return run(db, seed(db, 'existing')); }
  finally { db.close(); }
}

function check(name, run) {
  try {
    run();
    results.push({ name, ok: true });
    console.log(`PASS ${name}`);
  } catch (error) {
    results.push({ name, ok: false, message: String(error.message).split('\n')[0] });
    console.error(`FAIL ${name}: ${String(error.message).split('\n')[0]}`);
  }
}

function assertRoundtrip(name, input, expected) {
  check(name, () => withSeededDb((db, before) => {
    const payload = clone(input);
    const validated = db.validateBackupSnapshot(payload);
    assert.deepEqual(db.createBackupSnapshot().tables, before.tables,
      'Preflight validation must not change stored rows');
    const result = db.restoreBackupSnapshot(payload);
    const after = db.createBackupSnapshot();
    assert.deepEqual(after.tables, expected.tables, 'All restored row fields must match');
    assert.deepEqual(after.counts, expected.counts);
    assert.deepEqual(result.beforeCounts, before.counts);
    assert.deepEqual(result.afterCounts, after.counts);
    assert.equal(result.restoredPhotoCount, after.tables.photos.length);
    assert.equal(result.restoredTrackedFolderCount, after.tables.tracked_folders.length);
    assert.equal(result.restoredWorldCacheCount, after.tables.world_metadata_cache.length);
    assert.equal(result.restoredTagCount, after.tables.tags.length);
    assert.equal(result.restoredPhotoTagCount, after.tables.photo_tags.length);
    // The main process restores the preflight result, while DB callers may
    // restore the original envelope. Both entry points must stay compatible.
    db.restoreBackupSnapshot(validated);
    assert.deepEqual(db.createBackupSnapshot().tables, expected.tables);
  }));
}

function assertRejectedWithoutChanges(name, input) {
  check(name, () => withSeededDb((db, before) => {
    const payload = clone(input);
    let validationRejection = null;
    try { db.validateBackupSnapshot(payload); }
    catch (error) { validationRejection = error; }
    assert.deepEqual(db.createBackupSnapshot().tables, before.tables,
      'Preflight validation must not change stored rows');
    let rejection = null;
    try { db.restoreBackupSnapshot(payload); }
    catch (error) { rejection = error; }
    const after = db.createBackupSnapshot();
    // Check actual rows even when the restore unexpectedly reported success.
    let dataError = null;
    try { assert.deepEqual(after.tables, before.tables); }
    catch (error) { dataError = error; }
    const changedTables = TABLE_NAMES.filter((table) =>
      JSON.stringify(after.tables[table]) !== JSON.stringify(before.tables[table]));
    assert.ok(!dataError,
      `Row data changed (${changedTables.join(', ')}); rejected=${Boolean(rejection)}`);
    assert.ok(rejection, 'Malformed backup was accepted, even though row data stayed unchanged');
    assert.ok(validationRejection, 'Preflight accepted malformed data; rejection happened only during restore');
    assert.deepEqual(after.counts, before.counts);
  }));
}

function mutatedSnapshot(source, mutate) {
  const value = clone(source);
  mutate(value);
  // Keep counts accurate so row-validation cases cannot pass only because of a
  // count mismatch. photo_tags is not counted by the current snapshot format.
  for (const [table, key] of Object.entries(COUNT_KEYS)) {
    if (Array.isArray(value.tables?.[table])) value.counts[key] = value.tables[table].length;
  }
  return value;
}

function main() {
  const sourceDb = initDatabase(':memory:');
  const emptyDb = initDatabase(':memory:');
  let source;
  let empty;
  try {
    source = seed(sourceDb, 'incoming');
    empty = emptyDb.createBackupSnapshot();
  } finally { sourceDb.close(); emptyDb.close(); }

  assertRoundtrip('current snapshot roundtrip', source, source);
  assertRoundtrip('current empty snapshot clears all five tables', empty, empty);
  assertRoundtrip('data wrapper roundtrip', { data: source }, source);
  assertRoundtrip('data wrapper accepts a valid empty snapshot', { data: empty }, empty);
  assertRoundtrip('official wrapper roundtrip', {
    appName: 'WorldShot Log', backupSchemaVersion: 1, data: source,
  }, source);
  const withoutCounts = clone(source);
  delete withoutCounts.counts;
  assertRoundtrip('counts are optional', withoutCounts, source);
  assertRoundtrip('counts are informational and ignored', {
    ...source, counts: { photoCount: -100, tagCount: 'not authoritative' },
  }, source);
  assertRoundtrip('wrapper preferences object is accepted', {
    data: source, preferences: { backgroundImagePath: 'backup-fixture/background.png' },
  }, source);
  const futureColumns = clone(source);
  for (const table of TABLE_NAMES) futureColumns.tables[table][0].unknown_extension = 'ignored';
  assertRoundtrip('unknown additional columns are ignored', futureColumns, source);
  const olderColumns = clone(source);
  const olderExpected = clone(source);
  for (const column of ['image_width', 'memo_text', 'is_favorite']) {
    delete olderColumns.tables.photos[0][column];
    olderExpected.tables.photos[0][column] = column === 'is_favorite' ? 0 : null;
  }
  delete olderColumns.tables.tags[0].color_hex;
  olderExpected.tables.tags[0].color_hex = null;
  assertRoundtrip('old nullable and default columns may be absent', olderColumns, olderExpected);

  if (!baselineOnly) {
    for (const [name, value] of [
      ['empty object', {}], ['undefined snapshot', undefined], ['null snapshot', null], ['array snapshot', []],
      ['string snapshot', 'backup'], ['null data wrapper', { data: null }],
      ['empty data wrapper', { data: {} }], ['array data wrapper', { data: [] }],
    ]) assertRejectedWithoutChanges(name, value);
    for (const [name, value] of [
      ['missing schema version', undefined], ['future schema version', 2],
      ['null schema version', null], ['string schema version', '1'],
    ]) {
      const input = clone(source);
      if (value === undefined) delete input.schemaVersion;
      else input.schemaVersion = value;
      assertRejectedWithoutChanges(name, input);
    }
    for (const [name, wrapper] of [
      ['different app wrapper', { appName: 'Another App', backupSchemaVersion: 1 }],
      ['future wrapper version', { appName: 'WorldShot Log', backupSchemaVersion: 2 }],
      ['null wrapper app name', { appName: null }],
      ['string wrapper version', { backupSchemaVersion: '1' }],
    ]) assertRejectedWithoutChanges(name, { ...wrapper, data: source });
    for (const [name, value] of [
      ['null', null], ['array', []], ['string', 'preferences'], ['number', 1],
      ['boolean', false], ['undefined', undefined],
    ]) assertRejectedWithoutChanges(`${name} wrapper preferences`, { data: source, preferences: value });
    for (const [name, value] of [
      ['missing tables', undefined], ['null tables', null], ['array tables', []],
      ['string tables', 'invalid'], ['numeric tables', 1],
    ]) {
      const input = clone(source);
      if (value === undefined) delete input.tables;
      else input.tables = value;
      assertRejectedWithoutChanges(name, input);
    }
    for (const table of TABLE_NAMES) {
      assertRejectedWithoutChanges(`${table}: missing table`, mutatedSnapshot(source, (input) => {
        delete input.tables[table];
        // Missing photos/tags must be rejected even when there are no link rows
        // that would accidentally make SQLite's foreign-key check reject it.
        if (table === 'photos' || table === 'tags') input.tables.photo_tags = [];
      }));
      for (const [name, value] of [['null', null], ['object', {}], ['string', 'rows']]) {
        assertRejectedWithoutChanges(`${table}: ${name} instead of array`,
          mutatedSnapshot(source, (input) => {
            input.tables[table] = value;
            if (table === 'photos' || table === 'tags') input.tables.photo_tags = [];
          }));
      }
      for (const [name, value] of [
        ['null', null], ['array', []], ['string', 'row'], ['number', 12],
        ['empty object', {}], ['unknown fields only', { unknown_field: 'not a row' }],
      ]) {
        assertRejectedWithoutChanges(`${table}: ${name} row`,
          mutatedSnapshot(source, (input) => { input.tables[table].push(value); }));
      }
    }
    const requiredColumns = {
      photos: ['id', 'file_name'], tracked_folders: ['id', 'folder_path'],
      world_metadata_cache: ['world_id', 'fetch_status'], tags: ['id', 'normalized_name'],
      photo_tags: ['photo_id', 'tag_id', 'created_at'],
    };
    for (const [table, columns] of Object.entries(requiredColumns)) {
      for (const column of columns) {
        for (const kind of ['missing', 'null', 'wrong type']) {
          assertRejectedWithoutChanges(`${table}.${column}: ${kind}`,
            mutatedSnapshot(source, (input) => {
              const row = input.tables[table][0];
              if (kind === 'missing') delete row[column];
              else if (kind === 'null') row[column] = null;
              else row[column] = typeof row[column] === 'number' ? 'wrong integer' : 123;
            }));
        }
      }
    }
    for (const column of ['photo_id', 'tag_id']) {
      assertRejectedWithoutChanges(`photo_tags: dangling ${column}`,
        mutatedSnapshot(source, (input) => { input.tables.photo_tags[0][column] = 999999; }));
    }
    for (const [table, column] of [
      ['photos', 'id'], ['photos', 'image_width'], ['photos', 'is_favorite'],
      ['tracked_folders', 'id'], ['tags', 'id'],
      ['photo_tags', 'photo_id'], ['photo_tags', 'tag_id'],
    ]) {
      assertRejectedWithoutChanges(`${table}.${column}: numeric string is not an integer`,
        mutatedSnapshot(source, (input) => {
          input.tables[table][0][column] = String(input.tables[table][0][column]);
        }));
    }
    for (const [table, column] of [
      ['photos', 'memo_text'], ['tags', 'color_hex'], ['world_metadata_cache', 'world_description'],
    ]) {
      assertRejectedWithoutChanges(`${table}.${column}: nullable text rejects a number`,
        mutatedSnapshot(source, (input) => { input.tables[table][0][column] = 123; }));
    }
    for (const [table, column] of [
      ['photos', 'file_name'], ['photos', 'memo_text'], ['photos', 'is_favorite'],
      ['tags', 'color_hex'], ['world_metadata_cache', 'world_description'],
    ]) {
      assertRejectedWithoutChanges(`${table}.${column}: explicit undefined is invalid`,
        mutatedSnapshot(source, (input) => { input.tables[table][0][column] = undefined; }));
    }
    for (const [name, value] of [
      ['fraction', 320.5], ['unsafe integer', Number.MAX_SAFE_INTEGER + 1],
      ['NaN', NaN], ['positive infinity', Infinity], ['negative infinity', -Infinity],
    ]) {
      assertRejectedWithoutChanges(`photos.image_width: ${name}`,
        mutatedSnapshot(source, (input) => { input.tables.photos[0].image_width = value; }));
    }
    for (const table of TABLE_NAMES) {
      assertRejectedWithoutChanges(`${table}: duplicate primary key`,
        mutatedSnapshot(source, (input) => {
          const duplicate = clone(input.tables[table][0]);
          // Distinguish the primary-key test from the other UNIQUE constraints.
          if (table === 'photos') duplicate.file_hash = 'different-hash';
          if (table === 'tracked_folders') duplicate.folder_path = 'different-folder';
          if (table === 'tags') duplicate.normalized_name = 'different-normalized-tag';
          input.tables[table].push(duplicate);
        }));
    }
    for (const [table, column] of [
      ['photos', 'file_hash'], ['tracked_folders', 'folder_path'], ['tags', 'normalized_name'],
    ]) {
      assertRejectedWithoutChanges(`${table}.${column}: duplicate unique value`,
        mutatedSnapshot(source, (input) => {
          const duplicate = clone(input.tables[table][0]);
          duplicate.id = 999999;
          input.tables[table].push(duplicate);
        }));
    }
  }
  const failed = results.filter((result) => !result.ok);
  console.log(`Backup validation: ${results.length - failed.length}/${results.length} passed`);
  return failed.length === 0;
}

try { process.exit(main() ? 0 : 1); }
catch (error) { console.error(error); process.exit(1); }

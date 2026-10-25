// Contract tests for the stage-3 extraction. Run with:
// node --test scripts/test-app-data-backup-service.js
// Only fake dependencies are used: no Electron, filesystem operations, or SQLite.
// Until the service is implemented, the missing-module failure is intentional.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const path = require('node:path');
const { createAppDataBackupService } = require('../src/app-data-backup-service');

const SELECTED_FILE = path.join('synthetic', '選択した backup.json');
const AUTOMATIC_BACKUP = {
  filePath: path.join('synthetic', 'backups', 'before-restore.json'),
  fileName: 'before-restore.json',
};
const TABLE_NAMES = ['photos', 'tracked_folders', 'world_metadata_cache', 'tags', 'photo_tags'];
const clone = value => JSON.parse(JSON.stringify(value));

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function snapshot(empty = false) {
  // Minimal rows identify data flow. Real schema/row validity is covered by
  // smoke-backup-validation.js and the Electron IPC restoration smoke.
  return {
    schemaVersion: 1,
    tables: empty ? Object.fromEntries(TABLE_NAMES.map(name => [name, []])) : {
      photos: [{ id: 7, memo_text: 'restored memo' }],
      tracked_folders: [{ id: 3, folder_path: 'synthetic/tracked' }],
      world_metadata_cache: [{ world_id: 'wrld_synthetic' }],
      tags: [{ id: 8, name: 'restored label' }],
      photo_tags: [{ photo_id: 7, tag_id: 8 }],
    },
  };
}

function envelope() {
  return {
    appName: 'WorldShot Log', appVersion: '2.5.3', backupSchemaVersion: 1,
    exportedAt: '2026-09-09T00:00:00.000Z',
    preferences: { backgroundImagePath: 'synthetic/background.png', customPreference: 'preserved' },
    data: snapshot(),
  };
}

function counts(tables) {
  return {
    photoCount: tables.photos.length,
    trackedFolderCount: tables.tracked_folders.length,
    worldCacheCount: tables.world_metadata_cache.length,
    tagCount: tables.tags.length,
  };
}

function harness(options = {}) {
  const h = {
    events: [], backups: [], validatedInput: null, restoredInput: null,
    restoreResult: null,
    error: options.error || new Error(`failure at ${options.failAt || 'unspecified step'}`),
    automaticStarted: deferred(), preferencesStarted: deferred(),
    persisted: { database: { existing: 'original database' }, preferences: { existing: 'original preferences' } },
  };
  h.input = Object.prototype.hasOwnProperty.call(options, 'input') ? options.input : envelope();
  h.initialPersisted = clone(h.persisted);
  const event = (name, value) => {
    h.events.push({ name, value });
    if (options.failAt === name) throw h.error;
  };
  h.names = () => h.events.map(item => item.name);
  h.service = createAppDataBackupService({
    selectRestoreFile: async () => {
      event('select');
      return Object.prototype.hasOwnProperty.call(options, 'selectedFile') ? options.selectedFile : SELECTED_FILE;
    },
    readText: async filePath => {
      event('read', filePath);
      return Object.prototype.hasOwnProperty.call(options, 'rawText') ? options.rawText : JSON.stringify(h.input);
    },
    photoDb: {
      validateBackupSnapshot: input => {
        h.validatedInput = input;
        event('validate', input);
        // The real DB validator returns the inner snapshot without preferences.
        return Object.prototype.hasOwnProperty.call(input, 'tables') ? input : input.data;
      },
      restoreBackupSnapshot: input => {
        h.restoredInput = input;
        event('restore', input);
        const restoredSnapshot = Object.prototype.hasOwnProperty.call(input, 'tables') ? input : input.data;
        h.persisted.database = clone(restoredSnapshot.tables);
        h.restoreResult = {
          beforeCounts: { photoCount: 12, trackedFolderCount: 2, worldCacheCount: 4, tagCount: 3 },
          afterCounts: counts(restoredSnapshot.tables),
          restoredPhotoCount: restoredSnapshot.tables.photos.length,
          restoredTrackedFolderCount: restoredSnapshot.tables.tracked_folders.length,
          restoredWorldCacheCount: restoredSnapshot.tables.world_metadata_cache.length,
          restoredTagCount: restoredSnapshot.tables.tags.length,
          restoredPhotoTagCount: restoredSnapshot.tables.photo_tags.length,
        };
        return h.restoreResult;
      },
    },
    createAutomaticBackup: async reason => {
      event('automatic-backup', reason);
      h.automaticStarted.resolve();
      if (options.automaticWaiter) await options.automaticWaiter;
      h.backups.push(clone(h.persisted));
      return { ...AUTOMATIC_BACKUP };
    },
    onDatabaseRestored: () => event('database-restored'),
    writePreferences: async preferences => {
      event('write-preferences', preferences);
      h.preferencesStarted.resolve();
      if (options.preferencesWaiter) await options.preferencesWaiter;
      h.persisted.preferences = clone(preferences);
    },
  });
  return h;
}

function expectedResponse(h) {
  return {
    ok: true,
    canceled: false,
    filePath: SELECTED_FILE,
    fileName: path.basename(SELECTED_FILE),
    automaticBackupFilePath: AUTOMATIC_BACKUP.filePath,
    automaticBackupFileName: AUTOMATIC_BACKUP.fileName,
    ...h.restoreResult.afterCounts,
    restoredPhotoCount: h.restoreResult.restoredPhotoCount,
    restoredTrackedFolderCount: h.restoreResult.restoredTrackedFolderCount,
    restoredWorldCacheCount: h.restoreResult.restoredWorldCacheCount,
    restoredTagCount: h.restoreResult.restoredTagCount,
  };
}

test('creating the service performs no I/O or persistence operations', () => {
  const h = harness();
  assert.equal(typeof h.service.restoreBackup, 'function');
  assert.deepEqual(h.events, []);
  assert.deepEqual(h.persisted, h.initialPersisted);
});

test('cancelled selection returns the existing cancellation response without reading or writing', async () => {
  const h = harness({ selectedFile: null });
  assert.deepEqual(await h.service.restoreBackup(), { ok: true, canceled: true });
  assert.deepEqual(h.names(), ['select']);
  assert.deepEqual(h.persisted, h.initialPersisted);
  assert.deepEqual(h.backups, []);
});

for (const [failAt, expectedEvents] of [
  ['select', ['select']],
  ['read', ['select', 'read']],
]) {
  test(`${failAt} failure is propagated without validation or persistence changes`, async () => {
    const h = harness({ failAt });
    await assert.rejects(h.service.restoreBackup(), error => error === h.error);
    assert.deepEqual(h.names(), expectedEvents);
    assert.deepEqual(h.persisted, h.initialPersisted);
    assert.deepEqual(h.backups, []);
  });
}

test('malformed JSON is rejected before DB validation and before automatic backup', async () => {
  const h = harness({ rawText: '{"unfinished":' });
  await assert.rejects(h.service.restoreBackup(), SyntaxError);
  assert.deepEqual(h.names(), ['select', 'read']);
  assert.deepEqual(h.persisted, h.initialPersisted);
  assert.deepEqual(h.backups, []);
});

for (const input of [null, true, 12, 'not a backup']) {
  test(`root ${JSON.stringify(input)} keeps the existing unreadable-backup message`, async () => {
    const h = harness({ input });
    await assert.rejects(h.service.restoreBackup(), { message: 'バックアップファイルを読み込めませんでした' });
    assert.deepEqual(h.names(), ['select', 'read']);
    assert.deepEqual(h.persisted, h.initialPersisted);
    assert.deepEqual(h.backups, []);
  });
}

test('another application keeps the existing application-mismatch message', async () => {
  const h = harness({ input: { ...envelope(), appName: 'Another application' } });
  await assert.rejects(h.service.restoreBackup(), { message: 'WorldShot Log のバックアップではありません' });
  assert.deepEqual(h.names(), ['select', 'read']);
  assert.deepEqual(h.persisted, h.initialPersisted);
  assert.deepEqual(h.backups, []);
});

for (const [name, input] of [
  ['array root', []],
  ['empty object', {}],
  ['future schema version', { ...envelope(), backupSchemaVersion: 2 }],
  ['invalid preference array', { ...envelope(), preferences: [] }],
  ['invalid row data', { ...envelope(), data: { schemaVersion: 1, tables: { photos: [null] } } }],
]) {
  test(`${name} reaches the actual validator boundary before any automatic backup`, async () => {
    const h = harness({ input, failAt: 'validate' });
    await assert.rejects(h.service.restoreBackup(), error => error === h.error);
    assert.deepEqual(h.names(), ['select', 'read', 'validate']);
    assert.deepEqual(h.validatedInput, input);
    assert.deepEqual(h.persisted, h.initialPersisted);
    assert.deepEqual(h.backups, []);
  });
}

test('official envelope preserves operation order, all metadata, and the existing IPC response', async () => {
  const h = harness();
  const result = await h.service.restoreBackup();
  assert.deepEqual(h.names(), [
    'select', 'read', 'validate', 'automatic-backup', 'restore', 'database-restored', 'write-preferences',
  ]);
  assert.equal(h.events.find(item => item.name === 'read').value, SELECTED_FILE);
  assert.equal(h.events.find(item => item.name === 'automatic-backup').value, 'before-restore');
  assert.equal(h.restoredInput, h.validatedInput,
    'Restore must receive the original parsed envelope, not the validator return value');
  assert.deepEqual(h.restoredInput, h.input);
  assert.equal(h.events.find(item => item.name === 'write-preferences').value, h.validatedInput.preferences);
  assert.deepEqual(h.backups, [h.initialPersisted]);
  assert.deepEqual(h.persisted.database, h.input.data.tables);
  assert.deepEqual(h.persisted.preferences, h.input.preferences);
  assert.deepEqual(result, expectedResponse(h));
});

for (const [name, input, expectedTables] of [
  ['bare snapshot', snapshot(), snapshot().tables],
  ['data wrapper', { data: snapshot() }, snapshot().tables],
  ['valid empty snapshot', snapshot(true), snapshot(true).tables],
]) {
  test(`${name} restores without changing preferences`, async () => {
    const h = harness({ input });
    assert.deepEqual(await h.service.restoreBackup(), expectedResponse(h));
    assert.deepEqual(h.names(), ['select', 'read', 'validate', 'automatic-backup', 'restore', 'database-restored']);
    assert.equal(h.restoredInput, h.validatedInput);
    assert.deepEqual(h.persisted.database, expectedTables);
    assert.deepEqual(h.persisted.preferences, h.initialPersisted.preferences);
    assert.deepEqual(h.backups, [h.initialPersisted]);
  });
}

test('automatic backup failure prevents restore, cache invalidation, and preference writes', async () => {
  const h = harness({ failAt: 'automatic-backup' });
  await assert.rejects(h.service.restoreBackup(), error => error === h.error);
  assert.deepEqual(h.names(), ['select', 'read', 'validate', 'automatic-backup']);
  assert.deepEqual(h.persisted, h.initialPersisted);
  assert.deepEqual(h.backups, []);
});

test('restore waits for automatic backup to finish before changing the DB', { timeout: 5000 }, async () => {
  const gate = deferred();
  const h = harness({ automaticWaiter: gate.promise });
  const pending = h.service.restoreBackup();
  await h.automaticStarted.promise;
  assert.deepEqual(h.names(), ['select', 'read', 'validate', 'automatic-backup']);
  assert.deepEqual(h.persisted, h.initialPersisted);
  gate.resolve();
  assert.deepEqual(await pending, expectedResponse(h));
  assert.deepEqual(h.backups, [h.initialPersisted]);
});

test('DB restore failure retains its automatic backup and skips all post-restore work', async () => {
  const h = harness({ failAt: 'restore' });
  await assert.rejects(h.service.restoreBackup(), error => error === h.error);
  assert.deepEqual(h.names(), ['select', 'read', 'validate', 'automatic-backup', 'restore']);
  assert.deepEqual(h.backups, [h.initialPersisted]);
  assert.deepEqual(h.persisted, h.initialPersisted);
});

test('post-restore callback failure propagates and skips preference writes', async () => {
  const h = harness({ failAt: 'database-restored' });
  await assert.rejects(h.service.restoreBackup(), error => error === h.error);
  assert.deepEqual(h.names(), ['select', 'read', 'validate', 'automatic-backup', 'restore', 'database-restored']);
  assert.deepEqual(h.persisted.database, h.input.data.tables);
  assert.deepEqual(h.persisted.preferences, h.initialPersisted.preferences);
  assert.deepEqual(h.backups, [h.initialPersisted]);
});

test('preference failure is propagated after DB restore and leaves the pre-restore backup available', async () => {
  const h = harness({ failAt: 'write-preferences' });
  await assert.rejects(h.service.restoreBackup(), error => error === h.error);
  assert.deepEqual(h.names(), [
    'select', 'read', 'validate', 'automatic-backup', 'restore', 'database-restored', 'write-preferences',
  ]);
  assert.deepEqual(h.persisted.database, h.input.data.tables);
  assert.deepEqual(h.persisted.preferences, h.initialPersisted.preferences);
  assert.deepEqual(h.backups, [h.initialPersisted]);
});

test('successful return waits for asynchronous preference persistence', { timeout: 5000 }, async () => {
  const gate = deferred();
  const h = harness({ preferencesWaiter: gate.promise });
  let settled = false;
  const pending = h.service.restoreBackup().then(value => { settled = true; return value; });
  await h.preferencesStarted.promise;
  assert.equal(settled, false);
  assert.deepEqual(h.persisted.database, h.input.data.tables);
  assert.deepEqual(h.persisted.preferences, h.initialPersisted.preferences);
  gate.resolve();
  assert.deepEqual(await pending, expectedResponse(h));
  assert.equal(settled, true);
  assert.deepEqual(h.persisted.preferences, h.input.preferences);
});

// Run with Electron so better-sqlite3 uses the application's native ABI.
// Only temporary SQLite databases are opened; the application and network are not started.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createHash } = require('node:crypto');
const Database = require('better-sqlite3');
const { initDatabase } = require('../src/db');

const RUN_ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-world-identity-'));
const DB_SOURCE_SHA256 = createHash('sha256')
  .update(fs.readFileSync(path.resolve(__dirname, '..', 'src', 'db.js'))).digest('hex');
const WORLD_A = 'wrld_aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
const WORLD_B = 'wrld_bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';
const WORLD_C = 'wrld_cccccccc-cccc-cccc-cccc-cccccccccccc';
const results = [];
const url = id => id ? `https://vrchat.com/home/world/${id}` : null;

function photo(key = 'photo', overrides = {}) {
  return {
    filePath: path.join(RUN_ROOT, `${key}.png`), fileName: `${key}.png`, fileHash: key,
    takenAt: '2026/06/07 22:20:00', takenAtTimestamp: new Date(2026, 5, 7, 22, 20).getTime(),
    groupDate: '2026-06-07', year: 2026, month: 6, day: 7,
    worldId: WORLD_A, worldName: 'Extracted A', worldNameManual: null, worldUrl: url(WORLD_A),
    worldIdentitySource: 'automatic', thumbnailPath: null,
    imageWidth: 320, imageHeight: 180, resolutionTier: 'other', orientationTier: 'landscape',
    printNoteText: '', memoText: null,
    createdAt: '2026-06-07T13:20:00.000Z', updatedAt: '2026-06-07T13:20:00.000Z',
    ...overrides,
  };
}

function assertIdentity(row, worldId, worldName, worldUrl = url(worldId)) {
  assert(row, 'Expected a saved photo row');
  assert.deepEqual(
    { worldId: row.world_id, worldName: row.world_name, worldUrl: row.world_url },
    { worldId, worldName, worldUrl },
    'ID, URL and automatic name must stay associated with the same world'
  );
}

function withDatabase(callback, seed = null) {
  const dbPath = path.join(RUN_ROOT, `case-${results.length}.sqlite`);
  if (seed) seed(dbPath);
  const db = initDatabase(dbPath);
  try {
    callback(db);
  } finally {
    db.close();
  }
}

function test(name, callback) {
  try {
    callback();
    results.push({ name, ok: true });
    console.log(`PASS: ${name}`);
  } catch (error) {
    results.push({ name, ok: false, error: error.message });
    console.error(`FAIL: ${name}\n${error.message}`);
  }
}

function setManualWorld(db, id, name = null) {
  return db.updateManualWorldSettings(id, {
    worldNameManual: name, worldId: WORLD_B, worldUrl: url(WORLD_B),
  });
}

function seedLegacy(dbPath, { worldId = WORLD_B, worldUrl = url(WORLD_B) } = {}) {
  const db = new Database(dbPath);
  try {
    // Deliberately predates world_identity_source and the other optional migrations.
    db.exec(`CREATE TABLE photos (
      id INTEGER PRIMARY KEY AUTOINCREMENT, file_path TEXT NOT NULL,
      file_name TEXT NOT NULL, file_hash TEXT NOT NULL UNIQUE,
      taken_at TEXT NOT NULL, taken_at_timestamp INTEGER NOT NULL,
      group_date TEXT NOT NULL, year INTEGER NOT NULL, month INTEGER NOT NULL, day INTEGER NOT NULL,
      world_id TEXT, world_name TEXT, world_url TEXT, thumbnail_path TEXT,
      created_at TEXT NOT NULL, updated_at TEXT NOT NULL
    )`);
    const fixture = photo('legacy');
    db.prepare(`INSERT INTO photos (file_path, file_name, file_hash, taken_at, taken_at_timestamp,
      group_date, year, month, day, world_id, world_name, world_url, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`).run(
      fixture.filePath, fixture.fileName, fixture.fileHash, fixture.takenAt,
      fixture.takenAtTimestamp, fixture.groupDate, fixture.year, fixture.month, fixture.day,
      worldId, 'Legacy B', worldUrl, fixture.createdAt, fixture.updatedAt
    );
  } finally {
    db.close();
  }
}

test('baseline: automatic identity can refresh and become empty', () => withDatabase(db => {
  db.insertOrUpdatePhoto(photo());
  db.insertOrUpdatePhoto(photo('photo', {
    worldId: WORLD_C, worldName: 'Extracted C', worldUrl: url(WORLD_C),
  }));
  assertIdentity(db.getPhotoByHash('photo'), WORLD_C, 'Extracted C');
  db.insertOrUpdatePhotos([photo('photo', { worldId: null, worldName: null, worldUrl: null })]);
  assertIdentity(db.getPhotoByHash('photo'), null, null);
}));

test('baseline: reimport preserves manual name, favorite and memo', () => withDatabase(db => {
  const row = db.insertOrUpdatePhoto(photo('photo', { memoText: 'Keep my memo' }));
  db.updateManualWorldName(row.id, 'My custom name');
  db.updateFavoriteStatus(row.id, true);
  db.insertOrUpdatePhoto(photo());
  const saved = db.getPhotoById(row.id);
  assert.equal(saved.world_name_manual, 'My custom name');
  assert.equal(saved.is_favorite, 1);
  assert.equal(saved.memo_text, 'Keep my memo');
}));

test('new extraction has automatic provenance; name-only edits do not change it', () => withDatabase(db => {
  const row = db.insertOrUpdatePhoto(photo());
  assert.equal(row.world_identity_source, 'automatic');
  assert.equal(db.getPhotosByMonth(2026, 6)[0].world_identity_source, 'automatic',
    'Listing rows must include provenance for renderer/editor metadata transfer');
  assert.equal(db.updateManualWorldName(row.id, 'Name only').world_identity_source, 'automatic');
  assert.equal(db.updateManualWorldName(row.id, '').world_identity_source, 'automatic');
}));

function renameThroughWorldSettings(db, photoId, expectedSource) {
  // The UI uses update-world-settings for both a name change and its clear
  // action, submitting the current URL even when the user never edited it.
  for (const name of ['Name from the editor', '']) {
    const before = db.getPhotoById(photoId);
    const after = db.updateManualWorldSettings(photoId, {
      worldNameManual: name, worldId: before.world_id, worldUrl: before.world_url,
    });
    assert.equal(after.world_identity_source, expectedSource,
      `${name ? 'Adding' : 'Clearing'} only the name must retain identity provenance`);
    assert.equal(after.world_name_manual, name || null);
    assertIdentity(after, before.world_id, before.world_name, before.world_url);
  }
}

test('UI name add/clear preserves automatic identity updates', () => withDatabase(db => {
  const row = db.insertOrUpdatePhoto(photo());
  renameThroughWorldSettings(db, row.id, 'automatic');
  db.insertOrUpdatePhotos([photo('photo', {
    worldId: WORLD_C, worldName: 'Extracted C', worldUrl: url(WORLD_C),
  })]);
  assertIdentity(db.getPhotoById(row.id), WORLD_C, 'Extracted C');
  assert.equal(db.getPhotoById(row.id).world_identity_source, 'automatic');
}));

test('UI name add/clear does not lock an empty automatic identity', () => withDatabase(db => {
  const row = db.insertOrUpdatePhoto(photo('photo', {
    worldId: null, worldName: null, worldUrl: null,
  }));
  renameThroughWorldSettings(db, row.id, 'automatic');
  db.insertOrUpdatePhoto(photo());
  assertIdentity(db.getPhotoById(row.id), WORLD_A, 'Extracted A');
  assert.equal(db.getPhotoById(row.id).world_identity_source, 'automatic');
}));

test('UI name add/clear does not lock an empty legacy identity', () => withDatabase(db => {
  const row = db.getPhotoByHash('legacy');
  renameThroughWorldSettings(db, row.id, 'legacy');
  db.insertOrUpdatePhoto(photo('legacy'));
  assertIdentity(db.getPhotoById(row.id), WORLD_A, 'Extracted A');
  assert.equal(db.getPhotoById(row.id).world_identity_source, 'automatic');
}, dbPath => seedLegacy(dbPath, { worldId: null, worldUrl: null })));

test('UI name add/clear keeps an existing manual identity protected', () => withDatabase(db => {
  const row = db.insertOrUpdatePhoto(photo());
  const manual = setManualWorld(db, row.id);
  renameThroughWorldSettings(db, row.id, 'manual');
  db.insertOrUpdatePhoto(photo());
  assertIdentity(db.getPhotoById(row.id), WORLD_B, manual.world_name);
  assert.equal(db.getPhotoById(row.id).world_identity_source, 'manual');
}));

test('UI name add/clear keeps an existing legacy identity protected', () => withDatabase(db => {
  const row = db.getPhotoByHash('legacy');
  renameThroughWorldSettings(db, row.id, 'legacy');
  db.insertOrUpdatePhoto(photo('legacy'));
  assertIdentity(db.getPhotoById(row.id), WORLD_B, 'Legacy B');
  assert.equal(db.getPhotoById(row.id).world_identity_source, 'legacy');
}, seedLegacy));

for (const method of ['insertOrUpdatePhoto', 'insertOrUpdatePhotos']) {
  for (const extractedWorld of ['different', 'empty']) {
    test(`manual link survives ${method} with ${extractedWorld} extracted identity`, () => withDatabase(db => {
      const row = db.insertOrUpdatePhoto(photo());
      const manual = setManualWorld(db, row.id);
      const incoming = photo('photo', extractedWorld === 'empty'
        ? { worldId: null, worldName: null, worldUrl: null } : {});
      db[method](method === 'insertOrUpdatePhotos' ? [incoming] : incoming);
      const saved = db.getPhotoById(row.id);
      assertIdentity(saved, WORLD_B, manual.world_name);
      assert.equal(saved.world_identity_source, 'manual');
    }));
  }
}

test('clearing only the manual name never unlocks a manual link', () => withDatabase(db => {
  const row = db.insertOrUpdatePhoto(photo());
  const manual = setManualWorld(db, row.id, 'My name');
  db.updateManualWorldName(row.id, '');
  db.insertOrUpdatePhoto(photo());
  const saved = db.getPhotoById(row.id);
  assertIdentity(saved, WORLD_B, manual.world_name);
  assert.equal(saved.world_name_manual, null);
  assert.equal(saved.world_identity_source, 'manual');
}));

test('explicit reread protects its resolved identity and clears manual name', () => withDatabase(db => {
  const row = db.insertOrUpdatePhoto(photo());
  db.updateManualWorldName(row.id, 'Temporary name');
  db.updateAutoWorldInfo(row.id, { worldId: WORLD_B, worldName: 'Official B', worldUrl: url(WORLD_B) });
  db.insertOrUpdatePhoto(photo());
  const saved = db.getPhotoById(row.id);
  assertIdentity(saved, WORLD_B, 'Official B');
  assert.equal(saved.world_name_manual, null);
  assert.equal(saved.world_identity_source, 'manual');
}));

test('official sync updates the preserved world name without unlocking its link', () => withDatabase(db => {
  const row = db.insertOrUpdatePhoto(photo());
  setManualWorld(db, row.id);
  db.updateAutoWorldInfoByWorldId(WORLD_B, { worldName: 'Official B', worldUrl: url(WORLD_B) });
  assertIdentity(db.getPhotoById(row.id), WORLD_B, 'Official B');
  db.insertOrUpdatePhoto(photo());
  assertIdentity(db.getPhotoById(row.id), WORLD_B, 'Official B');
  assert.equal(db.getPhotoById(row.id).world_identity_source, 'manual');
}));

for (const legacy of [
  { name: 'ID and URL', worldId: WORLD_B, worldUrl: url(WORLD_B) },
  { name: 'ID only', worldId: WORLD_B, worldUrl: null },
  { name: 'URL only', worldId: null, worldUrl: url(WORLD_B) },
]) {
  test(`migration protects legacy identity with ${legacy.name}`, () => withDatabase(db => {
    db.insertOrUpdatePhoto(photo('legacy'));
    const saved = db.getPhotoByHash('legacy');
    assertIdentity(saved, legacy.worldId, 'Legacy B', legacy.worldUrl);
    assert.equal(saved.world_identity_source, 'legacy');
  }, dbPath => seedLegacy(dbPath, legacy)));
}

test('legacy with no ID or URL can be populated by new extraction', () => withDatabase(db => {
  db.insertOrUpdatePhoto(photo('legacy'));
  const saved = db.getPhotoByHash('legacy');
  assertIdentity(saved, WORLD_A, 'Extracted A');
  assert.equal(saved.world_identity_source, 'automatic');
}, dbPath => seedLegacy(dbPath, { worldId: null, worldUrl: null })));

test('old backup without provenance restores as protected legacy', () => withDatabase(db => {
  const row = db.insertOrUpdatePhoto(photo());
  setManualWorld(db, row.id);
  const backup = db.createBackupSnapshot();
  for (const backupRow of backup.tables.photos) delete backupRow.world_identity_source;
  db.restoreBackupSnapshot(backup);
  db.insertOrUpdatePhoto(photo());
  const restored = db.getPhotoByHash('photo');
  assertIdentity(restored, WORLD_B, 'Extracted A');
  assert.equal(restored.world_identity_source, 'legacy');
}));

test('new backup roundtrip preserves manual and automatic provenance', () => withDatabase(db => {
  const manual = db.insertOrUpdatePhoto(photo('manual'));
  setManualWorld(db, manual.id);
  db.insertOrUpdatePhoto(photo('automatic'));
  db.restoreBackupSnapshot(db.createBackupSnapshot());
  db.insertOrUpdatePhotos([photo('manual'), photo('automatic', {
    worldId: WORLD_C, worldName: 'Extracted C', worldUrl: url(WORLD_C),
  })]);
  assertIdentity(db.getPhotoByHash('manual'), WORLD_B, 'Extracted A');
  assert.equal(db.getPhotoByHash('manual').world_identity_source, 'manual');
  assertIdentity(db.getPhotoByHash('automatic'), WORLD_C, 'Extracted C');
  assert.equal(db.getPhotoByHash('automatic').world_identity_source, 'automatic');
}));

for (const source of ['manual', 'legacy', 'automatic']) {
  test(`edited photo inherits ${source} provenance before later reimport`, () => withDatabase(db => {
    const row = db.insertOrUpdatePhoto(photo('edited'));
    db.updateEditedPhotoSourceMetadata(row.id, {
      ...photo('edited'), worldId: WORLD_B, worldName: 'Source B',
      worldUrl: url(WORLD_B), worldIdentitySource: source,
    });
    assert.equal(db.getPhotoById(row.id).world_identity_source, source,
      'Saving source metadata must copy its provenance');
    db.insertOrUpdatePhoto(photo('edited'));
    assertIdentity(db.getPhotoById(row.id), source === 'automatic' ? WORLD_A : WORLD_B,
      source === 'automatic' ? 'Extracted A' : 'Source B');
  }));
}

const failed = results.filter(result => !result.ok);
const reportPath = path.join(RUN_ROOT, 'result.json');
fs.writeFileSync(reportPath, JSON.stringify({ ok: failed.length === 0,
  dbSourceSha256: DB_SOURCE_SHA256, passed: results.length - failed.length,
  failed: failed.length, results }, null, 2));
console.log(`World identity smoke: ${results.length - failed.length}/${results.length} passed. ${reportPath}`);
process.exit(failed.length > 0 ? 1 : 0);

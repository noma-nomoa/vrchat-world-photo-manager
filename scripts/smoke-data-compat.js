const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const Database = require('better-sqlite3');
const { initDatabase } = require('../src/db');

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function createLegacyDatabase(dbPath) {
  const db = new Database(dbPath);
  db.exec(`
    CREATE TABLE photos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      file_path TEXT NOT NULL,
      file_name TEXT NOT NULL,
      file_hash TEXT NOT NULL UNIQUE,
      taken_at TEXT NOT NULL,
      taken_at_timestamp INTEGER NOT NULL,
      group_date TEXT NOT NULL,
      year INTEGER NOT NULL,
      month INTEGER NOT NULL,
      day INTEGER NOT NULL,
      world_id TEXT,
      world_name TEXT,
      world_url TEXT,
      thumbnail_path TEXT,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE tracked_folders (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      folder_path TEXT NOT NULL UNIQUE,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL,
      last_imported_at TEXT NOT NULL
    );

    CREATE TABLE world_metadata_cache (
      world_id TEXT PRIMARY KEY,
      source_url TEXT,
      world_name_official TEXT,
      world_description TEXT,
      world_tags_json TEXT,
      author_id TEXT,
      author_name TEXT,
      release_status TEXT,
      image_url TEXT,
      thumbnail_image_url TEXT,
      fetch_status TEXT NOT NULL,
      fetch_error TEXT,
      fetched_at TEXT,
      last_attempted_at TEXT NOT NULL,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE tags (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      normalized_name TEXT NOT NULL UNIQUE,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE photo_tags (
      photo_id INTEGER NOT NULL,
      tag_id INTEGER NOT NULL,
      created_at TEXT NOT NULL,
      PRIMARY KEY (photo_id, tag_id)
    );
  `);

  const nowIso = new Date().toISOString();
  db.prepare(`
    INSERT INTO photos (
      file_path, file_name, file_hash, taken_at, taken_at_timestamp,
      group_date, year, month, day, world_id, world_name, world_url,
      thumbnail_path, created_at, updated_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    path.join(path.dirname(dbPath), 'legacy-photo.png'),
    'legacy-photo.png',
    'legacy-hash',
    '2026/06/07 22:20:00',
    new Date(2026, 5, 7, 22, 20, 0).getTime(),
    '2026-06-07',
    2026,
    6,
    7,
    'wrld_legacy',
    'Legacy World',
    'https://vrchat.com/home/world/wrld_legacy',
    null,
    nowIso,
    nowIso
  );
  db.close();
}

function getColumnNames(dbPath, tableName) {
  const db = new Database(dbPath, { readonly: true });
  try {
    return db
      .prepare(`PRAGMA table_info(${tableName})`)
      .all()
      .map((column) => column.name);
  } finally {
    db.close();
  }
}

function assertMigratedColumns(dbPath) {
  const photoColumns = new Set(getColumnNames(dbPath, 'photos'));
  [
    'world_name_manual',
    'is_favorite',
    'image_width',
    'image_height',
    'resolution_tier',
    'orientation_tier',
    'print_note_text',
    'memo_text',
  ].forEach((columnName) => {
    assert(photoColumns.has(columnName), `Missing migrated column: ${columnName}`);
  });

  const tagColumns = new Set(getColumnNames(dbPath, 'tags'));
  assert(tagColumns.has('color_hex'), 'Missing migrated column: color_hex');
}

function seedCurrentData(db, sampleRoot) {
  const nowIso = new Date().toISOString();
  const photo = db.insertOrUpdatePhoto({
    filePath: path.join(sampleRoot, 'current-photo.png'),
    fileName: 'current-photo.png',
    fileHash: 'current-hash',
    takenAt: '2026/06/08 00:10:00',
    takenAtTimestamp: new Date(2026, 5, 8, 0, 10, 0).getTime(),
    groupDate: '2026-06-08',
    year: 2026,
    month: 6,
    day: 8,
    worldId: 'wrld_current',
    worldName: 'Current World',
    worldNameManual: null,
    worldUrl: 'https://vrchat.com/home/world/wrld_current',
    thumbnailPath: path.join(sampleRoot, 'current-thumb.png'),
    imageWidth: 3840,
    imageHeight: 2160,
    resolutionTier: '4K',
    orientationTier: 'landscape',
    printNoteText: 'print note',
    memoText: 'memo',
    createdAt: nowIso,
    updatedAt: nowIso,
  });

  assert(photo?.id, 'Current photo was not inserted.');
  db.updateFavoriteStatus(photo.id, true);
  db.updateManualWorldName(photo.id, 'Manual Current World');
  db.replacePhotoTags(photo.id, [
    { name: 'integral', colorHex: '#FF3344' },
    { name: 'photo walk', colorHex: '#14B8A6' },
  ]);
  db.upsertTrackedFolder(path.join(sampleRoot, 'Tracked'));
  db.upsertWorldMetadata({
    worldId: 'wrld_current',
    sourceUrl: 'https://vrchat.com/home/world/wrld_current',
    worldNameOfficial: 'Current World Official',
    worldDescription: 'Compatibility smoke sample.',
    worldTags: ['debug', 'compat'],
    authorId: 'usr_compat',
    authorName: 'WorldShot Log',
    releaseStatus: 'public',
    imageUrl: 'https://example.com/world.png',
    thumbnailImageUrl: 'https://example.com/world-thumb.png',
    fetchStatus: 'success',
  });

  return photo.id;
}

function assertRestoredData(db) {
  const summary = db.getApplicationDataSummary();
  assert(summary.photoCount === 2, `Expected 2 photos, got ${summary.photoCount}`);
  assert(
    summary.trackedFolderCount === 1,
    `Expected 1 tracked folder, got ${summary.trackedFolderCount}`
  );
  assert(
    summary.worldCacheCount === 1,
    `Expected 1 world cache row, got ${summary.worldCacheCount}`
  );
  assert(summary.tagCount === 2, `Expected 2 tags, got ${summary.tagCount}`);

  const photos = db.getAllPhotosWithWorldInfo();
  const current = photos.find((photo) => photo.file_hash === 'current-hash');
  const legacy = photos.find((photo) => photo.file_hash === 'legacy-hash');
  assert(current, 'Current photo was not restored.');
  assert(legacy, 'Legacy photo was not preserved.');
  assert(current.world_name_manual === 'Manual Current World', 'Manual world name was not restored.');
  assert(current.is_favorite === 1, 'Favorite flag was not restored.');
  assert(current.memo_text === 'memo', 'Memo was not restored.');
  assert(current.print_note_text === 'print note', 'Print note was not restored.');

  const tags = db.getPhotoTags(current.id);
  assert(tags.length === 2, `Expected 2 tags, got ${tags.length}`);
  assert(tags.some((tag) => tag.name === 'integral' && tag.color_hex === '#FF3344'), 'Colored tag was not restored.');
  assert(db.getTrackedFolders().length === 1, 'Tracked folder was not restored.');
  assert(db.getWorldMetadataByWorldId('wrld_current'), 'World metadata was not restored.');
}

function main() {
  const runRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-data-compat-'));
  const dbPath = path.join(runRoot, 'app.sqlite');

  createLegacyDatabase(dbPath);

  let db = initDatabase(dbPath);
  try {
    assertMigratedColumns(dbPath);
    assert(db.getApplicationDataSummary().photoCount === 1, 'Legacy row was not preserved.');

    seedCurrentData(db, runRoot);
    const backup = db.createBackupSnapshot();
    assert(backup.counts.photoCount === 2, 'Backup photo count is incorrect.');
    assert(backup.counts.tagCount === 2, 'Backup tag count is incorrect.');

    const resetCounts = db.resetApplicationData();
    assert(resetCounts.photoCount === 2, 'Reset did not report original photo count.');
    assert(db.getApplicationDataSummary().photoCount === 0, 'Reset did not clear photos.');

    const restoreResult = db.restoreBackupSnapshot({ data: backup });
    assert(restoreResult.restoredPhotoCount === 2, 'Restore photo count is incorrect.');
    assertRestoredData(db);
  } finally {
    db.close();
  }

  console.log(`Data compatibility smoke passed: ${runRoot}`);
}

try {
  main();
  process.exit(0);
} catch (error) {
  console.error(error);
  process.exit(1);
}

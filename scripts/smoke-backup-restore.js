// Run with: node node_modules/electron/cli.js scripts/smoke-backup-restore.js
// On Windows, use the CLI wrapper so the caller waits for Electron to exit and
// keeps its output pipes open. A direct GUI executable launch may detach.
// Exercise the real main/preload backup IPC with an isolated profile and native
// dialogs stubbed. No user database, user files, or external network are used.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { createHash } = require('node:crypto');
const Database = require('better-sqlite3');
const { app, BrowserWindow, dialog, nativeImage, session } = require('electron');
const { initDatabase } = require('../src/db');

const SMOKE_ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-backup-restore-'));
const PROFILE_DIR = path.join(SMOKE_ROOT, 'Profile');
const APPDATA_DIR = path.join(PROFILE_DIR, 'AppData', 'Roaming');
const LOCALAPPDATA_DIR = path.join(PROFILE_DIR, 'AppData', 'Local');
const TEMP_DIR = path.join(SMOKE_ROOT, 'Temp');
const OUTPUT_DIR = path.resolve(__dirname, '..', 'out', 'smoke-backup-restore', path.basename(SMOKE_ROOT));
const DATA_DIR = path.join(APPDATA_DIR, 'WorldShot Log', 'data');
const DB_PATH = path.join(DATA_DIR, 'app.sqlite');
const PREFERENCES_PATH = path.join(DATA_DIR, 'preferences.json');
const BACKUP_DIR = path.join(DATA_DIR, 'backups');
const FIXTURE_DIR = path.join(SMOKE_ROOT, 'Fixtures');
const TABLE_ORDER = {
  photos: 'id',
  tracked_folders: 'id',
  world_metadata_cache: 'world_id',
  tags: 'id',
  photo_tags: 'photo_id, tag_id',
};
const report = { ok: false, smokeRoot: SMOKE_ROOT, checks: [], blockedRequests: [] };
let nextOpenDialog = null;
let nextSavePath = null;
let win;
let watchdog = null;
let finished = false;
let loggingEnabled = true;

function finish(exitCode, error = null) {
  if (finished) return;
  finished = true;
  clearTimeout(watchdog);
  if (error) {
    report.ok = false;
    report.error = error.stack || String(error);
    writeLog('stderr', report.error);
  }
  const reportPath = path.join(OUTPUT_DIR, 'result.json');
  try {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
    writeLog('stdout', `Backup restore report: ${reportPath}`);
  } catch (reportError) {
    exitCode = 1;
    writeLog('stderr', `Could not write the test report: ${reportError.stack || reportError}`);
  }
  app.exit(exitCode);
}

function handleLogStreamError(error) {
  if (error?.code === 'EPIPE') {
    // A closed parent output pipe must not open Electron's native exception
    // dialog. Disable console output; the artifact still records test failures.
    loggingEnabled = false;
    return;
  }
  finish(1, error);
}

function writeLog(streamName, message) {
  if (!loggingEnabled) return;
  try {
    process[streamName].write(`${message}\n`);
  } catch (error) {
    handleLogStreamError(error);
  }
}

process.stdout.on('error', handleLogStreamError);
process.stderr.on('error', handleLogStreamError);
// Real asynchronous test failures remain failures, including EPIPE errors from
// anything other than stdout/stderr. Do not let Electron show a native dialog.
process.on('uncaughtException', error => finish(1, error));
process.on('unhandledRejection', reason => finish(1,
  reason instanceof Error ? reason : new Error(String(reason))));

for (const directory of [PROFILE_DIR, APPDATA_DIR, LOCALAPPDATA_DIR, TEMP_DIR, DATA_DIR, FIXTURE_DIR, OUTPUT_DIR]) {
  fs.mkdirSync(directory, { recursive: true });
}
process.env.APPDATA = process.env.WORLDSHOT_APPDATA_ROOT = APPDATA_DIR;
process.env.LOCALAPPDATA = process.env.WORLDSHOT_LOCALAPPDATA_ROOT = LOCALAPPDATA_DIR;
process.env.USERPROFILE = PROFILE_DIR;
process.env.WORLDSHOT_HOME_ROOT = PROFILE_DIR;
process.env.TEMP = process.env.TMP = TEMP_DIR;
app.setPath('home', PROFILE_DIR);
app.setPath('appData', APPDATA_DIR);
app.setPath('userData', path.dirname(DATA_DIR));
app.setPath('temp', TEMP_DIR);
for (const key of ['pictures', 'documents', 'downloads', 'desktop']) {
  const directory = path.join(PROFILE_DIR, key);
  fs.mkdirSync(directory, { recursive: true });
  app.setPath(key, directory);
}

function assertIsolated(filePath) {
  const relative = path.relative(SMOKE_ROOT, filePath);
  assert(relative && !relative.startsWith('..') && !path.isAbsolute(relative), filePath);
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function seedDatabase(dbPath, prefix, count) {
  assertIsolated(dbPath);
  const db = initDatabase(dbPath);
  try {
    const timestamp = '2026-06-08T00:10:00.000Z';
    const worldId = prefix === 'original'
      ? 'wrld_11111111-1111-4111-8111-111111111111'
      : 'wrld_22222222-2222-4222-8222-222222222222';
    for (let index = 0; index < count; index += 1) {
      const filePath = path.join(FIXTURE_DIR, `${prefix}-${index}.png`);
      const pixels = Buffer.alloc(16 * 12 * 4, index + 100);
      for (let offset = 3; offset < pixels.length; offset += 4) pixels[offset] = 255;
      const bytes = nativeImage.createFromBitmap(pixels, { width: 16, height: 12 }).toPNG();
      fs.writeFileSync(filePath, bytes);
      const photo = db.insertOrUpdatePhoto({
        filePath, fileName: path.basename(filePath),
        fileHash: createHash('sha256').update(bytes).digest('hex'),
        takenAt: `2026/06/08 00:10:0${index}`,
        takenAtTimestamp: Date.parse(timestamp) + index * 1000,
        groupDate: '2026-06-08', year: 2026, month: 6, day: 8,
        worldId, worldName: `${prefix} world`, worldNameManual: `${prefix} manual world`,
        worldUrl: `https://vrchat.com/home/world/${worldId}`,
        thumbnailPath: filePath, imageWidth: 16, imageHeight: 12,
        resolutionTier: 'other', orientationTier: 'landscape',
        printNoteText: `${prefix} print note`, memoText: `${prefix} memo ${index}`,
        createdAt: timestamp, updatedAt: timestamp,
      });
      db.updateFavoriteStatus(photo.id, true);
      db.replacePhotoTags(photo.id, [{ name: `${prefix} label`, colorHex: '#14B8A6' }]);
    }
    const trackedPath = path.join(FIXTURE_DIR, `${prefix}-tracked`);
    fs.mkdirSync(trackedPath, { recursive: true });
    db.upsertTrackedFolder(trackedPath);
    db.upsertWorldMetadata({
      worldId, sourceUrl: `https://vrchat.com/home/world/${worldId}`,
      worldNameOfficial: `${prefix} official`, worldDescription: `${prefix} description`,
      worldTags: ['backup', prefix], authorId: 'usr_smoke', authorName: 'Backup smoke',
      releaseStatus: 'public', fetchStatus: 'success',
      fetchedAt: timestamp, lastAttemptedAt: timestamp,
    });
    return db.createBackupSnapshot();
  } finally {
    db.close();
  }
}

function readDatabase() {
  const db = new Database(DB_PATH, { readonly: true, fileMustExist: true });
  try {
    return {
      tables: Object.fromEntries(Object.entries(TABLE_ORDER).map(([table, order]) =>
        [table, db.prepare(`SELECT * FROM ${table} ORDER BY ${order}`).all()])),
      sequences: db.prepare('SELECT * FROM sqlite_sequence ORDER BY name').all(),
    };
  } finally {
    db.close();
  }
}

function readAutomaticBackups() {
  if (!fs.existsSync(BACKUP_DIR)) return {};
  return Object.fromEntries(fs.readdirSync(BACKUP_DIR).sort().map(name =>
    [name, fs.readFileSync(path.join(BACKUP_DIR, name)).toString('base64')]));
}

function capturePersistedState() {
  return {
    database: readDatabase(),
    preferences: fs.readFileSync(PREFERENCES_PATH).toString('base64'),
    automaticBackups: readAutomaticBackups(),
  };
}

async function until(label, predicate) {
  const started = Date.now();
  while (Date.now() - started < 30000) {
    if (await predicate()) return;
    await new Promise(resolve => setTimeout(resolve, 60));
  }
  throw new Error(`Timed out: ${label}`);
}

async function check(name, run) {
  try {
    await run();
    report.checks.push({ name, ok: true });
    writeLog('stdout', `PASS ${name}`);
  } catch (error) {
    report.checks.push({ name, ok: false, error: error.message });
    throw error;
  }
}

async function restoreText(name, text) {
  assert.equal(nextOpenDialog, null);
  const filePath = path.join(FIXTURE_DIR, `${name}.json`);
  fs.writeFileSync(filePath, text, 'utf8');
  nextOpenDialog = { canceled: false, filePaths: [filePath] };
  const result = await win.webContents.executeJavaScript('window.electronAPI.restoreAppDataBackup()');
  assert.equal(nextOpenDialog, null, 'Restore IPC did not open the expected native dialog');
  return result;
}

async function rejectWithoutChanges(name, payload, raw = false) {
  await check(name, async () => {
    const before = capturePersistedState();
    const result = await restoreText(name, raw ? payload : JSON.stringify(payload));
    assert.equal(result.ok, false, `Invalid backup accepted: ${JSON.stringify(result)}`);
    assert.equal(typeof result.message, 'string');
    assert.ok(result.message.length > 0);
    assert.deepEqual(capturePersistedState(), before,
      'Rejected restore must preserve all rows, sequences, preference bytes, and automatic backup files');
  });
}

async function restoreValid(name, payload, expectedSnapshot, expectedPreferences = null) {
  await check(name, async () => {
    const before = capturePersistedState();
    const beforePreferences = JSON.parse(Buffer.from(before.preferences, 'base64').toString('utf8'));
    const result = await restoreText(name, JSON.stringify(payload));
    assert.equal(result.ok, true, JSON.stringify(result));
    assert.equal(result.canceled, false);
    assert.deepEqual(readDatabase().tables, expectedSnapshot.tables);
    assert.equal(result.restoredPhotoCount, expectedSnapshot.tables.photos.length);
    assertIsolated(result.automaticBackupFilePath);
    assert.equal(path.dirname(result.automaticBackupFilePath), BACKUP_DIR);
    const automaticBackup = JSON.parse(fs.readFileSync(result.automaticBackupFilePath, 'utf8'));
    assert.equal(automaticBackup.appName, 'WorldShot Log');
    assert.equal(automaticBackup.backupSchemaVersion, 1);
    assert.deepEqual(automaticBackup.data.tables, before.database.tables,
      'Automatic backup must contain the pre-restore rows, including labels and metadata');
    assert.deepEqual(automaticBackup.preferences, beforePreferences,
      'Automatic backup must preserve the pre-restore preferences');
    if (expectedPreferences) {
      const restoredPreferences = JSON.parse(fs.readFileSync(PREFERENCES_PATH, 'utf8'));
      assert.equal(restoredPreferences.backgroundImagePath, expectedPreferences.backgroundImagePath);
      assert.equal(restoredPreferences.smokeMarker, expectedPreferences.smokeMarker);
    } else {
      assert.equal(fs.readFileSync(PREFERENCES_PATH).toString('base64'), before.preferences,
        'Snapshot without preferences must preserve the existing preference bytes');
    }
  });
}

async function main() {
  await app.whenReady();
  session.defaultSession.webRequest.onBeforeRequest(
    { urls: ['http://*/*', 'https://*/*', 'ws://*/*', 'wss://*/*'] },
    (details, callback) => {
      report.blockedRequests.push(details.url);
      callback({ cancel: true });
    }
  );
  globalThis.fetch = async (input) => {
    report.blockedRequests.push(String(input));
    throw new Error('External network is disabled in backup restore smoke');
  };
  dialog.showOpenDialog = async () => {
    assert(nextOpenDialog, 'Unexpected open dialog');
    const result = nextOpenDialog;
    for (const filePath of result.filePaths) assertIsolated(filePath);
    nextOpenDialog = null;
    return result;
  };
  dialog.showSaveDialog = async () => {
    assert(nextSavePath, 'Unexpected save dialog');
    assertIsolated(nextSavePath);
    const filePath = nextSavePath;
    nextSavePath = null;
    return { canceled: false, filePath };
  };
  app.on('browser-window-created', (_event, window) => window.hide());
  seedDatabase(DB_PATH, 'original', 1);
  const target = seedDatabase(path.join(FIXTURE_DIR, 'target.sqlite'), 'restored', 2);
  const originalPreferences = {
    backgroundImagePath: path.join(FIXTURE_DIR, 'original-0.png'),
    pendingInstalledAppUpdateNotice: null,
    lastShownInstalledAppUpdateNoticeVersion: '',
    smokeMarker: 'original preferences',
  };
  fs.writeFileSync(PREFERENCES_PATH, JSON.stringify(originalPreferences, null, 2));
  require('../src/index');
  for (const key of ['home', 'appData', 'userData', 'sessionData', 'temp']) assertIsolated(app.getPath(key));
  await until('application window', () => {
    win = BrowserWindow.getAllWindows()[0];
    return Boolean(win);
  });
  await until('renderer startup', () => win.webContents.executeJavaScript(
    `Boolean(window.electronAPI?.restoreAppDataBackup) &&
     !document.querySelector('.app')?.classList.contains('is-app-initializing')`
  ));

  let originalBackup;
  await check('real create-backup IPC writes the official envelope', async () => {
    const before = capturePersistedState();
    const backupPath = path.join(FIXTURE_DIR, 'original-export.json');
    nextSavePath = backupPath;
    const result = await win.webContents.executeJavaScript('window.electronAPI.createAppDataBackup()');
    assert.equal(nextSavePath, null);
    assert.equal(result.ok, true, JSON.stringify(result));
    originalBackup = JSON.parse(fs.readFileSync(backupPath, 'utf8'));
    assert.equal(originalBackup.appName, 'WorldShot Log');
    assert.equal(originalBackup.backupSchemaVersion, 1);
    assert.equal(originalBackup.data.schemaVersion, 1);
    assert.deepEqual(originalBackup.data.tables, before.database.tables);
    assert.deepEqual(capturePersistedState(), before);
  });
  const targetPreferences = { ...originalPreferences, smokeMarker: 'restored preferences' };
  const targetBackup = { ...clone(originalBackup), data: target, preferences: targetPreferences };
  await check('cancelled restore preserves all persisted state', async () => {
    const before = capturePersistedState();
    nextOpenDialog = { canceled: true, filePaths: [] };
    const result = await win.webContents.executeJavaScript('window.electronAPI.restoreAppDataBackup()');
    assert.equal(nextOpenDialog, null);
    assert.equal(result.ok, true);
    assert.equal(result.canceled, true);
    assert.deepEqual(capturePersistedState(), before);
  });
  await rejectWithoutChanges('malformed-json', '{"unfinished":', true);
  await rejectWithoutChanges('null-root', null);
  await rejectWithoutChanges('array-root', []);
  await rejectWithoutChanges('empty-object', {});
  await rejectWithoutChanges('catalog-is-not-backup', { appName: 'WorldShot Log', photos: [] });
  await rejectWithoutChanges('future-envelope-version', { ...targetBackup, backupSchemaVersion: 2 });
  await rejectWithoutChanges('wrong-application', { ...targetBackup, appName: 'Another application' });
  await rejectWithoutChanges('future-data-version', {
    ...targetBackup, data: { ...target, schemaVersion: 2 },
  });
  const missingTable = clone(targetBackup);
  delete missingTable.data.tables.photo_tags;
  await rejectWithoutChanges('missing-table', missingTable);
  const wrongTable = clone(targetBackup);
  wrongTable.data.tables.photos = {};
  await rejectWithoutChanges('table-is-not-array', wrongTable);
  const nullRow = clone(targetBackup);
  nullRow.data.tables.photos.push(null);
  await rejectWithoutChanges('null-row', nullRow);
  const missingRequired = clone(targetBackup);
  delete missingRequired.data.tables.photos[0].file_path;
  await rejectWithoutChanges('missing-required-column', missingRequired);
  const invalidValue = clone(targetBackup);
  invalidValue.data.tables.photos[0].memo_text = ['invalid column value'];
  await rejectWithoutChanges('invalid-row-value', invalidValue);
  const duplicatePhoto = clone(targetBackup);
  duplicatePhoto.data.tables.photos.push(clone(duplicatePhoto.data.tables.photos[0]));
  await rejectWithoutChanges('duplicate-photo-key', duplicatePhoto);
  const danglingTag = clone(targetBackup);
  danglingTag.data.tables.photo_tags[0].photo_id = 999999;
  await rejectWithoutChanges('dangling-photo-tag', danglingTag);
  await rejectWithoutChanges('preferences-array', { ...targetBackup, preferences: [] });

  await restoreValid('official-envelope', targetBackup, target, targetPreferences);
  await restoreValid('bare-snapshot', originalBackup.data, originalBackup.data);
  await restoreValid('data-wrapper', { data: target }, target);
  const empty = { schemaVersion: 1, tables: Object.fromEntries(Object.keys(TABLE_ORDER).map(name => [name, []])) };
  await restoreValid('valid-empty-backup', { ...targetBackup, data: empty }, empty, targetPreferences);
  await restoreValid('restore-real-export-after-empty', originalBackup, originalBackup.data, originalPreferences);
  report.ok = true;
}

watchdog = setTimeout(() => finish(1,
  new Error('Backup restore smoke exceeded 120 seconds')), 120000);
main().then(() => finish(report.ok ? 0 : 1), error => finish(1, error));

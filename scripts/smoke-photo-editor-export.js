// Run with: electron scripts/smoke-photo-editor-export.js
// Repeat with --legacy-world to verify imported legacy identity inheritance.
// Exercises the real renderer export and save IPC, with only the native dialog stubbed.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { createHash } = require('node:crypto');
const Database = require('better-sqlite3');
const { app, BrowserWindow, dialog, nativeImage, session } = require('electron');

const SMOKE_ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-export-'));
const RUN_ID = path.basename(SMOKE_ROOT);
const OUTPUT_DIR = path.resolve(__dirname, '..', 'out', 'smoke-export', RUN_ID);
const PROFILE_DIR = path.join(SMOKE_ROOT, 'Profile');
const APPDATA_DIR = path.join(PROFILE_DIR, 'AppData', 'Roaming');
const LOCALAPPDATA_DIR = path.join(PROFILE_DIR, 'AppData', 'Local');
const TEMP_DIR = path.join(SMOKE_ROOT, 'Temp');
const DB_PATH = path.join(APPDATA_DIR, 'WorldShot Log', 'data', 'app.sqlite');
const SOURCE_PATH = path.join(SMOKE_ROOT, 'VRChat_2026-06-07_22-20-00.000_320x180.png');
const WORLD_ID = 'wrld_bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';
const WORLD_URL = `https://vrchat.com/home/world/${WORLD_ID}`;
const WORLD_NAME = 'Export smoke B';
const WORLD_IDENTITY_SOURCE = process.argv.includes('--legacy-world') ? 'legacy' : 'manual';
const WIDTH = 320;
const HEIGHT = 180;
const SOURCE_COLORS = [[32, 80, 160, 255], [200, 96, 40, 255]];
const EXPORT_FORMATS = {
  png: { mimeType: 'image/png', extension: 'png' },
  jpeg: { mimeType: 'image/jpeg', extension: 'jpg' },
  webp: { mimeType: 'image/webp', extension: 'webp' },
};
const blockedRequests = [];
const savedFiles = [];
let nextSavePath = null;

for (const dir of [PROFILE_DIR, APPDATA_DIR, LOCALAPPDATA_DIR, TEMP_DIR, OUTPUT_DIR]) {
  fs.mkdirSync(dir, { recursive: true });
}
process.env.APPDATA = process.env.WORLDSHOT_APPDATA_ROOT = APPDATA_DIR;
process.env.LOCALAPPDATA = process.env.WORLDSHOT_LOCALAPPDATA_ROOT = LOCALAPPDATA_DIR;
process.env.USERPROFILE = PROFILE_DIR;
process.env.WORLDSHOT_HOME_ROOT = PROFILE_DIR;
process.env.TEMP = process.env.TMP = TEMP_DIR;
app.setPath('home', PROFILE_DIR);
app.setPath('appData', APPDATA_DIR);
app.setPath('userData', path.join(APPDATA_DIR, 'WorldShot Log'));
app.setPath('temp', TEMP_DIR);
for (const key of ['pictures', 'documents', 'downloads', 'desktop']) {
  const dir = path.join(PROFILE_DIR, key);
  fs.mkdirSync(dir, { recursive: true });
  app.setPath(key, dir);
}

function hash(buffer) {
  return createHash('sha256').update(buffer).digest('hex');
}

function assertIsolated(filePath) {
  const relative = path.relative(SMOKE_ROOT, filePath);
  assert(relative && !relative.startsWith('..') && !path.isAbsolute(relative), filePath);
}

function seedPhoto() {
  // Electron's Windows bitmap layout is BGRA; all input pixels are opaque.
  const bitmap = Buffer.alloc(WIDTH * HEIGHT * 4);
  for (let y = 0; y < HEIGHT; y += 1) {
    for (let x = 0; x < WIDTH; x += 1) {
      const [r, g, b, a] = SOURCE_COLORS[x < WIDTH / 2 ? 0 : 1];
      bitmap.set([b, g, r, a], (y * WIDTH + x) * 4);
    }
  }
  fs.writeFileSync(SOURCE_PATH, nativeImage.createFromBitmap(bitmap, {
    width: WIDTH, height: HEIGHT, scaleFactor: 1,
  }).toPNG());
  const dataDir = path.join(app.getPath('userData'), 'data');
  fs.mkdirSync(dataDir, { recursive: true });
  const db = require('../src/db').initDatabase(DB_PATH);
  try {
    // Cached metadata lets the real sync queue run deterministically without a network call.
    db.upsertWorldMetadata({ worldId: WORLD_ID, sourceUrl: WORLD_URL,
      worldNameOfficial: WORLD_NAME, fetchStatus: 'success' });
    return db.insertOrUpdatePhoto({
      filePath: SOURCE_PATH, fileName: path.basename(SOURCE_PATH),
      fileHash: hash(fs.readFileSync(SOURCE_PATH)),
      takenAt: '2026/06/07 22:20:00',
      takenAtTimestamp: new Date(2026, 5, 7, 22, 20).getTime(),
      groupDate: '2026-06-07', year: 2026, month: 6, day: 7,
      worldId: WORLD_ID, worldName: WORLD_NAME, worldNameManual: null, worldUrl: WORLD_URL,
      worldIdentitySource: WORLD_IDENTITY_SOURCE,
      thumbnailPath: SOURCE_PATH, imageWidth: WIDTH, imageHeight: HEIGHT,
      resolutionTier: 'other', orientationTier: 'landscape',
      printNoteText: '', memoText: 'export fixture',
      createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
    }).id;
  } finally {
    db.close();
  }
}

function assertWorldIdentity(photo) {
  assert.equal(photo.worldId, WORLD_ID);
  assert.equal(photo.worldUrl, WORLD_URL);
  assert.equal(photo.worldName, WORLD_NAME);
  assert.equal(photo.rawWorldName, WORLD_NAME);
  assert.equal(photo.worldNameManual, null);
  assert.equal(photo.worldIdentitySource, WORLD_IDENTITY_SOURCE);
}

function assertStoredWorldIdentity(photoId) {
  const db = new Database(DB_PATH, { readonly: true });
  try {
    const row = db.prepare(`SELECT world_id, world_url, world_name, world_name_manual,
      world_identity_source FROM photos WHERE id = ?`).get(photoId);
    assert.deepEqual(row, { world_id: WORLD_ID, world_url: WORLD_URL, world_name: WORLD_NAME,
      world_name_manual: null, world_identity_source: WORLD_IDENTITY_SOURCE });
  } finally {
    db.close();
  }
}

function inspectStoredImage(photoId) {
  const db = new Database(DB_PATH, { readonly: true });
  try {
    const row = db.prepare(`SELECT image_width, image_height, thumbnail_path
      FROM photos WHERE id = ?`).get(photoId);
    assert.equal(row.image_width, WIDTH);
    assert.equal(row.image_height, HEIGHT);
    assert.equal(typeof row.thumbnail_path, 'string');
    assertIsolated(row.thumbnail_path);
    const thumbnail = nativeImage.createFromPath(row.thumbnail_path);
    assert.equal(thumbnail.isEmpty(), false, 'Saved image must have a decodable managed thumbnail');
    const thumbnailSize = thumbnail.getSize();
    assert(thumbnailSize.width > 0 && thumbnailSize.height > 0);
    return { thumbnailPath: row.thumbnail_path, thumbnailSize };
  } finally {
    db.close();
  }
}

async function until(label, predicate, timeoutMs = 30000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const result = await predicate();
    if (result) return result;
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  throw new Error(`Timed out: ${label}`);
}

function runInPage(win, source) {
  return win.webContents.executeJavaScript(source);
}

async function decodeSavedImage(win, format, filePath, fileBytes) {
  if (format !== 'webp') {
    const decoded = nativeImage.createFromPath(filePath);
    assert.deepEqual(decoded.getSize(), { width: WIDTH, height: HEIGHT });
    return decoded.toBitmap();
  }
  // Windows nativeImage does not decode WebP; decode the actual saved bytes in
  // Chromium, the same image engine that displays these images in the app.
  const decoded = await runInPage(win, `(async () => {
    const bytes = Uint8Array.from(atob(${JSON.stringify(fileBytes.toString('base64'))}), char => char.charCodeAt(0));
    const bitmap = await createImageBitmap(new Blob([bytes], { type: 'image/webp' }));
    try {
      const canvas = document.createElement('canvas');
      canvas.width = bitmap.width; canvas.height = bitmap.height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(bitmap, 0, 0);
      return { width: bitmap.width, height: bitmap.height,
        rgba: Array.from(ctx.getImageData(0, 0, canvas.width, canvas.height).data) };
    } finally { bitmap.close(); }
  })()`);
  assert.equal(decoded.width, WIDTH);
  assert.equal(decoded.height, HEIGHT);
  const bgra = Buffer.from(decoded.rgba);
  for (let offset = 0; offset < bgra.length; offset += 4) {
    const r = bgra[offset];
    bgra[offset] = bgra[offset + 2];
    bgra[offset + 2] = r;
  }
  return bgra;
}

async function importMask(win) {
  await runInPage(win, `new Promise((resolve, reject) => {
    const canvas = document.createElement('canvas');
    canvas.width = ${WIDTH}; canvas.height = ${HEIGHT};
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#000'; ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width / 2, canvas.height);
    canvas.toBlob(blob => {
      try {
        const transfer = new DataTransfer();
        transfer.items.add(new File([blob], 'half-mask.png', { type: 'image/png' }));
        const input = document.getElementById('photo-editor-subject-file');
        input.files = transfer.files;
        input.dispatchEvent(new Event('change', { bubbles: true }));
        resolve(true);
      } catch (error) { reject(error); }
    }, 'image/png');
  })`);
  await until('imported mask', () => runInPage(win,
    `photoEditorState?.subjectMask?.source === 'imported' &&
     photoEditorState.subjectMask.status === 'ready'`));
  await runInPage(win, '(async () => { await waitForPhotoEditorSubjectMaskToLoad(); return true; })()');
}

async function exportAndSave(win, name, transparent, { format = 'png', quality = 92,
  expectedRoute = null } = {}) {
  const formatMeta = EXPORT_FORMATS[format];
  assert(formatMeta, `Unsupported smoke format: ${format}`);
  assert.equal(nextSavePath, null);
  nextSavePath = path.join(SMOKE_ROOT, 'Saved', `${name}.${formatMeta.extension}`);
  const expectedPath = nextSavePath;
  const { result, renderDetails, sourceMetadata, workerCalls } = await runInPage(win, `(async () => {
    commitPhotoEditorHistoryMutation();
    const settings = normalizePhotoEditorExportSettings(${JSON.stringify({ format, quality, maxEdge: 0 })});
    const workerCallOffset = window.__exportSmokeWorkerCalls.length;
    const rendered = await renderPhotoEditorExportDataUrl(
      document.createElement('canvas'), settings,
      { transparentSubjectBackground: ${transparent} }
    );
    const source = photoEditorState.sourcePhoto;
    const sourceMetadata = buildPhotoEditorSourceMetadataPayload(source);
    const result = await window.electronAPI.saveEditedPhoto({
      sourceFilePath: source.filePath, sourceFileName: source.fileName,
      sourceTakenAtTimestamp: source.takenAtTimestamp,
      sourcePhotoMetadata: sourceMetadata,
      outputFormat: settings.format, dataUrl: rendered.dataUrl,
    });
    return { result, sourceMetadata,
      renderDetails: { mimeType: rendered.dataUrl.slice(5, rendered.dataUrl.indexOf(';')),
        width: rendered.outputSize.width, height: rendered.outputSize.height, settings },
      workerCalls: window.__exportSmokeWorkerCalls.slice(workerCallOffset) };
  })()`);
  assert.equal(nextSavePath, null, 'Expected the real save IPC to open the stubbed dialog');
  assert.equal(result.ok, true, JSON.stringify(result));
  assert.equal(result.canceled, false);
  assert.equal(result.importFailed, false, result.importMessage);
  assert.equal(result.filePath, expectedPath);
  assert.deepEqual(renderDetails, { mimeType: formatMeta.mimeType, width: WIDTH, height: HEIGHT,
    settings: { format, quality, maxEdge: 0 } });
  assert.equal(result.photo.memoText, 'export fixture');
  for (const key of ['takenAt', 'takenAtTimestamp', 'groupDate', 'year', 'month', 'day']) {
    assert.equal(result.photo[key], sourceMetadata[key], `${name}: source ${key} must survive save/import`);
  }
  assert.equal(result.photo.imageWidth, WIDTH);
  assert.equal(result.photo.imageHeight, HEIGHT);
  assert(Math.abs(fs.statSync(result.filePath).mtimeMs - sourceMetadata.takenAtTimestamp) <= 1,
    `${name}: saved file timestamp must match the source capture time`);
  assertWorldIdentity(result.photo);
  assertStoredWorldIdentity(result.photo.id);
  assert.deepEqual(result.importResult.worldMetadataTargets, [{ worldId: WORLD_ID, worldUrl: WORLD_URL }],
    'Save must synchronize B after copying source metadata, not the embedded/empty identity');
  assertIsolated(result.filePath);
  const fileBytes = fs.readFileSync(result.filePath);
  if (format === 'png') {
    assert.equal(fileBytes.subarray(0, 8).toString('hex'), '89504e470d0a1a0a');
  } else if (format === 'jpeg') {
    assert.equal(fileBytes.subarray(0, 3).toString('hex'), 'ffd8ff');
    assert.equal(fileBytes.subarray(-2).toString('hex'), 'ffd9');
  } else {
    assert.equal(fileBytes.subarray(0, 4).toString('ascii'), 'RIFF');
    assert.equal(fileBytes.subarray(8, 12).toString('ascii'), 'WEBP');
    assert.equal(fileBytes.readUInt32LE(4) + 8, fileBytes.length);
  }
  const bitmap = await decodeSavedImage(win, format, result.filePath, fileBytes);
  assert.equal(bitmap.length, WIDTH * HEIGHT * 4);
  if (expectedRoute === 'worker') {
    assert.equal(workerCalls.length, 1, `${name}: must exercise the real Worker export`);
    assert.deepEqual(workerCalls[0], { mimeType: formatMeta.mimeType, quality: quality / 100,
      responseType: 'blob', returnedMimeType: formatMeta.mimeType, returnedBuffer: true });
  } else if (expectedRoute === 'main') {
    assert.deepEqual(workerCalls, [], `${name}: subject mask must use the main renderer path`);
  }
  savedFiles.push({ name, path: result.filePath, sha256: hash(fileBytes), pixelSha256: hash(bitmap),
    format, quality, byteLength: fileBytes.length, renderDetails,
    ...inspectStoredImage(result.photo.id),
    ...(expectedRoute ? { renderRoute: expectedRoute, workerCalls } : {}),
    photoId: result.photo.id, worldIdentitySource: result.photo.worldIdentitySource,
    worldMetadataTargets: result.importResult.worldMetadataTargets });
  return bitmap;
}

async function verifyReimportAndSync(win, sourceHash) {
  const result = await runInPage(win,
    'window.electronAPI.reimportRegisteredPhotos({ year: 2026, month: 6 })');
  assert.equal(result.ok, true, JSON.stringify(result));
  assert.equal(result.failedCount, 0);
  const uniqueImageCount = new Set([sourceHash, ...savedFiles.map(file => file.sha256)]).size;
  assert.equal(result.importedCount, uniqueImageCount);
  assert.deepEqual(result.worldMetadataTargets, [{ worldId: WORLD_ID, worldUrl: WORLD_URL }],
    'Reimport must synchronize the preserved B identity, even though these image files contain no world ID');
  const photos = await runInPage(win, 'window.electronAPI.getPhotosByMonth(2026, 6)');
  assert.equal(photos.length, uniqueImageCount);
  for (const photo of photos) {
    assertWorldIdentity(photo);
    assertStoredWorldIdentity(photo.id);
    assert.equal(photo.imageWidth, WIDTH, 'Image width must survive month reimport');
    assert.equal(photo.imageHeight, HEIGHT, 'Image height must survive month reimport');
  }

  const sync = await runInPage(win, `(async () => {
    let unsubscribe;
    let timer;
    const update = new Promise((resolve, reject) => {
      timer = setTimeout(() => {
        unsubscribe(); reject(new Error('World B sync did not emit an update'));
      }, 10000);
      unsubscribe = window.electronAPI.onWorldMetadataUpdated(resolve);
    });
    try {
      const queued = await window.electronAPI.startWorldMetadataSync(${JSON.stringify(result.worldMetadataTargets)});
      return { queued, update: await update };
    } finally {
      clearTimeout(timer); unsubscribe();
    }
  })()`);
  assert.equal(sync.queued.ok, true);
  assert.equal(sync.queued.queuedCount, 1);
  assert.equal(sync.update.photos.length, uniqueImageCount);
  for (const photo of sync.update.photos) {
    assertWorldIdentity(photo);
    assertStoredWorldIdentity(photo.id);
  }
  return { reimportedCount: result.importedCount, targets: result.worldMetadataTargets,
    synchronizedPhotoIds: sync.update.photos.map(photo => photo.id) };
}

function assertPixels(bitmap, visibleHalf = 'both') {
  for (let y = 0; y < HEIGHT; y += 1) {
    for (let x = 0; x < WIDTH; x += 1) {
      const isLeft = x < WIDTH / 2;
      const visible = visibleHalf === 'both' || (visibleHalf === 'left') === isLeft;
      const [r, g, b] = SOURCE_COLORS[isLeft ? 1 : 0]; // Source is flipped horizontally.
      const offset = (y * WIDTH + x) * 4;
      const expected = visible ? [b, g, r, 255] : [0, 0, 0, 0];
      for (let channel = 0; channel < 4; channel += 1) {
        assert.equal(bitmap[offset + channel], expected[channel],
          `${visibleHalf}: pixel (${x}, ${y}), BGRA channel ${channel}`);
      }
    }
  }
}

function assertLossyPixels(bitmap, name) {
  let absoluteError = 0;
  let maxInteriorError = 0;
  // JPEG/WebP chroma subsampling can blend colors at the two-color boundary.
  // Check every alpha and overall RGB error, and check the flat interiors tightly.
  for (let y = 0; y < HEIGHT; y += 1) {
    for (let x = 0; x < WIDTH; x += 1) {
      const [r, g, b] = SOURCE_COLORS[x < WIDTH / 2 ? 1 : 0];
      const expected = [b, g, r];
      const offset = (y * WIDTH + x) * 4;
      assert.equal(bitmap[offset + 3], 255, `${name}: all pixels must remain opaque`);
      for (let channel = 0; channel < 3; channel += 1) {
        const error = Math.abs(bitmap[offset + channel] - expected[channel]);
        absoluteError += error;
        if (Math.abs(x - WIDTH / 2) > 8) maxInteriorError = Math.max(maxInteriorError, error);
      }
    }
  }
  const meanAbsoluteError = absoluteError / (WIDTH * HEIGHT * 3);
  assert(meanAbsoluteError <= 4, `${name}: mean RGB error ${meanAbsoluteError} exceeds 4/255`);
  assert(maxInteriorError <= 8, `${name}: flat interior error ${maxInteriorError} exceeds 8/255`);
  return { meanAbsoluteError, maxInteriorError };
}

async function verifyLossyFormats(win, route) {
  for (const format of ['jpeg', 'webp']) {
    const records = [];
    for (const quality of [60, 100]) {
      const name = `${format}-${route}-quality-${quality}`;
      const bitmap = await exportAndSave(win, name, false, { format, quality, expectedRoute: route });
      const record = savedFiles.at(-1);
      record.pixelError = assertLossyPixels(bitmap, name);
      records.push(record);
    }
    assert.notEqual(records[0].sha256, records[1].sha256,
      `${format}/${route}: changing quality must change the encoded image`);
  }
}

async function main() {
  await app.whenReady();
  session.defaultSession.webRequest.onBeforeRequest(
    { urls: ['http://*/*', 'https://*/*', 'ws://*/*', 'wss://*/*'] },
    (details, callback) => {
      blockedRequests.push(details.url);
      callback({ cancel: true });
    }
  );
  // World metadata uses Node fetch, outside Chromium's session request hooks.
  globalThis.fetch = async (input) => {
    blockedRequests.push(String(input));
    throw new Error('External network is disabled in export smoke');
  };
  dialog.showSaveDialog = async () => {
    assert(nextSavePath, 'Unexpected save dialog');
    assertIsolated(nextSavePath);
    const filePath = nextSavePath;
    nextSavePath = null;
    return { canceled: false, filePath };
  };
  const photoId = seedPhoto();
  const sourceHash = hash(fs.readFileSync(SOURCE_PATH));
  require('../src/index');
  for (const key of ['home', 'appData', 'userData', 'sessionData', 'temp']) {
    assertIsolated(app.getPath(key));
  }
  const win = await until('application window', () => BrowserWindow.getAllWindows()[0]);
  win.webContents.on('render-process-gone', (_event, details) => {
    console.error('Renderer exited:', details);
    app.exit(1);
  });
  await until('source photo card', () => runInPage(win,
    `Boolean(document.querySelector('.photo-card[data-photo-id="${photoId}"]'))`));
  await runInPage(win, `document.querySelector('.photo-card[data-photo-id="${photoId}"]').click()`);
  await runInPage(win, 'openPhotoEditorModal()');
  await until('editor image', () => runInPage(win, 'Boolean(photoEditorState?.sourceImage)'));
  await runInPage(win, `(() => {
    window.__exportSmokeWorkerCalls = [];
    const original = applyPhotoEditorEffectsWithWorker;
    applyPhotoEditorEffectsWithWorker = async function (...args) {
      if (args[2]?.purpose !== 'export') return original.apply(this, args);
      const call = { mimeType: args[2].exportSettings.mimeType,
        quality: args[2].exportSettings.quality, responseType: args[2].responseType };
      window.__exportSmokeWorkerCalls.push(call);
      const result = await original.apply(this, args);
      call.returnedMimeType = result?.mimeType;
      call.returnedBuffer = Boolean(result?.buffer?.byteLength);
      return result;
    };
  })()`);
  assertWorldIdentity(await runInPage(win, 'photoEditorState.sourcePhoto'));
  assertStoredWorldIdentity(photoId);
  await runInPage(win, `updatePhotoEditorCrop({ flipX: true }); commitPhotoEditorHistoryMutation();`);

  assertPixels(await exportAndSave(win, 'normal', false));
  await verifyLossyFormats(win, 'worker');
  await importMask(win);
  const transparent = await exportAndSave(win, 'transparent', true);
  assertPixels(transparent, 'right');
  await runInPage(win, `updatePhotoEditorSubjectMask({ invert: true }, { interactive: false });
    commitPhotoEditorHistoryMutation();`);
  const inverted = await exportAndSave(win, 'inverted', true);
  assertPixels(inverted, 'left');
  await runInPage(win, 'undoPhotoEditorEdit()');
  assert.deepEqual(await exportAndSave(win, 'undo', true), transparent,
    'Undo must restore all saved PNG pixels');
  await runInPage(win, 'redoPhotoEditorEdit()');
  assert.deepEqual(await exportAndSave(win, 'redo', true), inverted,
    'Redo must restore all saved PNG pixels');
  await verifyLossyFormats(win, 'main');
  const worldMetadata = await verifyReimportAndSync(win, sourceHash);
  assert.equal(hash(fs.readFileSync(SOURCE_PATH)), sourceHash, 'Original file changed');

  const report = { ok: true, smokeRoot: SMOKE_ROOT, sourcePath: SOURCE_PATH,
    sourceSha256: sourceHash, width: WIDTH, height: HEIGHT,
    worldIdentitySource: WORLD_IDENTITY_SOURCE, worldMetadata, savedFiles, blockedRequests };
  fs.writeFileSync(path.join(OUTPUT_DIR, 'result.json'), JSON.stringify(report, null, 2));
  console.log(`Photo editor export smoke passed: ${OUTPUT_DIR}`);
  app.quit();
}

const watchdog = setTimeout(() => {
  console.error('Export smoke exceeded 120 seconds');
  app.exit(1);
}, 120000);
main().catch(error => {
  fs.writeFileSync(path.join(OUTPUT_DIR, 'result.json'), JSON.stringify({
    ok: false, smokeRoot: SMOKE_ROOT, sourcePath: SOURCE_PATH,
    worldIdentitySource: WORLD_IDENTITY_SOURCE,
    error: error.stack || String(error), savedFiles, blockedRequests,
  }, null, 2));
  console.error(error.stack || error);
  console.error(`Isolated artifacts: ${SMOKE_ROOT}`);
  app.exit(1);
}).finally(() => clearTimeout(watchdog));

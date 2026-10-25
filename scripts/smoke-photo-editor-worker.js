// Real renderer/Worker comparison. Worker failure cannot fall back to renderer and pass.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { app, BrowserWindow, nativeImage, session } = require('electron');

const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-worker-'));
const OUTPUT_DIR = path.resolve(__dirname, '..', 'out', 'smoke-worker', path.basename(ROOT));
const PROFILE_DIR = path.join(ROOT, 'Profile');
const APPDATA_DIR = path.join(PROFILE_DIR, 'AppData', 'Roaming');
const LOCALAPPDATA_DIR = path.join(PROFILE_DIR, 'AppData', 'Local');
const TEMP_DIR = path.join(ROOT, 'Temp');
const SOURCE_PATH = path.join(ROOT, 'VRChat_2026-06-07_22-20-00.000_320x180.png');
const WIDTH = 320;
const HEIGHT = 180;
const blockedRequests = [];

for (const directory of [PROFILE_DIR, APPDATA_DIR, LOCALAPPDATA_DIR, TEMP_DIR, OUTPUT_DIR]) {
  fs.mkdirSync(directory, { recursive: true });
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
  const directory = path.join(PROFILE_DIR, key);
  fs.mkdirSync(directory, { recursive: true });
  app.setPath(key, directory);
}

const sha256 = buffer => createHash('sha256').update(buffer).digest('hex');

function seedPhoto() {
  const pixels = Buffer.alloc(WIDTH * HEIGHT * 4);
  for (let y = 0; y < HEIGHT; y += 1) {
    for (let x = 0; x < WIDTH; x += 1) {
      pixels.set([(x * 13 + y * 7) % 256, (x * 7 + y * 3) % 256,
        (x * 3 + y * 5) % 256, 255], (y * WIDTH + x) * 4);
    }
  }
  fs.writeFileSync(SOURCE_PATH, nativeImage.createFromBitmap(pixels, {
    width: WIDTH, height: HEIGHT, scaleFactor: 1,
  }).toPNG());
  const dataDir = path.join(app.getPath('userData'), 'data');
  fs.mkdirSync(dataDir, { recursive: true });
  const db = require('../src/db').initDatabase(path.join(dataDir, 'app.sqlite'));
  try {
    return db.insertOrUpdatePhoto({
      filePath: SOURCE_PATH, fileName: path.basename(SOURCE_PATH), fileHash: sha256(fs.readFileSync(SOURCE_PATH)),
      takenAt: '2026/06/07 22:20:00', takenAtTimestamp: new Date(2026, 5, 7, 22, 20).getTime(),
      groupDate: '2026-06-07', year: 2026, month: 6, day: 7,
      worldId: null, worldName: 'Worker fixture', worldNameManual: null, worldUrl: null,
      thumbnailPath: SOURCE_PATH, imageWidth: WIDTH, imageHeight: HEIGHT,
      resolutionTier: 'other', orientationTier: 'landscape', printNoteText: '', memoText: null,
      createdAt: '2026-06-07T13:20:00.000Z', updatedAt: '2026-06-07T13:20:00.000Z',
    }).id;
  } finally {
    db.close();
  }
}

async function until(label, predicate) {
  const deadline = Date.now() + 30000;
  while (Date.now() < deadline) {
    const value = await predicate();
    if (value) return value;
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  throw new Error(`Timed out: ${label}`);
}

async function main() {
  await app.whenReady();
  session.defaultSession.webRequest.onBeforeRequest(
    { urls: ['http://*/*', 'https://*/*', 'ws://*/*', 'wss://*/*'] },
    (details, callback) => { blockedRequests.push(details.url); callback({ cancel: true }); }
  );
  globalThis.fetch = async input => {
    blockedRequests.push(String(input));
    throw new Error('Network disabled in Worker smoke');
  };
  app.on('browser-window-created', (_event, window) => {
    window.webContents.setBackgroundThrottling(false);
    window.hide();
  });
  const photoId = seedPhoto();
  const sourceHash = sha256(fs.readFileSync(SOURCE_PATH));
  require('../src/index');
  for (const key of ['home', 'appData', 'userData', 'sessionData', 'temp']) {
    const relative = path.relative(ROOT, app.getPath(key));
    assert(relative && !relative.startsWith('..') && !path.isAbsolute(relative), `Unisolated ${key}`);
  }
  const win = await until('application window', () => BrowserWindow.getAllWindows()[0]);
  win.webContents.on('render-process-gone', (_event, details) => {
    console.error(details); app.exit(1);
  });
  const inPage = (fn, ...args) => win.webContents.executeJavaScript(
    `(${fn.toString()})(...${JSON.stringify(args)})`
  );
  await until('source photo card', () => inPage(id => Boolean(
    document.querySelector(`.photo-card[data-photo-id="${id}"]`)), photoId));
  await inPage(async id => {
    document.querySelector(`.photo-card[data-photo-id="${id}"]`).click();
    await openPhotoEditorModal();
  }, photoId);
  await until('source image', () => inPage(() => Boolean(photoEditorState?.sourceImage)));

  const comparisons = await inPage(async ({ width, height }) => {
    const check = (condition, message) => { if (!condition) throw new Error(message); };
    const digest = async data => Array.from(new Uint8Array(
      await crypto.subtle.digest('SHA-256', data)
    ), byte => byte.toString(16).padStart(2, '0')).join('');
    const compare = (actual, expected) => {
      check(actual.length === expected.length, 'Pixel buffer lengths differ');
      let differentChannels = 0;
      let maxDifference = 0;
      for (let index = 0; index < actual.length; index += 1) {
        const difference = Math.abs(actual[index] - expected[index]);
        if (difference) differentChannels += 1;
        maxDifference = Math.max(maxDifference, difference);
      }
      return { differentChannels, maxDifference };
    };
    const results = [];
    for (const mode of ['rgb', 'hsv', 'combined']) {
      photoEditorState.values = clonePhotoEditValues({
        brightness: 7, exposure: 13, contrast: -14, highlights: -21, whites: 11,
        shadows: 24, blacks: -9, gamma: 8, temperature: 12, tint: -7,
        saturation: 16, vibrance: 22, fade: 9, grain: 4,
      });
      const curve = getDefaultPhotoEditorCurveState();
      curve.mode = mode === 'hsv' ? 'hsv' : 'rgb';
      if (mode !== 'hsv') {
        curve.points.rgb.master = [0.03, 0.20, 0.55, 0.84, 0.98];
        curve.points.rgb.r = [0.02, 0.30, 0.57, 0.78, 1];
        curve.points.rgb.b = [0, 0.20, 0.46, 0.73, 0.96];
      }
      if (mode !== 'rgb') {
        curve.points.hsv.master = [0.01, 0.23, 0.53, 0.81, 0.99];
        curve.points.hsv.h = [0.02, 0.28, 0.51, 0.77, 0.97];
        curve.points.hsv.s = [0, 0.20, 0.57, 0.82, 1];
        curve.points.hsv.v = [0.03, 0.28, 0.54, 0.79, 0.99];
      }
      photoEditorState.curve = normalizePhotoEditorCurveState(curve);
      const directCanvas = document.createElement('canvas');
      const directBase = drawPhotoEditorBaseToCanvas(directCanvas);
      check(directBase.outputSize.width === width && directBase.outputSize.height === height,
        'Direct output dimensions changed');
      const original = directBase.ctx.getImageData(0, 0, width, height).data;
      applyPhotoEditorEffectsToCanvas(directBase.ctx, directBase.outputSize, directBase.sourceRect,
        { drawOverlays: false, drawTextOverlay: false, drawImageOverlay: false });
      const expected = directBase.ctx.getImageData(0, 0, width, height).data;
      const originalDifference = compare(expected, original);
      check(originalDifference.differentChannels > width * height,
        `${mode}: non-default adjustments must visibly change the fixture`);
      const expectedPixelHash = await digest(expected);

      for (const responseType of ['bitmap', 'blob']) {
        const workerCanvas = document.createElement('canvas');
        const workerBase = drawPhotoEditorBaseToCanvas(workerCanvas);
        // This helper rejects on Worker failure. Do not use the preview/export wrappers,
        // which can silently fall back to the main renderer and conceal Worker failures.
        const response = await applyPhotoEditorEffectsWithWorker(workerCanvas, workerBase, {
          responseType, purpose: 'export', exportSettings: { mimeType: 'image/png' },
        });
        check(response?.ok === true && response.type === 'render-result', `${mode}: missing Worker success`);
        check(Number.isInteger(response.requestId) && response.requestId > 0, 'Missing Worker request ID');
        check(response.width === width && response.height === height, 'Worker dimensions changed');
        let bitmap;
        if (responseType === 'bitmap') {
          check(response.bitmap instanceof ImageBitmap, 'Worker did not transfer an ImageBitmap');
          bitmap = response.bitmap;
        } else {
          check(response.buffer instanceof ArrayBuffer && response.buffer.byteLength > 8,
            'Worker did not transfer an encoded buffer');
          check(response.mimeType === 'image/png', 'Worker encoded the wrong format');
          bitmap = await createImageBitmap(new Blob([response.buffer], { type: response.mimeType }));
        }
        try {
          workerBase.ctx.clearRect(0, 0, width, height);
          workerBase.ctx.drawImage(bitmap, 0, 0);
        } finally {
          bitmap.close();
        }
        const actual = workerBase.ctx.getImageData(0, 0, width, height).data;
        const difference = compare(actual, expected);
        check(difference.differentChannels === 0,
          `${mode}/${responseType}: renderer/Worker pixels differ: ${JSON.stringify(difference)}`);
        results.push({ mode, responseType, requestId: response.requestId,
          width: response.width, height: response.height,
          originalDifference, ...difference, expectedPixelHash, actualPixelHash: await digest(actual) });
      }
    }
    return results;
  }, { width: WIDTH, height: HEIGHT });
  assert.equal(comparisons.length, 6);
  assert.equal(new Set(comparisons.map(result => result.requestId)).size, 6);
  assert.equal(new Set(comparisons.map(result => result.expectedPixelHash)).size, 3,
    'RGB, HSV and combined curves must produce distinct pixels');
  assert.equal(sha256(fs.readFileSync(SOURCE_PATH)), sourceHash, 'Source image was modified');
  const report = { ok: true, root: ROOT, sourceSha256: sourceHash, comparisons, blockedRequests };
  fs.writeFileSync(path.join(OUTPUT_DIR, 'result.json'), JSON.stringify(report, null, 2));
  console.log(`Photo editor real Worker smoke passed: ${OUTPUT_DIR}`);
  app.quit();
}

const watchdog = setTimeout(() => { console.error('Worker smoke exceeded 90 seconds'); app.exit(1); }, 90000);
main().catch(error => {
  console.error(error.stack || error);
  console.error(`Isolated artifacts: ${ROOT}`);
  app.exit(1);
}).finally(() => clearTimeout(watchdog));

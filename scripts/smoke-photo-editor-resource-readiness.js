// Real renderer export with synthetic images and no writes to the user's library.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { pathToFileURL } = require('node:url');
const { app, BrowserWindow, nativeImage, session } = require('electron');

const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-resource-readiness-'));
const OUTPUT_DIR = path.resolve(__dirname, '..', 'out', 'smoke-resource-readiness', path.basename(ROOT));
const PROFILE_DIR = path.join(ROOT, 'Profile');
const APPDATA_DIR = path.join(PROFILE_DIR, 'AppData', 'Roaming');
const LOCALAPPDATA_DIR = path.join(PROFILE_DIR, 'AppData', 'Local');
const TEMP_DIR = path.join(ROOT, 'Temp');
for (const directory of [PROFILE_DIR, APPDATA_DIR, LOCALAPPDATA_DIR, TEMP_DIR, OUTPUT_DIR]) fs.mkdirSync(directory, { recursive: true });
process.env.APPDATA = process.env.WORLDSHOT_APPDATA_ROOT = APPDATA_DIR;
process.env.LOCALAPPDATA = process.env.WORLDSHOT_LOCALAPPDATA_ROOT = LOCALAPPDATA_DIR;
process.env.USERPROFILE = process.env.WORLDSHOT_HOME_ROOT = PROFILE_DIR;
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

const report = { ok: false, root: ROOT, checks: [], goodOutputs: {}, blockedRequests: [] };
const baselinePath = path.join(__dirname, 'fixtures', 'photo-editor-resource-readiness-baseline.json');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
let win;
const inPage = (fn, ...args) => win.webContents.executeJavaScript(`(${fn.toString()})(...${JSON.stringify(args)})`);
async function until(label, predicate) {
  const deadline = Date.now() + 10000;
  while (Date.now() < deadline) {
    if (await predicate()) return;
    await new Promise(resolve => setTimeout(resolve, 20));
  }
  throw new Error(`Timed out: ${label}`);
}
async function check(name, run) {
  try { await run(); report.checks.push({ name, ok: true }); console.log(`PASS ${name}`); }
  catch (error) { report.checks.push({ name, ok: false, error: error.message }); throw error; }
}
function createPng(name, mask = false) {
  const width = 64, height = 48;
  const pixels = Buffer.alloc(width * height * 4);
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
    const values = mask ? (x < width / 2 ? [255, 255, 255, 255] : [0, 0, 0, 255])
      : [20 + y * 3, 35 + x * 2, 80 + x, 255];
    pixels.set(values, (y * width + x) * 4);
  }
  const filePath = path.join(ROOT, name);
  fs.writeFileSync(filePath, nativeImage.createFromBitmap(pixels, { width, height }).toPNG());
  return { filePath, fileUrl: pathToFileURL(filePath).href, fileName: name,
    dataUrl: `data:image/png;base64,${fs.readFileSync(filePath).toString('base64')}` };
}

async function main() {
  await app.whenReady();
  session.defaultSession.webRequest.onBeforeRequest(
    { urls: ['http://*/*', 'https://*/*', 'ws://*/*', 'wss://*/*'] },
    (details, callback) => { report.blockedRequests.push(details.url); callback({ cancel: true }); }
  );
  globalThis.fetch = async input => { report.blockedRequests.push(String(input)); throw new Error('Network disabled in resource readiness smoke'); };
  app.on('browser-window-created', (_event, window) => { window.webContents.setBackgroundThrottling(false); window.hide(); });
  require('../src/index');
  for (const key of ['home', 'appData', 'userData', 'sessionData', 'temp']) {
    const relative = path.relative(ROOT, app.getPath(key));
    assert(relative && !relative.startsWith('..') && !path.isAbsolute(relative), `Unisolated ${key}`);
  }
  await until('application window', () => { win = BrowserWindow.getAllWindows()[0]; return Boolean(win); });
  await until('renderer startup', () => inPage(() => typeof photoEditorSaveController !== 'undefined' &&
    !document.querySelector('.app')?.classList.contains('is-app-initializing')));
  const source = { ...createPng('source.png'), id: 1, takenAtTimestamp: 1780790400000 };
  const overlay = createPng('overlay.png', true);
  const mask = overlay.dataUrl;
  const corruptPath = path.join(ROOT, 'corrupt.png');
  fs.writeFileSync(corruptPath, 'not an image');
  const missingUrl = pathToFileURL(path.join(ROOT, 'missing.png')).href;
  const corruptUrl = pathToFileURL(corruptPath).href;
  const invalidMask = 'data:image/png;base64,bm90IGEgcG5n';
  const originalHash = hash(fs.readFileSync(source.filePath));
  await inPage(() => {
    const h = window.resourceReadinessSmoke = { errors: [], persistCalls: [], jobs: [] };
    const copy = value => JSON.parse(JSON.stringify(value));
    window.addEventListener('unhandledrejection', event => h.errors.push(String(event.reason)));
    schedulePhotoEditorRender = () => {};
    persistPhotoEditorExport = async payload => { h.persistCalls.push(payload); return { ok: true, canceled: true }; };
    h.capture = () => ({
      photoId: photoEditorState?.sourcePhoto?.id,
      active: photoEditorSession.isActive(photoEditorState),
      snapshot: capturePhotoEditorHistorySnapshot(),
      history: photoEditorSession.getHistoryState(),
      isSaving: photoEditorState?.isSaving,
      status: photoEditorStatus.textContent,
      saveDisabled: photoEditorSaveButton.disabled,
    });
    h.configure = (overlayUrl, maskUrl, options = {}) => {
      photoEditorOverlayImageCache.clear();
      photoEditorSubjectMaskImageCache.clear();
      clearPhotoEditorSubjectMaskEditCanvas();
      photoEditorState.values.brightness = options.brightness || 0;
      photoEditorState.adjustmentTarget = options.adjustmentTarget || 'whole';
      photoEditorState.imageOverlays = overlayUrl ? [normalizePhotoEditorImageOverlayState({
        id: 'readiness-overlay', assetId: 'readiness-asset', fileName: 'readiness-overlay.png', fileUrl: overlayUrl,
        x: 0.5, y: 0.5, width: 0.6, height: 0.6, opacity: options.opacity ?? 0.65,
        space: 'output', maskMode: options.maskMode || 'normal',
      })] : [];
      photoEditorState.subjectMask = normalizePhotoEditorSubjectMaskState({
        enabled: options.enabled !== false, status: 'ready', source: 'imported', maskDataUrl: maskUrl,
        width: 64, height: 48, showOverlay: options.showOverlay || false, opacity: options.maskOpacity ?? 0.55,
      });
    };
    h.decodeDataUrl = async dataUrl => {
      const response = await fetch(dataUrl);
      const bitmap = await createImageBitmap(await response.blob());
      const canvas = document.createElement('canvas');
      canvas.width = bitmap.width; canvas.height = bitmap.height;
      const ctx = canvas.getContext('2d'); ctx.drawImage(bitmap, 0, 0); bitmap.close();
      return { width: canvas.width, height: canvas.height, rgba: [...ctx.getImageData(0, 0, canvas.width, canvas.height).data] };
    };
    h.render = async options => {
      const result = await renderPhotoEditorExportDataUrl(document.createElement('canvas'),
        { format: 'png', maxEdge: 0, quality: 92 }, options);
      return h.decodeDataUrl(result.dataUrl);
    };
    h.startSave = options => {
      const job = { settled: false };
      h.jobs.push(job);
      job.promise = savePhotoEditorImage(options).then(() => { job.settled = true; });
      return h.jobs.length - 1;
    };
    h.deferResource = (kind, key) => {
      const entry = { image: photoEditorState.sourceImage, loaded: false, failed: false };
      entry.loadPromise = new Promise(resolve => { h.resolveResource = success => {
        entry.loaded = success; entry.failed = !success; resolve(entry);
      }; });
      (kind === 'overlay' ? photoEditorOverlayImageCache : photoEditorSubjectMaskImageCache).set(key, entry);
    };
  });
  async function open() {
    await inPage(async photo => { currentModalPhoto = photo; await openPhotoEditorModal(); }, source);
  }
  async function configure(overlayUrl = '', maskUrl = '', options = {}) {
    await open();
    await inPage((image, subject, settings) => window.resourceReadinessSmoke.configure(image, subject, settings), overlayUrl, maskUrl, options);
  }
  async function render() { return inPage(() => window.resourceReadinessSmoke.render()); }
  const goodCases = [
    ['plain', '', '', {}],
    ['overlay', overlay.fileUrl, '', { brightness: 14 }],
    ['transparent-overlay', corruptUrl, '', { opacity: 0 }],
    ['disabled-mask', '', invalidMask, { enabled: false }],
    ['hidden-mask-subject-adjustment', '', mask, { brightness: 35, adjustmentTarget: 'subject', maskOpacity: 0, showOverlay: false }],
    ['subject-only-overlay', overlay.fileUrl, mask, { brightness: 14, adjustmentTarget: 'subject', maskMode: 'subject-only' }],
    ['background-only-overlay', overlay.fileUrl, mask, { brightness: 14, adjustmentTarget: 'background', maskMode: 'background-only' }],
  ];
  const baseline = JSON.parse(fs.readFileSync(baselinePath, 'utf8'));
  for (const [name, image, subject, options] of goodCases) {
    await check(`valid output remains pixel-identical: ${name}`, async () => {
      await configure(image, subject, options);
      const result = await render();
      const measured = { width: result.width, height: result.height, rgbaSha256: hash(Buffer.from(result.rgba)) };
      report.goodOutputs[name] = measured;
      assert.deepEqual(measured, baseline.goodOutputs[name]);
    });
  }
  assert.equal(originalHash, baseline.sourceSha256);
  assert.deepEqual(report.goodOutputs.plain, report.goodOutputs['transparent-overlay']);
  assert.deepEqual(report.goodOutputs.plain, report.goodOutputs['disabled-mask']);
  assert.notDeepEqual(report.goodOutputs.plain, report.goodOutputs['hidden-mask-subject-adjustment']);

  for (const [name, image, subject, options, message] of [
    ['missing overlay', missingUrl, '', {}, '画像オーバーレイ素材を読み込めませんでした'],
    ['corrupt overlay', corruptUrl, '', {}, '画像オーバーレイ素材を読み込めませんでした'],
    ['corrupt enabled mask', '', invalidMask, { brightness: 35, adjustmentTarget: 'subject' }, 'マスク画像を読み込めませんでした'],
    ['hidden zero-preview-opacity enabled mask', '', invalidMask, { brightness: 35, adjustmentTarget: 'background', maskOpacity: 0 }, 'マスク画像を読み込めませんでした'],
  ]) {
    await check(`${name} blocks persistence and leaves editing usable`, async () => {
      await configure(image, subject, options);
      const result = await inPage(async () => {
        const h = window.resourceReadinessSmoke;
        const before = h.capture();
        const count = h.persistCalls.length;
        await savePhotoEditorImage();
        return { before, after: h.capture(), newPersistCalls: h.persistCalls.length - count };
      });
      assert.equal(result.newPersistCalls, 0, 'Incomplete images must never reach the persistence IPC');
      assert.equal(result.after.active, true);
      assert.equal(result.after.isSaving, false);
      assert.equal(result.after.saveDisabled, false);
      assert.deepEqual(result.after.snapshot, result.before.snapshot);
      assert.deepEqual(result.after.history, result.before.history);
      assert(result.after.status.includes(message), result.after.status);
    });
  }

  await check('Undo removes a broken layer so export resumes; Redo requires it again', async () => {
    await configure();
    await inPage(url => {
      beginPhotoEditorHistoryMutation();
      photoEditorState.imageOverlays = [normalizePhotoEditorImageOverlayState({ id: 'undo-broken', fileUrl: url, opacity: 1 })];
      commitPhotoEditorHistoryMutation();
      undoPhotoEditorEdit();
    }, missingUrl);
    assert.equal(hash(Buffer.from((await render()).rgba)), baseline.goodOutputs.plain.rgbaSha256);
    const result = await inPage(async () => {
      redoPhotoEditorEdit();
      const h = window.resourceReadinessSmoke;
      const count = h.persistCalls.length;
      await savePhotoEditorImage();
      return { count: h.persistCalls.length - count, layers: photoEditorState.imageOverlays.length, canUndo: photoEditorSession.getHistoryStatus().canUndo };
    });
    assert.deepEqual(result, { count: 0, layers: 1, canUndo: true });
  });

  await check('an active brush canvas remains the mask source while its stored image is unavailable', async () => {
    await configure('', invalidMask, { brightness: 35, adjustmentTarget: 'subject', maskOpacity: 0 });
    await inPage(() => {
      const canvas = document.createElement('canvas');
      canvas.width = 64; canvas.height = 48;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#000'; ctx.fillRect(0, 0, 64, 48);
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, 32, 48);
      photoEditorSubjectMaskEditCanvas = canvas;
    });
    assert.equal(hash(Buffer.from((await render()).rgba)), baseline.goodOutputs['hidden-mask-subject-adjustment'].rgbaSha256);
  });
  await check('public wait helpers retain their serializable void result', async () => {
    await configure(overlay.fileUrl, mask);
    assert.equal(await inPage(() => waitForPhotoEditorSubjectMaskToLoad()), undefined);
    assert.equal(await inPage(() => waitForPhotoEditorImageOverlaysToLoad(photoEditorState.imageOverlays)), undefined);
  });
  await check('transparent export rejects a corrupt required mask before persistence', async () => {
    await configure('', invalidMask);
    const result = await inPage(async () => {
      const h = window.resourceReadinessSmoke, count = h.persistCalls.length;
      await savePhotoEditorImage({ transparentSubjectBackground: true });
      return { calls: h.persistCalls.length - count, state: h.capture() };
    });
    assert.equal(result.calls, 0);
    assert.equal(result.state.isSaving, false);
    assert(result.state.status.includes('マスク画像を読み込めませんでした'));
  });

  async function defer(kind, options = {}) {
    await configure(kind === 'overlay' ? overlay.fileUrl : '', kind === 'mask' ? mask : '', options);
    await inPage((type, key) => window.resourceReadinessSmoke.deferResource(type, key), kind, kind === 'overlay' ? overlay.fileUrl : mask);
  }
  async function start() { return inPage(() => window.resourceReadinessSmoke.startSave()); }
  async function settle(job, success = true) {
    return inPage(async (id, ready) => {
      const h = window.resourceReadinessSmoke, before = h.capture(), count = h.persistCalls.length;
      h.resolveResource(ready);
      await h.jobs[id].promise;
      return { before, after: h.capture(), newPersistCalls: h.persistCalls.length - count };
    }, job, success);
  }
  for (const kind of ['overlay', 'mask']) {
    await check(`${kind} loading holds persistence until decoded`, async () => {
      await defer(kind);
      const job = await start();
      assert.equal(await inPage(id => window.resourceReadinessSmoke.jobs[id].settled, job), false);
      const result = await settle(job);
      assert.equal(result.newPersistCalls, 1);
      assert.equal(result.after.isSaving, false);
      assert.equal(result.after.saveDisabled, false);
    });
    for (const success of [true, false]) {
      await check(`late ${kind} ${success ? 'success' : 'failure'} cannot save or alter a new editor`, async () => {
        await defer(kind);
        const job = await start();
        await inPage(async photo => { currentModalPhoto = { ...photo, id: 2 }; await openPhotoEditorModal(); }, source);
        const result = await settle(job, success);
        assert.equal(result.newPersistCalls, 0);
        assert.deepEqual(result.after, result.before);
      });
    }
  }
  for (const reopen of [false, true]) {
    await check(`closing during resource decode ${reopen ? 'and reopening the same photo' : ''} prevents stale persistence`, async () => {
      await defer('overlay');
      const job = await start();
      await inPage(() => closePhotoEditorModal());
      if (reopen) await open();
      else await until('closed editor', () => inPage(() => photoEditorState === null));
      const result = await settle(job);
      assert.equal(result.newPersistCalls, 0);
      assert.deepEqual(result.after, result.before);
    });
  }
  await check('Undo while a layer is loading cancels export and retains the new edit state', async () => {
    await defer('overlay');
    await inPage(() => {
      beginPhotoEditorHistoryMutation();
      photoEditorState.values.brightness = 20;
      commitPhotoEditorHistoryMutation();
    });
    const job = await start();
    await inPage(() => undoPhotoEditorEdit());
    const result = await settle(job);
    assert.equal(result.newPersistCalls, 0);
    assert.equal(result.after.snapshot.values.brightness, 0);
    assert.deepEqual(result.after.history, result.before.history);
    assert.equal(result.after.isSaving, false);
    assert.equal(result.after.saveDisabled, false);
  });
  await check('Undo and Redo with identical final settings still invalidate an evicted mask dependency', async () => {
    await defer('mask');
    await inPage(() => {
      beginPhotoEditorHistoryMutation();
      photoEditorState.values.brightness = 20;
      commitPhotoEditorHistoryMutation();
    });
    const job = await start();
    await inPage(() => { undoPhotoEditorEdit(); redoPhotoEditorEdit(); });
    const result = await settle(job);
    assert.equal(result.newPersistCalls, 0);
    assert.equal(result.after.snapshot.values.brightness, 20);
    assert.deepEqual(result.after.history, result.before.history);
    assert.equal(result.after.isSaving, false);
    assert.equal(result.after.saveDisabled, false);
  });
  for (const action of ['edit', 'evict-resource']) {
    await check(`${action} during Worker rendering rejects the old bitmap and restores save controls`, async () => {
      await configure(overlay.fileUrl);
      await inPage(() => {
        const h = window.resourceReadinessSmoke;
        h.originalWorker = applyPhotoEditorEffectsWithWorker;
        h.resolveWorker = null;
        h.bitmapClosed = false;
        applyPhotoEditorEffectsWithWorker = () => new Promise(resolve => { h.resolveWorker = resolve; });
      });
      try {
        const job = await start();
        await until('Worker export', () => inPage(() => Boolean(window.resourceReadinessSmoke.resolveWorker)));
        const result = await inPage(async (id, change) => {
          const h = window.resourceReadinessSmoke, count = h.persistCalls.length;
          if (change === 'edit') {
            beginPhotoEditorHistoryMutation(); photoEditorState.values.brightness = 30; commitPhotoEditorHistoryMutation();
          } else photoEditorOverlayImageCache.clear();
          const before = h.capture();
          h.resolveWorker({ bitmap: { close() { h.bitmapClosed = true; } } });
          await h.jobs[id].promise;
          return { before, after: h.capture(), newPersistCalls: h.persistCalls.length - count, bitmapClosed: h.bitmapClosed };
        }, job, action);
        assert.equal(result.newPersistCalls, 0);
        assert.equal(result.bitmapClosed, true);
        assert.deepEqual(result.after.snapshot, result.before.snapshot);
        assert.deepEqual(result.after.history, result.before.history);
        assert.equal(result.after.isSaving, false);
        assert.equal(result.after.saveDisabled, false);
      } finally {
        await inPage(() => { applyPhotoEditorEffectsWithWorker = window.resourceReadinessSmoke.originalWorker; });
      }
    });
  }

  // Read a font already supplied by Windows. Its bytes stay in memory and are
  // never copied into the repository or distribution.
  const fontPath = path.join(process.env.WINDIR || 'C:/Windows', 'Fonts', 'times.ttf');
  const fontBytes = fs.readFileSync(fontPath);
  report.fontSource = { fileName: path.basename(fontPath), sha256: hash(fontBytes) };
  await inPage(async base64 => {
    const h = window.resourceReadinessSmoke;
    h.originalFontLoad = document.fonts.load.bind(document.fonts);
    const bytes = Uint8Array.from(atob(base64), character => character.charCodeAt(0));
    h.testFont = await new FontFace('Bebas Neue', bytes, { weight: '400' }).load();
    h.setFontMode = mode => {
      document.fonts.delete(h.testFont);
      h.fontRequests = [];
      h.fontGate = new Promise((resolve, reject) => { h.releaseFont = async success => {
        if (success) { document.fonts.add(h.testFont); resolve([h.testFont]); }
        else reject(new Error('controlled offline font'));
      }; });
      // Attach rejection handling even when a test switches before a load starts.
      h.fontGate.catch(() => {});
      document.fonts.load = (font, text) => {
        if (!font.includes('Bebas Neue')) return h.originalFontLoad(font, text);
        h.fontRequests.push({ font, text });
        if (mode === 'offline') return Promise.reject(new Error('controlled offline font'));
        if (mode === 'ready') { document.fonts.add(h.testFont); return Promise.resolve([h.testFont]); }
        return h.fontGate;
      };
    };
    h.setTestText = () => {
      updatePhotoEditorTextOverlay({ text: 'WWW iii', fontKey: 'bebasNeue', weight: '400', size: 220,
        strokeType: 'none', strokeWidth: 0, color: '#ffffff', space: 'output' }, { interactive: false });
      commitPhotoEditorHistoryMutation();
    };
  }, fontBytes.toString('base64'));
  try {
    await check('saving while the selected font loads waits and exports its final glyphs', async () => {
      await configure();
      await inPage(() => { const h = window.resourceReadinessSmoke; h.setFontMode('pending'); h.setTestText(); });
      const result = await inPage(async () => {
        const h = window.resourceReadinessSmoke;
        const count = h.persistCalls.length;
        const job = h.startSave();
        const early = await Promise.race([h.jobs[job].promise.then(() => 'settled'),
          new Promise(resolve => setTimeout(() => resolve('pending'), 200))]);
        const savedBeforeFont = h.persistCalls.length > count ? await h.decodeDataUrl(h.persistCalls.at(-1).dataUrl) : null;
        await h.releaseFont(true);
        await h.jobs[job].promise;
        const saved = await h.decodeDataUrl(h.persistCalls.at(-1).dataUrl);
        const ready = await h.render();
        return { early, savedBeforeFont, saved, ready, newPersistCalls: h.persistCalls.length - count, state: h.capture() };
      });
      report.fontLoadingObservation = {
        early: result.early,
        savedBeforeFontSha256: result.savedBeforeFont ? hash(Buffer.from(result.savedBeforeFont.rgba)) : null,
        savedSha256: hash(Buffer.from(result.saved.rgba)), readySha256: hash(Buffer.from(result.ready.rgba)),
      };
      assert.equal(result.early, 'pending', 'Saving must wait for the selected font instead of using fallback glyphs');
      assert.equal(result.newPersistCalls, 1);
      assert.deepEqual(result.saved, result.ready);
      assert.equal(result.state.isSaving, false);
      assert.equal(result.state.saveDisabled, false);
    });
    await check('offline font loading keeps the existing fallback glyphs and successful save', async () => {
      await configure();
      await inPage(() => { const h = window.resourceReadinessSmoke; h.setFontMode('offline'); h.setTestText(); });
      const result = await inPage(async () => {
        const h = window.resourceReadinessSmoke, count = h.persistCalls.length;
        await savePhotoEditorImage();
        const saved = await h.decodeDataUrl(h.persistCalls.at(-1).dataUrl);
        return { saved, fallback: await h.render(), calls: h.persistCalls.length - count, state: h.capture() };
      });
      report.offlineFontSha256 = hash(Buffer.from(result.saved.rgba));
      assert.equal(result.calls, 1);
      assert.deepEqual(result.saved, result.fallback);
      assert.notEqual(report.offlineFontSha256, report.fontLoadingObservation.readySha256);
      assert.equal(result.state.isSaving, false);
    });
    await check('a permanently pending font uses fallback after the five-second deadline', async () => {
      await configure();
      await inPage(() => { const h = window.resourceReadinessSmoke; h.setFontMode('pending'); h.setTestText(); });
      const result = await inPage(async () => {
        const h = window.resourceReadinessSmoke, count = h.persistCalls.length, started = performance.now();
        await savePhotoEditorImage();
        const elapsedMs = performance.now() - started;
        const saved = await h.decodeDataUrl(h.persistCalls.at(-1).dataUrl);
        await h.releaseFont(false);
        return { saved, elapsedMs, calls: h.persistCalls.length - count, state: h.capture() };
      });
      report.fontTimeoutMs = result.elapsedMs;
      assert.equal(result.calls, 1);
      assert(result.elapsedMs >= 4500 && result.elapsedMs < 10000, `Unexpected font wait ${result.elapsedMs}ms`);
      assert.equal(hash(Buffer.from(result.saved.rgba)), report.offlineFontSha256);
      assert.equal(result.state.isSaving, false);
      assert.equal(result.state.saveDisabled, false);
    });
    await check('invisible text does not wait for a font that cannot affect the image', async () => {
      await configure();
      await inPage(() => {
        const h = window.resourceReadinessSmoke;
        h.setFontMode('pending'); h.setTestText();
        updatePhotoEditorTextOverlay({ fillTransparent: true, strokeType: 'none' }, { interactive: false });
        commitPhotoEditorHistoryMutation();
      });
      const result = await inPage(async () => {
        const h = window.resourceReadinessSmoke, requests = h.fontRequests.length, count = h.persistCalls.length;
        await savePhotoEditorImage();
        const saved = await h.decodeDataUrl(h.persistCalls.at(-1).dataUrl);
        await h.releaseFont(false);
        return { saved, addedFontRequests: h.fontRequests.length - requests, calls: h.persistCalls.length - count };
      });
      assert.equal(result.calls, 1);
      assert.equal(result.addedFontRequests, 0);
      assert.equal(hash(Buffer.from(result.saved.rgba)), baseline.goodOutputs.plain.rgbaSha256);
    });
    for (const action of ['switch-editor', 'undo', 'evict-resource']) {
      await check(`${action} during font loading prevents stale or incomplete persistence`, async () => {
        await configure(overlay.fileUrl);
        await inPage(() => { const h = window.resourceReadinessSmoke; h.setFontMode('pending'); h.setTestText(); });
        const job = await start();
        await until('export font request', () => inPage(() => window.resourceReadinessSmoke.fontRequests.length >= 2));
        if (action === 'switch-editor') {
          await inPage(async photo => { currentModalPhoto = { ...photo, id: 2 }; await openPhotoEditorModal(); }, source);
        } else await inPage(change => {
          if (change === 'undo') undoPhotoEditorEdit();
          else photoEditorOverlayImageCache.clear();
        }, action);
        const result = await inPage(async id => {
          const h = window.resourceReadinessSmoke, before = h.capture(), count = h.persistCalls.length;
          await h.releaseFont(true);
          await h.jobs[id].promise;
          return { before, after: h.capture(), calls: h.persistCalls.length - count };
        }, job);
        assert.equal(result.calls, 0);
        assert.deepEqual(result.after.snapshot, result.before.snapshot);
        assert.deepEqual(result.after.history, result.before.history);
        assert.equal(result.after.isSaving, false);
        assert.equal(result.after.saveDisabled, false);
        if (action === 'switch-editor') assert.deepEqual(result.after, result.before);
      });
    }
  } finally {
    await inPage(() => {
      const h = window.resourceReadinessSmoke;
      document.fonts.load = h.originalFontLoad;
      document.fonts.delete(h.testFont);
    });
  }

  assert.deepEqual(await inPage(() => window.resourceReadinessSmoke.errors), []);
  assert.equal(hash(fs.readFileSync(source.filePath)), originalHash);
  report.ok = true;
  console.log(`Resource readiness smoke passed: ${OUTPUT_DIR}`);
}
const watchdog = setTimeout(() => { console.error('Resource readiness smoke exceeded 90 seconds'); app.exit(1); }, 90000);
main().then(() => {
  fs.writeFileSync(path.join(OUTPUT_DIR, 'result.json'), JSON.stringify(report, null, 2));
  clearTimeout(watchdog); app.quit();
}).catch(error => {
  report.error = error.stack || String(error);
  fs.writeFileSync(path.join(OUTPUT_DIR, 'result.json'), JSON.stringify(report, null, 2));
  console.error(error); clearTimeout(watchdog); app.exit(1);
});

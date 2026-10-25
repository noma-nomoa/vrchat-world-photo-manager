// Real renderer/preload with controlled overlay IPC, confirmation and Worker boundaries.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { pathToFileURL } = require('node:url');
const { app, BrowserWindow, nativeImage, session, ipcMain } = require('electron');

const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-overlay-session-'));
const OUTPUT_DIR = path.resolve(__dirname, '..', 'out', 'smoke-overlay-session', path.basename(ROOT));
const PROFILE_DIR = path.join(ROOT, 'Profile');
const APPDATA_DIR = path.join(PROFILE_DIR, 'AppData', 'Roaming');
const LOCALAPPDATA_DIR = path.join(PROFILE_DIR, 'AppData', 'Local');
const TEMP_DIR = path.join(ROOT, 'Temp');
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

const report = { ok: false, root: ROOT, checks: [], blockedRequests: [] };
const gates = [];
let win;
const inPage = (fn, ...args) => win.webContents.executeJavaScript(`(${fn.toString()})(...${JSON.stringify(args)})`);
const hashFile = file => createHash('sha256').update(fs.readFileSync(file)).digest('hex');
async function until(label, predicate) {
  const deadline = Date.now() + 10000;
  while (Date.now() < deadline) {
    if (await predicate()) return;
    await new Promise(resolve => setTimeout(resolve, 20));
  }
  throw new Error(`Timed out: ${label}`);
}
function createPhoto(id, width, height) {
  const filePath = path.join(ROOT, `source-${id}.png`);
  const pixels = Buffer.alloc(width * height * 4);
  for (let index = 0; index < pixels.length; index += 4) pixels.set([40, 100, 180, 255], index);
  fs.writeFileSync(filePath, nativeImage.createFromBitmap(pixels, { width, height }).toPNG());
  return { id, filePath, fileName: path.basename(filePath), fileUrl: pathToFileURL(filePath).href };
}
async function openPhoto(photo) {
  await inPage(async value => { currentModalPhoto = value; await openPhotoEditorModal(); }, photo);
  await until('source image', () => inPage(() => Boolean(photoEditorState?.sourceImage)));
}
async function start(kind, asset) {
  return inPage((name, value) => window.overlaySessionSmoke.start(name, value), kind, asset);
}
async function waitIpc(kind) {
  await until(`${kind} IPC`, () => gates.some(gate => gate.kind === kind && !gate.settled));
  return gates.find(gate => gate.kind === kind && !gate.settled);
}
function release(gate, result, reject = false) {
  gate.settled = true;
  if (reject) gate.reject(new Error(result));
  else gate.resolve(result);
}
async function finish(job) {
  const outcome = await inPage(async id => {
    const job = window.overlaySessionSmoke.jobs[id];
    await job.promise;
    return job.outcome;
  }, job);
  assert.equal(outcome, 'resolved', 'The UI action handles its own failure');
}
async function capture() { return inPage(() => window.overlaySessionSmoke.capture()); }
async function confirm(job, value = true) {
  await inPage((id, confirmed) => window.overlaySessionSmoke.jobs[id].confirm(confirmed), job, value);
}
async function check(name, run) {
  try {
    await run();
    report.checks.push({ name, ok: true });
    console.log(`PASS ${name}`);
  } catch (error) {
    report.checks.push({ name, ok: false, error: error.message });
    throw error;
  }
}

async function main() {
  await app.whenReady();
  session.defaultSession.webRequest.onBeforeRequest(
    { urls: ['http://*/*', 'https://*/*', 'ws://*/*', 'wss://*/*'] },
    (details, callback) => { report.blockedRequests.push(details.url); callback({ cancel: true }); }
  );
  globalThis.fetch = async input => {
    report.blockedRequests.push(String(input));
    throw new Error('Network disabled in overlay session smoke');
  };
  app.on('browser-window-created', (_event, window) => {
    window.webContents.setBackgroundThrottling(false);
    window.hide();
  });
  require('../src/index');
  for (const key of ['home', 'appData', 'userData', 'sessionData', 'temp']) {
    const relative = path.relative(ROOT, app.getPath(key));
    assert(relative && !relative.startsWith('..') && !path.isAbsolute(relative), `Unisolated ${key}`);
  }
  await until('application window', () => { win = BrowserWindow.getAllWindows()[0]; return Boolean(win); });
  await until('renderer startup', () => inPage(() =>
    typeof selectPhotoEditorOverlayImages === 'function' &&
    !document.querySelector('.app')?.classList.contains('is-app-initializing')));
  for (const [channel, kind] of [
    ['select-photo-editor-overlay-images', 'import'],
    ['delete-photo-editor-overlay-asset', 'delete'],
  ]) {
    ipcMain.removeHandler(channel);
    ipcMain.handle(channel, (_event, payload) => new Promise((resolve, reject) => {
      gates.push({ kind, payload, resolve, reject, settled: false });
    }));
  }
  await inPage(() => {
    const h = window.overlaySessionSmoke = { jobs: [], errors: [], toasts: [], previewCalls: [] };
    const copy = value => JSON.parse(JSON.stringify(value));
    window.addEventListener('unhandledrejection', event => h.errors.push(String(event.reason)));
    // Explicit renders below isolate callback effects from unrelated animation frames.
    schedulePhotoEditorRender = () => {};
    showToast = message => h.toasts.push(message);
    h.capture = () => ({
      photoId: photoEditorState?.sourcePhoto?.id ?? null,
      active: photoEditorSession.isActive(photoEditorState),
      snapshot: photoEditorState ? capturePhotoEditorHistorySnapshot() : null,
      history: photoEditorSession.getHistoryState(),
      status: document.getElementById('photo-editor-status').textContent,
      modalClass: document.getElementById('photo-editor-modal').className,
      toasts: copy(h.toasts),
    });
    openConfirmModal = () => new Promise(resolve => { h.jobs[h.jobs.length - 1].confirm = resolve; });
    h.start = (kind, asset) => {
      const job = { outcome: 'pending' };
      h.jobs.push(job);
      job.promise = (kind === 'import' ? selectPhotoEditorOverlayImages() : deletePhotoEditorManagedOverlayAsset(asset))
        .then(() => { job.outcome = 'resolved'; }, error => { job.outcome = String(error); });
      return h.jobs.length - 1;
    };
  });

  const photoA = createPhoto(1, 80, 60);
  const photoB = createPhoto(2, 100, 50);
  const sourceHashes = [photoA, photoB].map(photo => hashFile(photo.filePath));
  const asset = { id: 'controlled-overlay', fileName: 'controlled-overlay.png', fileUrl: photoA.fileUrl, width: 80, height: 60 };
  const asset2 = { ...asset, id: 'second-overlay', fileName: 'second-overlay.png', fileUrl: photoB.fileUrl };
  const imported = (assets = [asset]) => ({ ok: true, assets, importedAssets: assets });

  await check('import adds overlays with undo and redo preserved', async () => {
    await openPhoto(photoA);
    const job = await start('import');
    release(await waitIpc('import'), imported());
    await finish(job);
    assert.equal((await capture()).snapshot.imageOverlays.length, 1);
    const history = await inPage(() => {
      undoPhotoEditorEdit();
      const undo = photoEditorState.imageOverlays.length;
      redoPhotoEditorEdit();
      return { undo, redo: photoEditorState.imageOverlays.length };
    });
    assert.deepEqual(history, { undo: 0, redo: 1 });
  });
  await check('cancelled import keeps editing state and refreshes managed library', async () => {
    await openPhoto(photoA);
    const before = await capture();
    const job = await start('import');
    release(await waitIpc('import'), { ok: true, canceled: true, assets: [asset2] });
    await finish(job);
    assert.deepEqual(await capture(), before);
    assert.deepEqual(await inPage(() => photoEditorOverlayAssets.map(value => value.id)), [asset2.id]);
  });
  await check('active import failure remains visible without adding layers', async () => {
    await openPhoto(photoA);
    const job = await start('import');
    release(await waitIpc('import'), { ok: false, message: 'controlled import failure', assets: [] });
    await finish(job);
    const state = await capture();
    assert.equal(state.snapshot.imageOverlays.length, 0);
    assert.equal(state.status, 'controlled import failure');
  });
  await check('separate imports in one session both remain valid when resolved out of order', async () => {
    await openPhoto(photoA);
    const first = await start('import');
    const firstGate = await waitIpc('import');
    const second = await start('import');
    await until('second import IPC', () => gates.filter(gate => gate.kind === 'import' && !gate.settled).length === 2);
    const secondGate = gates.find(gate => gate.kind === 'import' && !gate.settled && gate !== firstGate);
    release(secondGate, imported([asset2]));
    await finish(second);
    release(firstGate, imported([asset]));
    await finish(first);
    assert.deepEqual((await capture()).snapshot.imageOverlays.map(value => value.assetId), [asset2.id, asset.id]);
  });

  for (const outcome of ['success', 'failure', 'rejection']) {
    await check(`A import ${outcome} after direct B open preserves B state and history`, async () => {
      await openPhoto(photoA);
      const job = await start('import');
      const gate = await waitIpc('import');
      await openPhoto(photoB);
      const before = await capture();
      release(gate, outcome === 'success' ? imported() : outcome === 'failure'
        ? { ok: false, message: 'stale import failure' } : 'stale IPC rejection', outcome === 'rejection');
      await finish(job);
      assert.deepEqual(await capture(), before);
      if (outcome === 'success') {
        assert.deepEqual(await inPage(() => photoEditorOverlayAssets.map(value => value.id)), [asset.id]);
      }
    });
  }
  await check('A import after close animation begins does not alter closing state', async () => {
    await openPhoto(photoA);
    const job = await start('import');
    const gate = await waitIpc('import');
    await inPage(() => {
      const h = window.overlaySessionSmoke;
      h.originalCloseElement = closeSubModalElement;
      closeSubModalElement = (_element, options) => { h.completeClose = options.onClosed; };
      closePhotoEditorModal();
    });
    try {
      const before = await capture();
      assert.equal(before.active, false);
      assert.equal(before.photoId, photoA.id, 'The old state still exists during the animation');
      release(gate, imported());
      await finish(job);
      assert.deepEqual(await capture(), before);
    } finally {
      await inPage(() => {
        const h = window.overlaySessionSmoke;
        closeSubModalElement = h.originalCloseElement;
        h.completeClose();
      });
    }
  });
  await check('same photo reopened has a new import owner', async () => {
    await openPhoto(photoA);
    const job = await start('import');
    const gate = await waitIpc('import');
    await inPage(() => closePhotoEditorModal());
    await openPhoto(photoA);
    const before = await capture();
    release(gate, imported());
    await finish(job);
    assert.deepEqual(await capture(), before);
  });

  async function prepareDelete() {
    await openPhoto(photoA);
    await inPage(value => addPhotoEditorImageOverlayFromAsset(value), asset);
    return start('delete', asset);
  }
  await check('confirmed active deletion removes matching layer and updates library', async () => {
    const job = await prepareDelete();
    await confirm(job);
    const gate = await waitIpc('delete');
    assert.equal(gate.payload.assetId, asset.id);
    release(gate, { ok: true, assets: [] });
    await finish(job);
    assert.equal((await capture()).snapshot.imageOverlays.length, 0);
    assert.deepEqual(await inPage(() => photoEditorOverlayAssets), []);
    assert.equal(await inPage(() => document.getElementById('photo-editor-image-overlay-library').children.length), 0);
  });
  await check('cancelled deletion does not invoke destructive IPC or change state', async () => {
    const job = await prepareDelete();
    const before = await capture();
    const callCount = gates.length;
    await confirm(job, false);
    await finish(job);
    assert.equal(gates.length, callCount);
    assert.deepEqual(await capture(), before);
  });
  await check('confirmation from A after B opens does not invoke deletion', async () => {
    const job = await prepareDelete();
    await openPhoto(photoB);
    const before = await capture();
    const callCount = gates.length;
    await confirm(job);
    await finish(job);
    assert.equal(gates.length, callCount);
    assert.deepEqual(await capture(), before);
  });
  for (const outcome of ['success', 'failure', 'rejection']) {
    await check(`A deletion ${outcome} after B opens preserves B references and cached image`, async () => {
      const job = await prepareDelete();
      await confirm(job);
      const gate = await waitIpc('delete');
      await openPhoto(photoB);
      await inPage(value => addPhotoEditorImageOverlayFromAsset(value), asset);
      await until('B cached overlay decoded', () => inPage(url => photoEditorOverlayImageCache.get(url)?.loaded, asset.fileUrl));
      await inPage(url => { window.overlaySessionSmoke.retainedEntry = photoEditorOverlayImageCache.get(url); }, asset.fileUrl);
      const before = await capture();
      release(gate, outcome === 'success' ? { ok: true, assets: [] } : outcome === 'failure'
        ? { ok: false, message: 'stale delete failure' } : 'stale delete IPC rejection', outcome === 'rejection');
      await finish(job);
      assert.deepEqual(await capture(), before);
      assert.equal(await inPage(url => photoEditorOverlayImageCache.get(url) === window.overlaySessionSmoke.retainedEntry &&
        photoEditorOverlayImageCache.get(url)?.loaded, asset.fileUrl), true);
      if (outcome === 'success') assert.deepEqual(await inPage(() => photoEditorOverlayAssets), []);
    });
  }
  await check('active deletion failure retains layer and explains failure', async () => {
    const job = await prepareDelete();
    await confirm(job);
    release(await waitIpc('delete'), { ok: false, message: 'controlled delete failure' });
    await finish(job);
    const state = await capture();
    assert.equal(state.snapshot.imageOverlays.length, 1);
    assert.match(state.status, /controlled delete failure/);
  });

  for (const outcome of ['bitmap', 'rejection']) {
    await check(`A preview Worker ${outcome} during B source load cannot commit an A frame`, async () => {
      await openPhoto(photoA);
      await inPage(() => {
        const h = window.overlaySessionSmoke;
        h.previewOriginals = { worker: applyPhotoEditorEffectsWithWorker, commit: commitPhotoEditorPreviewFrame, Image: window.Image };
        h.bitmapClosed = false;
        h.previewCalls = [];
        commitPhotoEditorPreviewFrame = () => h.previewCalls.push('commit');
        applyPhotoEditorEffectsWithWorker = () => new Promise((resolve, reject) => { h.resolveWorker = resolve; h.rejectWorker = reject; });
        h.previewPromise = renderPhotoEditorPreview();
        window.Image = class DeferredImage {
          constructor() { h.sourceImage = this; }
          set src(value) { this.srcValue = value; }
        };
      });
      try {
        await inPage(value => {
          currentModalPhoto = value;
          window.overlaySessionSmoke.openPromise = openPhotoEditorModal();
        }, photoB);
        const result = await inPage(async value => {
          const h = window.overlaySessionSmoke;
          const before = h.capture();
          if (value === 'rejection') h.rejectWorker(new Error('stale preview Worker failure'));
          else h.resolveWorker({ bitmap: { close() { h.bitmapClosed = true; } } });
          await h.previewPromise;
          return { before, after: h.capture(), calls: h.previewCalls, bitmapClosed: h.bitmapClosed };
        }, outcome);
        assert.deepEqual(result.after, result.before);
        assert.deepEqual(result.calls, []);
        if (outcome === 'bitmap') assert.equal(result.bitmapClosed, true);
      } finally {
        await inPage(async () => {
          const h = window.overlaySessionSmoke;
          window.Image = h.previewOriginals.Image;
          applyPhotoEditorEffectsWithWorker = h.previewOriginals.worker;
          commitPhotoEditorPreviewFrame = h.previewOriginals.commit;
          closePhotoEditorModal();
          h.sourceImage.onerror();
          await h.openPromise;
        });
      }
    });
  }
  assert.deepEqual(await inPage(() => window.overlaySessionSmoke.errors), []);
  assert.deepEqual([photoA, photoB].map(photo => hashFile(photo.filePath)), sourceHashes);
  report.sourceHashes = sourceHashes;
  report.ok = true;
  console.log(`Photo editor overlay session smoke passed: ${OUTPUT_DIR}`);
}

const watchdog = setTimeout(() => { console.error('Overlay session smoke exceeded 90 seconds.'); app.exit(1); }, 90000);
main().then(() => {
  fs.writeFileSync(path.join(OUTPUT_DIR, 'result.json'), JSON.stringify(report, null, 2));
  clearTimeout(watchdog);
  app.quit();
}).catch(error => {
  report.error = error.stack || String(error);
  fs.writeFileSync(path.join(OUTPUT_DIR, 'result.json'), JSON.stringify(report, null, 2));
  console.error(error);
  clearTimeout(watchdog);
  app.exit(1);
});

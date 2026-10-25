// Exercise the real save action while controlling only its asynchronous boundaries.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { pathToFileURL } = require('node:url');
const { app, BrowserWindow, nativeImage, session } = require('electron');

const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-save-session-'));
const OUTPUT_DIR = path.resolve(__dirname, '..', 'out', 'smoke-save-session', path.basename(ROOT));
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
const hashFile = filePath => createHash('sha256').update(fs.readFileSync(filePath)).digest('hex');
let win;
const inPage = (fn, ...args) => win.webContents.executeJavaScript(
  `(${fn.toString()})(...${JSON.stringify(args)})`
);

async function until(label, predicate) {
  const deadline = Date.now() + 10000;
  while (Date.now() < deadline) {
    if (await predicate()) return;
    await new Promise(resolve => setTimeout(resolve, 30));
  }
  throw new Error(`Timed out: ${label}`);
}

function createPhoto(id, width, height) {
  const pixels = Buffer.alloc(width * height * 4);
  for (let index = 0; index < pixels.length; index += 4) pixels.set([40, 100, 180, 255], index);
  const filePath = path.join(ROOT, `source-${id}.png`);
  fs.writeFileSync(filePath, nativeImage.createFromBitmap(pixels, { width, height }).toPNG());
  return {
    id, filePath, fileName: path.basename(filePath), fileUrl: pathToFileURL(filePath).href,
    takenAtTimestamp: 1780790400000 + id, worldId: `wrld-source-${id}`,
    worldUrl: `https://vrchat.com/home/world/wrld-source-${id}`,
    worldName: `Effective world ${id}`, rawWorldName: `Automatic world ${id}`,
    worldNameManual: `Manual world ${id}`, worldIdentitySource: 'manual',
    memoText: `Memo ${id}`, printNoteText: `Note ${id}`, isFavorite: true,
  };
}

async function openPhoto(photo) {
  await inPage(async value => {
    currentModalPhoto = value;
    await openPhotoEditorModal();
  }, photo);
  await until('source image', () => inPage(() => Boolean(photoEditorState?.sourceImage)));
}

async function startSave() {
  return inPage(() => window.saveSessionSmoke.start());
}

async function waitStage(job, stage) {
  await until(`${stage} boundary for job ${job}`, () => inPage((id, name) =>
    window.saveSessionSmoke.gates.some(gate => gate.jobId === id && gate.stage === name && !gate.settled), job, stage));
}

async function release(job, stage, outcome = 'success') {
  await inPage((id, name, result) => window.saveSessionSmoke.release(id, name, result), job, stage, outcome);
}

async function finish(job) {
  const outcome = await inPage(async id => {
    const job = window.saveSessionSmoke.jobs[id];
    await job.promise;
    return job.outcome;
  }, job);
  assert.equal(outcome, 'resolved', 'The save action must handle its own asynchronous failures');
}

async function finishWithoutChange(job, stage, outcome) {
  const result = await inPage(async (id, name, value) => {
    const h = window.saveSessionSmoke;
    const before = h.capture();
    h.release(id, name, value);
    await h.jobs[id].promise;
    return { before, after: h.capture(), outcome: h.jobs[id].outcome };
  }, job, stage, outcome);
  assert.equal(result.outcome, 'resolved');
  assert.deepEqual(result.after, result.before, 'The stale save must not modify the current editor');
}

async function finishSuccessfulPersistenceWithoutChange(job) {
  const result = await inPage(async id => {
    const h = window.saveSessionSmoke;
    const before = h.capture();
    h.release(id, 'persist', 'success');
    // Stay in one renderer task so a preview frame cannot alter the snapshot.
    for (let attempt = 0; attempt < 20; attempt += 1) {
      if (h.gates.some(gate => gate.jobId === id && gate.stage === 'sync')) break;
      await Promise.resolve();
    }
    h.release(id, 'sync', 'success');
    await h.jobs[id].promise;
    return { before, after: h.capture(), outcome: h.jobs[id].outcome };
  }, job);
  assert.equal(result.outcome, 'resolved');
  assert.deepEqual(result.after, result.before, 'Persistence completion and synchronization must both preserve B');
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
    throw new Error('Network disabled in save session smoke');
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
  await until('application window', () => {
    win = BrowserWindow.getAllWindows()[0];
    return Boolean(win);
  });
  await until('renderer startup', () => inPage(() =>
    typeof photoEditorSaveController !== 'undefined' &&
    typeof persistPhotoEditorExport === 'function' &&
    typeof synchronizePhotoEditorSavedResult === 'function' &&
    !document.querySelector('.app')?.classList.contains('is-app-initializing')
  ));
  await inPage(() => {
    const h = window.saveSessionSmoke = { jobs: [], gates: [], errors: [], syncCalls: [] };
    h.realSynchronize = synchronizePhotoEditorSavedResult;
    const copy = value => JSON.parse(JSON.stringify(value));
    window.addEventListener('unhandledrejection', event => h.errors.push(String(event.reason)));
    const originalMatchMedia = window.matchMedia.bind(window);
    window.matchMedia = query => query === '(prefers-reduced-motion: reduce)'
      ? { matches: false } : originalMatchMedia(query);
    h.capture = () => ({
      photoId: photoEditorState?.sourcePhoto?.id ?? null,
      sourcePath: photoEditorState?.sourcePhoto?.filePath ?? null,
      sourceSize: photoEditorState?.sourceImage
        ? [photoEditorState.sourceImage.naturalWidth, photoEditorState.sourceImage.naturalHeight] : null,
      isSaving: photoEditorState?.isSaving ?? null,
      active: photoEditorSession.isActive(photoEditorState),
      values: copy(photoEditorState?.values ?? null),
      history: photoEditorSession.getHistoryState(),
      status: document.getElementById('photo-editor-status').textContent,
      modalClass: document.getElementById('photo-editor-modal').className,
      fileName: document.getElementById('photo-editor-file-name')?.textContent ?? null,
      saveDisabled: document.getElementById('photo-editor-save-btn')?.disabled ?? null,
    });
    h.defer = (stage, jobId, data) => new Promise((resolve, reject) => {
      h.gates.push({ stage, jobId, data, resolve, reject, settled: false });
    });
    h.start = () => {
      const id = h.jobs.length;
      const job = { id, sourceId: photoEditorState.sourcePhoto.id, outcome: 'pending' };
      h.jobs.push(job);
      job.promise = savePhotoEditorImage().then(
        () => { job.outcome = 'resolved'; }, error => { job.outcome = String(error); }
      );
      return id;
    };
    renderPhotoEditorExportDataUrl = (_canvas, settings, options) => {
      const job = h.jobs[h.jobs.length - 1];
      job.dataUrl = `data:image/png;base64,${btoa(`save-session-${job.id}`)}`;
      return h.defer('render', job.id, { settings: copy(settings), transparent: Boolean(options?.transparentSubjectBackground) });
    };
    persistPhotoEditorExport = payload => {
      const job = h.jobs.find(value => value.dataUrl === payload.dataUrl);
      if (!job) throw new Error('Persistence did not receive the captured render result');
      return h.defer('persist', job.id, copy(payload));
    };
    synchronizePhotoEditorSavedResult = (result, { isCurrent }) => {
      const call = { jobId: result.smokeJobId, currentAtStart: isCurrent(), currentAtEnd: null };
      h.syncCalls.push(call);
      return h.defer('sync', result.smokeJobId, copy(result)).then(() => {
        call.currentAtEnd = isCurrent();
      });
    };
    h.release = (jobId, stage, outcome) => {
      const gate = h.gates.find(value => value.jobId === jobId && value.stage === stage && !value.settled);
      if (!gate) throw new Error(`No pending ${stage} boundary for job ${jobId}`);
      gate.settled = true;
      if (outcome === 'reject') return gate.reject(new Error(`controlled ${stage} rejection`));
      if (stage === 'render') return gate.resolve({ dataUrl: h.jobs[jobId].dataUrl, outputSize: { width: 32, height: 24 } });
      if (stage === 'sync') return gate.resolve();
      const result = outcome === 'cancel' ? { canceled: true }
        : outcome === 'failure' ? { ok: false, message: 'controlled persistence failure' }
          : { ok: true, fileName: `saved-${jobId}.png`,
            importResult: h.jobs[jobId].importResult || { smoke: true }, smokeJobId: jobId };
      gate.resolve(result);
    };
  });

  const photoA = createPhoto(1, 32, 24);
  const photoB = createPhoto(2, 24, 32);
  const sourceHashes = [photoA, photoB].map(photo => hashFile(photo.filePath));

  await check('current save captures A metadata, suppresses double save, synchronizes and closes', async () => {
    await openPhoto(photoA);
    const job = await startSave();
    await waitStage(job, 'render');
    const second = await startSave();
    await finish(second);
    assert.equal(await inPage(() => window.saveSessionSmoke.gates.filter(gate => gate.stage === 'render').length), 1);
    assert.equal(await inPage(() => photoEditorState.isSaving), true);
    await inPage(() => {
      // Another app update may refresh metadata while rendering is pending.
      photoEditorState.sourcePhoto.worldId = 'wrld-changed-after-save-start';
      photoEditorState.sourcePhoto.memoText = 'Changed after save start';
    });
    await release(job, 'render');
    await waitStage(job, 'persist');
    const payload = await inPage(id => window.saveSessionSmoke.gates.find(gate => gate.jobId === id && gate.stage === 'persist').data, job);
    assert.equal(payload.sourceFilePath, photoA.filePath);
    assert.equal(payload.sourceFileName, photoA.fileName);
    assert.equal(payload.sourceTakenAtTimestamp, photoA.takenAtTimestamp);
    assert.equal(payload.sourcePhotoMetadata.worldId, photoA.worldId);
    assert.equal(payload.sourcePhotoMetadata.worldUrl, photoA.worldUrl);
    assert.equal(payload.sourcePhotoMetadata.worldName, photoA.rawWorldName);
    assert.equal(payload.sourcePhotoMetadata.worldNameManual, photoA.worldNameManual);
    assert.equal(payload.sourcePhotoMetadata.worldIdentitySource, 'manual');
    assert.equal(payload.sourcePhotoMetadata.memoText, photoA.memoText);
    assert.equal(payload.outputFormat, 'png');
    await release(job, 'persist');
    await waitStage(job, 'sync');
    await release(job, 'sync');
    await finish(job);
    const result = await inPage(id => ({ active: photoEditorSession.isActive(photoEditorState),
      call: window.saveSessionSmoke.syncCalls.find(call => call.jobId === id) }), job);
    assert.equal(result.active, false);
    assert.deepEqual(result.call, { jobId: job, currentAtStart: true, currentAtEnd: true });
  });

  for (const outcome of ['cancel', 'failure', 'reject']) {
    await check(`current persistence ${outcome} restores controls and reports status`, async () => {
      await openPhoto(photoA);
      const job = await startSave();
      await release(job, 'render');
      await waitStage(job, 'persist');
      await release(job, 'persist', outcome);
      await finish(job);
      const result = await inPage(() => window.saveSessionSmoke.capture());
      assert.equal(result.active, true);
      assert.equal(result.isSaving, false);
      assert.match(result.status, outcome === 'cancel' ? /キャンセル/ : outcome === 'failure' ? /controlled persistence failure/ : /controlled persist rejection/);
    });
  }

  for (const outcome of ['success', 'reject']) {
    await check(`A rendering ${outcome} after close and B open never persists or modifies B`, async () => {
      await openPhoto(photoA);
      const job = await startSave();
      await inPage(() => closePhotoEditorModal());
      await openPhoto(photoB);
      await inPage(() => setPhotoEditorStatus('B editing continues'));
      await finishWithoutChange(job, 'render', outcome);
      assert.equal(await inPage(id => window.saveSessionSmoke.gates.some(gate => gate.jobId === id && gate.stage === 'persist'), job), false);
    });
  }

  for (const outcome of ['success', 'cancel', 'failure', 'reject']) {
    await check(`A persistence ${outcome} preserves B save and only successful A synchronizes`, async () => {
      await openPhoto(photoA);
      const job = await startSave();
      await release(job, 'render');
      await waitStage(job, 'persist');
      await inPage(() => closePhotoEditorModal());
      await openPhoto(photoB);
      const bJob = await startSave();
      await waitStage(bJob, 'render');
      if (outcome === 'success') {
        await finishSuccessfulPersistenceWithoutChange(job);
      } else {
        await finishWithoutChange(job, 'persist', outcome);
      }
      const result = await inPage(id => ({ state: window.saveSessionSmoke.capture(),
        calls: window.saveSessionSmoke.syncCalls.filter(call => call.jobId === id),
        payload: window.saveSessionSmoke.gates.find(gate => gate.jobId === id && gate.stage === 'persist').data }), job);
      assert.equal(result.state.photoId, photoB.id);
      assert.equal(result.state.isSaving, true, 'A finally must not reset B saving flag');
      assert.equal(result.payload.sourceFilePath, photoA.filePath);
      assert.equal(result.payload.sourcePhotoMetadata.worldId, photoA.worldId);
      assert.deepEqual(result.calls, outcome === 'success'
        ? [{ jobId: job, currentAtStart: false, currentAtEnd: false }] : []);
      await release(bJob, 'render', 'reject');
      await finish(bJob);
    });
  }

  await check('A synchronization crossing B open cannot close B or clear B saving', async () => {
    await openPhoto(photoA);
    const job = await startSave();
    await release(job, 'render');
    await waitStage(job, 'persist');
    await release(job, 'persist');
    await waitStage(job, 'sync');
    await inPage(() => closePhotoEditorModal());
    await openPhoto(photoB);
    const bJob = await startSave();
    await waitStage(bJob, 'render');
    await finishWithoutChange(job, 'sync', 'success');
    assert.deepEqual(await inPage(id => window.saveSessionSmoke.syncCalls.find(call => call.jobId === id), job),
      { jobId: job, currentAtStart: true, currentAtEnd: false });
    await release(bJob, 'render', 'reject');
    await finish(bJob);
  });

  await check('closing and reopening the same photo creates a different save owner', async () => {
    await openPhoto(photoA);
    const job = await startSave();
    await inPage(() => closePhotoEditorModal());
    await openPhoto(photoA);
    await inPage(() => setPhotoEditorStatus('Reopened A editing continues'));
    await finishWithoutChange(job, 'render', 'success');
    assert.equal(await inPage(id => window.saveSessionSmoke.gates.some(gate => gate.jobId === id && gate.stage === 'persist'), job), false);
  });

  await check('a decoded source image delivered after close cannot replace B source', async () => {
    await inPage(photo => {
      const h = window.saveSessionSmoke;
      const OriginalImage = window.Image;
      window.Image = function (...args) {
        window.Image = OriginalImage;
        const image = new OriginalImage(...args);
        Object.defineProperty(image, 'onload', {
          set(handler) { image.addEventListener('load', event => { h.deliverSource = () => handler.call(image, event); }); },
          configurable: true,
        });
        return image;
      };
      currentModalPhoto = photo;
      h.opening = openPhotoEditorModal();
    }, photoA);
    await until('delayed source image callback', () => inPage(() => Boolean(window.saveSessionSmoke.deliverSource)));
    await inPage(() => closePhotoEditorModal());
    await openPhoto(photoB);
    const result = await inPage(async () => {
      const h = window.saveSessionSmoke;
      const before = h.capture();
      h.deliverSource();
      await h.opening;
      return { before, after: h.capture() };
    });
    assert.equal(result.before.photoId, photoB.id);
    assert.deepEqual(result.before.sourceSize, [24, 32]);
    assert.deepEqual(result.after, result.before);
  });

  await check('real synchronization rechecks ownership after sidebar refresh and preserves B month', async () => {
    const monthA = { year: 2026, month: 6 };
    const monthB = { year: 2026, month: 8 };
    const importResult = { selectedMonth: monthA, worldMetadataTargets: [{ worldId: photoA.worldId }] };
    await openPhoto(photoA);
    await inPage(selection => {
      const h = window.saveSessionSmoke;
      h.adapterOriginals = { synchronizePhotoEditorSavedResult, refreshSidebar,
        restoreMonthViewAfterDataChange, queueWorldMetadataSyncForResult };
      h.monthRestores = [];
      h.metadataQueues = [];
      h.resolveSidebar = null;
      currentSelection = selection;
      synchronizePhotoEditorSavedResult = h.realSynchronize;
      refreshSidebar = () => new Promise(resolve => { h.resolveSidebar = resolve; });
      restoreMonthViewAfterDataChange = async options => {
        h.monthRestores.push(JSON.parse(JSON.stringify(options)));
        currentSelection = options.preferredSelection;
        return true;
      };
      queueWorldMetadataSyncForResult = async result => {
        h.metadataQueues.push(JSON.parse(JSON.stringify(result)));
      };
    }, monthA);
    try {
      const job = await startSave();
      await inPage((id, result) => { window.saveSessionSmoke.jobs[id].importResult = result; }, job, importResult);
      await release(job, 'render');
      await waitStage(job, 'persist');
      await release(job, 'persist');
      await until('real synchronization sidebar refresh', () => inPage(() => Boolean(window.saveSessionSmoke.resolveSidebar)));
      assert.equal(await inPage(() => photoEditorSession.isActive(photoEditorState)), true,
        'A must still be current when synchronization enters the sidebar await');
      await inPage(() => closePhotoEditorModal());
      await openPhoto(photoB);
      await inPage(selection => {
        currentSelection = selection;
        setPhotoEditorStatus('B is viewing its own month');
      }, monthB);
      const result = await inPage(async id => {
        const h = window.saveSessionSmoke;
        const before = h.capture();
        h.resolveSidebar();
        await h.jobs[id].promise;
        return { before, after: h.capture(), outcome: h.jobs[id].outcome,
          selection: currentSelection, restores: h.monthRestores, metadataQueues: h.metadataQueues };
      }, job);
      assert.equal(result.outcome, 'resolved');
      assert.deepEqual(result.selection, monthB, 'A completion must not move the library back to A month');
      assert.equal(result.restores.length, 1);
      assert.deepEqual(result.restores[0].preferredSelection, monthB);
      assert.deepEqual(result.metadataQueues, [importResult], 'A saved file must still queue its World metadata');
      assert.deepEqual(result.after, result.before, 'The real synchronization adapter must preserve B editor state');
    } finally {
      await inPage(() => {
        const original = window.saveSessionSmoke.adapterOriginals;
        synchronizePhotoEditorSavedResult = original.synchronizePhotoEditorSavedResult;
        refreshSidebar = original.refreshSidebar;
        restoreMonthViewAfterDataChange = original.restoreMonthViewAfterDataChange;
        queueWorldMetadataSyncForResult = original.queueWorldMetadataSyncForResult;
      });
    }
  });

  assert.deepEqual(await inPage(() => window.saveSessionSmoke.errors), []);
  assert.deepEqual([photoA, photoB].map(photo => hashFile(photo.filePath)), sourceHashes);
  report.sourceHashes = sourceHashes;
  report.ok = true;
  console.log(`Photo editor save session smoke passed: ${OUTPUT_DIR}`);
}

const watchdog = setTimeout(() => {
  console.error('Save session smoke exceeded 90 seconds.');
  app.exit(1);
}, 90000);
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

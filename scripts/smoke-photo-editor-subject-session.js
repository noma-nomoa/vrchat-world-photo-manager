const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { app, BrowserWindow, nativeImage, session } = require('electron');

// Load the real renderer and control only the asynchronous model boundary.
// No model download, user database, or user photo is used by this smoke test.
const root = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-subject-session-'));
for (const name of ['home', 'appData', 'documents', 'pictures', 'temp']) {
  const directory = path.join(root, name);
  fs.mkdirSync(directory, { recursive: true });
  app.setPath(name, directory);
}
process.env.TEMP = process.env.TMP = app.getPath('temp');
process.env.APPDATA = app.getPath('appData');
process.env.LOCALAPPDATA = path.join(root, 'local');
process.env.WORLDSHOT_APPDATA_ROOT = process.env.APPDATA;
process.env.WORLDSHOT_HOME_ROOT = app.getPath('home');
process.env.WORLDSHOT_LOCALAPPDATA_ROOT = process.env.LOCALAPPDATA;
fs.mkdirSync(process.env.LOCALAPPDATA, { recursive: true });
app.setPath('userData', path.join(process.env.APPDATA, 'WorldShot Log'));

const report = { ok: false, root, checks: [] };
const outputDir = path.resolve(__dirname, '..', 'out', 'smoke-subject-session', path.basename(root));
fs.mkdirSync(outputDir, { recursive: true });
const baselineOnly = process.argv.includes('--baseline-only');
let win;
const deadline = setTimeout(() => {
  console.error('Subject session smoke exceeded 90 seconds.');
  app.exit(1);
}, 90000);

function inPage(fn, ...args) {
  return win.webContents.executeJavaScript(`(${fn.toString()})(...${JSON.stringify(args)})`);
}

async function waitFor(predicate, label) {
  const started = Date.now();
  while (Date.now() - started < 10000) {
    if (await predicate()) return;
    await new Promise((resolve) => setTimeout(resolve, 30));
  }
  throw new Error(`Timed out: ${label}`);
}

function createPhoto(id, width, height) {
  const pixels = Buffer.alloc(width * height * 4);
  for (let i = 0; i < pixels.length; i += 4) {
    pixels[i] = 40;
    pixels[i + 1] = 100;
    pixels[i + 2] = 180;
    pixels[i + 3] = 255;
  }
  const filePath = path.join(root, `photo-${id}.png`);
  fs.writeFileSync(filePath, nativeImage.createFromBitmap(pixels, { width, height }).toPNG());
  return { id, fileName: path.basename(filePath), filePath, fileUrl: pathToFileURL(filePath).href };
}

async function openPhoto(photo) {
  await inPage(async (value) => {
    currentModalPhoto = value;
    await openPhotoEditorModal();
  }, photo);
  await waitFor(() => inPage(() => Boolean(photoEditorState?.sourceImage)), 'source image');
}

async function start(modelId) {
  return inPage((id) => {
    const h = window.subjectSessionSmoke;
    const request = { outcome: 'pending' };
    h.jobs.push(request);
    request.promise = generatePhotoEditorSubjectMaskWithModel(id, id === 'u2netp' ? 'lightweight' : 'standard')
      .then(() => { request.outcome = 'resolved'; }, (error) => { request.outcome = error.message; });
    return h.jobs.length - 1;
  }, modelId);
}

async function settle(job, fail = false) {
  return inPage(async (index, reject) => {
    const h = window.subjectSessionSmoke;
    const pending = h.pending.shift();
    if (!pending) throw new Error('Expected an outstanding model request.');
    if (reject) pending.reject(new Error('controlled model failure'));
    else pending.resolve();
    await h.jobs[index].promise;
    return h.jobs[index].outcome;
  }, job, fail);
}

async function snapshot() {
  return inPage(() => window.subjectSessionSmoke.capture());
}

async function assertMaskControlsEnabled() {
  const disabled = await inPage(() => [
    document.getElementById('photo-editor-subject-auto-btn').disabled,
    document.getElementById('photo-editor-subject-high-quality-btn').disabled,
  ]);
  assert.deepEqual(disabled, [false, false], 'AI buttons must be enabled for another attempt');
}

async function settleWithoutChange(job, fail = false) {
  // Snapshot and finish in one renderer task. Unrelated rAF preview updates
  // between IPC calls must not be mistaken for writes by the stale request.
  const result = await inPage(async (index, reject) => {
    const h = window.subjectSessionSmoke;
    const before = JSON.stringify(h.capture());
    const pending = h.pending.shift();
    if (reject) pending.reject(new Error('controlled model failure'));
    else pending.resolve();
    await h.jobs[index].promise;
    return { before, after: JSON.stringify(h.capture()), outcome: h.jobs[index].outcome };
  }, job, fail);
  assert.equal(result.outcome, 'resolved');
  assert.equal(result.after, result.before);
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
    { urls: ['http://*/*', 'https://*/*'] }, (_details, callback) => callback({ cancel: true })
  );
  app.on('browser-window-created', (_event, window) => window.hide());
  require('../src/index');
  await waitFor(() => {
    win = BrowserWindow.getAllWindows()[0];
    return Boolean(win);
  }, 'application window');
  await waitFor(() => inPage(() =>
    typeof openPhotoEditorModal === 'function' &&
    !document.querySelector('.app')?.classList.contains('is-app-initializing')
  ), 'renderer startup');

  await inPage(() => {
    const h = window.subjectSessionSmoke = { pending: [], jobs: [], errors: [], refreshWaiter: null };
    h.capture = () => {
      const history = photoEditorSession.getHistoryState();
      return {
        photoId: photoEditorState?.sourcePhoto?.id || null,
        mask: photoEditorState?.subjectMask || null,
        undo: history.undo,
        redo: history.redo,
        pending: history.pending,
        status: document.getElementById('photo-editor-status').textContent,
      };
    };
    window.addEventListener('unhandledrejection', (event) => h.errors.push(String(event.reason)));
    // Ensure the close-animation interval is exercised even on reduced-motion desktops.
    const originalMatchMedia = window.matchMedia.bind(window);
    window.matchMedia = (query) => query === '(prefers-reduced-motion: reduce)'
      ? { matches: false } : originalMatchMedia(query);
    h.models = [
      { id: 'u2netp', label: 'Lightweight', ready: true, inputSize: 32 },
      { id: 'withoutbg-snap', label: 'Snap', ready: true, pipeline: 'withoutbg-snap' },
      { id: 'withoutbg-focus', label: 'Focus', ready: true, pipeline: 'withoutbg-focus' },
    ];
    photoEditorAiSubjectModelStatus = { ok: true, models: h.models };
    refreshPhotoEditorAiSubjectModelStatus = async () => {
      if (h.refreshWaiter) await h.refreshWaiter.promise;
      photoEditorAiSubjectModelStatus = { ok: true, models: h.models };
      return photoEditorAiSubjectModelStatus;
    };
    const mask = document.createElement('canvas');
    mask.width = 32;
    mask.height = 24;
    mask.getContext('2d').fillRect(8, 4, 16, 16);
    h.maskDataUrl = mask.toDataURL('image/png');
    runPhotoEditorSubjectModel = () => new Promise((resolve, reject) => h.pending.push({
      reject,
      resolve: () => {
        const data = new Float32Array(32 * 32);
        for (let y = 8; y < 24; y++) for (let x = 8; x < 24; x++) data[y * 32 + x] = 1;
        resolve({ data, dims: [1, 1, 32, 32], type: 'float32' });
      },
    }));
    createPhotoEditorWithoutBgSubjectMaskDataUrl = () =>
      new Promise((resolve, reject) => h.pending.push({ reject, resolve: () => resolve(h.maskDataUrl) }));
  });

  const photoA = createPhoto(1, 32, 24);
  const photoB = createPhoto(2, 24, 32);
  if (!baselineOnly) {
    await check('closing A discards its history timer without committing B pending edits', async () => {
      await openPhoto(photoA);
      const result = await inPage(async (nextPhoto) => {
        updatePhotoEditorSliderValue('brightness', 12);
        schedulePhotoEditorHistoryCommit(800);
        closePhotoEditorModal();
        currentModalPhoto = nextPhoto;
        await openPhotoEditorModal();
        updatePhotoEditorSliderValue('brightness', 27);
        schedulePhotoEditorHistoryCommit(4000);
        const before = photoEditorSession.getHistoryState();
        // Pass both A's old deadline and its close animation, but not B's commit deadline.
        await new Promise(resolve => setTimeout(resolve, 1100));
        return { before, after: photoEditorSession.getHistoryState(),
          photoId: photoEditorState?.sourcePhoto?.id,
          brightness: photoEditorState?.values.brightness };
      }, photoB);
      assert.ok(result.before.pending, 'B must have an uncommitted edit');
      assert.deepEqual(result.after, result.before, 'A must not commit or clear B history');
      assert.equal(result.photoId, photoB.id);
      assert.equal(result.brightness, 27);
    });
  }
  for (const model of ['u2netp', 'withoutbg-snap', 'withoutbg-focus']) {
    await check(`${model}: current result and retry after failure`, async () => {
      await openPhoto(photoA);
      assert.equal(await settle(await start(model), true), 'resolved');
      assert.equal((await snapshot()).mask.status, 'failed');
      await assertMaskControlsEnabled();
      assert.equal(await settle(await start(model)), 'resolved');
      const after = await snapshot();
      assert.equal(after.mask.status, 'ready');
      assert.equal(after.mask.modelId, model);
      assert.equal(after.mask.width, 32);
      assert.equal(after.mask.height, 24);
      assert.ok(after.mask.maskDataUrl.startsWith('data:image/png;base64,'));
    });
    if (baselineOnly) continue;

    await check(`${model}: brightness and crop edits preserve pending inference`, async () => {
      await openPhoto(photoA);
      const job = await start(model);
      await inPage(() => {
        updatePhotoEditorSliderValue('brightness', 12);
        updatePhotoEditorCrop({ zoom: 30, offsetX: 12 });
        commitPhotoEditorHistoryMutation();
      });
      assert.equal(await settle(job), 'resolved');
      assert.equal((await snapshot()).mask.status, 'ready');
      const values = await inPage(() => ({ brightness: photoEditorState.values.brightness, crop: photoEditorState.crop }));
      assert.equal(values.brightness, 12);
      assert.equal(values.crop.zoom, 30);
      assert.equal(values.crop.offsetX, 12);
    });

    for (const fail of [false, true]) {
      await check(`${model}: close animation ignores ${fail ? 'failure' : 'success'}`, async () => {
        await openPhoto(photoA);
        const job = await start(model);
        // Complete the model inside the same renderer task as closing so this is
        // guaranteed to run before the 520ms close finalizer, not after a sleep.
        const unchanged = await inPage(async (index, reject) => {
          const h = window.subjectSessionSmoke;
          clearPhotoEditorHistoryCommitTimer();
          document.getElementById('photo-editor-close').click();
          const oldState = photoEditorState;
          const before = JSON.stringify(capturePhotoEditorHistorySnapshot());
          const status = document.getElementById('photo-editor-status').textContent;
          const request = h.pending.shift();
          if (reject) request.reject(new Error('controlled model failure'));
          else request.resolve();
          await h.jobs[index].promise;
          return photoEditorState === oldState &&
            before === JSON.stringify(capturePhotoEditorHistorySnapshot()) &&
            status === document.getElementById('photo-editor-status').textContent &&
            h.jobs[index].outcome === 'resolved';
        }, job, fail);
        assert.equal(unchanged, true);
        await waitFor(() => inPage(() => photoEditorState === null), 'close completed');
      });

      await check(`${model}: old ${fail ? 'failure' : 'success'} cannot modify another photo`, async () => {
        await openPhoto(photoA);
        const job = await start(model);
        await inPage(() => closePhotoEditorModal());
        await openPhoto(photoB);
        await inPage(() => setPhotoEditorStatus('photo B status'));
        await settleWithoutChange(job, fail);
      });
    }

    await check(`${model}: result after complete close is harmless`, async () => {
      await openPhoto(photoA);
      const job = await start(model);
      await inPage(() => closePhotoEditorModal());
      await waitFor(() => inPage(() => photoEditorState === null), 'closed editor');
      assert.equal(await settle(job), 'resolved');
      assert.equal((await snapshot()).mask, null);
    });

    for (const fail of [false, true]) {
      await check(`${model}: newer request survives an older ${fail ? 'failure' : 'success'}`, async () => {
        await openPhoto(photoA);
        const first = await start(model);
        const second = await start(model);
        await inPage(() => { window.subjectSessionSmoke.pending.reverse(); });
        assert.equal(await settle(second), 'resolved');
        await settleWithoutChange(first, fail);
      });
    }
    for (const fail of [false, true]) {
      await check(`${model}: cancelled ${fail ? 'failure' : 'success'} preserves pending replacement`, async () => {
        await openPhoto(photoA);
        const first = await start(model);
        const second = await start(model);
        await settleWithoutChange(first, fail);
        assert.equal(await settle(second), 'resolved');
        assert.equal((await snapshot()).mask.status, 'ready');
      });
    }

    await check(`${model}: undo and redo completed inference cannot restore loading`, async () => {
      await openPhoto(photoA);
      const job = await start(model);
      // The normal debounce timer commits this loading snapshot during slow AI.
      await inPage(() => commitPhotoEditorHistoryMutation());
      assert.equal(await settle(job), 'resolved');
      const readyMask = (await snapshot()).mask;
      await inPage(() => undoPhotoEditorEdit());
      assert.notEqual((await snapshot()).mask.status, 'loading');
      await assertMaskControlsEnabled();
      await inPage(() => redoPhotoEditorEdit());
      assert.deepEqual((await snapshot()).mask, readyMask);
    });

    for (const action of ['clearPhotoEditorSubjectMask', 'resetPhotoEditorAll', 'undoPhotoEditorEdit']) {
      await check(`${model}: ${action} cancels pending mask`, async () => {
        await openPhoto(photoA);
        const job = await start(model);
        await inPage((name) => window[name](), action);
        await settleWithoutChange(job);
        if (action === 'undoPhotoEditorEdit') {
          await inPage(() => redoPhotoEditorEdit());
          assert.notEqual((await snapshot()).mask.status, 'loading');
          await assertMaskControlsEnabled();
        }
        assert.equal(await settle(await start(model)), 'resolved');
        assert.equal((await snapshot()).mask.status, 'ready');
      });
    }
  }

  if (!baselineOnly) {
    // FileReader and Image are controlled only while creating an import. Real
    // image decoding/serialization is exercised separately by smoke:export.
    const beginImport = () => inPage(() => {
      const h = window.subjectSessionSmoke;
      const OriginalReader = window.FileReader;
      window.FileReader = class {
        readAsDataURL() { h.reader = this; }
      };
      try { loadPhotoEditorSubjectMaskFile({ type: 'image/png' }); }
      finally { window.FileReader = OriginalReader; }
    });
    const beginDecode = () => inPage(() => {
      const h = window.subjectSessionSmoke;
      const OriginalImage = window.Image;
      window.Image = class {
        constructor() { h.image = this; this.naturalWidth = 32; this.naturalHeight = 24; }
        set src(_value) {}
      };
      try { h.reader.result = h.maskDataUrl; h.reader.onload(); }
      finally { window.Image = OriginalImage; }
    });
    for (const stage of ['read', 'decode']) {
      for (const fail of [false, true]) {
        await check(`import ${stage}: stale ${fail ? 'failure' : 'success'} cannot change photo B`, async () => {
          await openPhoto(photoA);
          await beginImport();
          if (stage === 'decode') await beginDecode();
          await inPage(() => closePhotoEditorModal());
          await openPhoto(photoB);
          const unchanged = await inPage((step, reject) => {
            const h = window.subjectSessionSmoke;
            const before = JSON.stringify(h.capture());
            const target = step === 'read' ? h.reader : h.image;
            if (reject) target.onerror();
            else { target.result = h.maskDataUrl; target.onload(); }
            return before === JSON.stringify(h.capture());
          }, stage, fail);
          assert.equal(unchanged, true);
        });
      }
    }
    await check('mask import supersedes pending inference', async () => {
      await openPhoto(photoA);
      const job = await start('u2netp');
      await beginImport();
      await beginDecode();
      await inPage(() => window.subjectSessionSmoke.image.onload());
      assert.equal((await snapshot()).mask.source, 'imported');
      await settleWithoutChange(job);
    });
    await check('new inference supersedes pending mask import', async () => {
      await openPhoto(photoA);
      await beginImport();
      await beginDecode();
      assert.equal(await settle(await start('u2netp')), 'resolved');
      const unchanged = await inPage(() => {
        const h = window.subjectSessionSmoke;
        const before = JSON.stringify(h.capture());
        h.image.onload();
        return before === JSON.stringify(h.capture());
      });
      assert.equal(unchanged, true);
    });
    for (const stage of ['read', 'decode']) {
      await check(`failed mask ${stage} releases cancelled inference controls`, async () => {
        await openPhoto(photoA);
        const job = await start('u2netp');
        await beginImport();
        if (stage === 'decode') await beginDecode();
        await inPage((step) => {
          const h = window.subjectSessionSmoke;
          (step === 'read' ? h.reader : h.image).onerror();
        }, stage);
        assert.notEqual((await snapshot()).mask.status, 'loading');
        await assertMaskControlsEnabled();
        await settleWithoutChange(job, true);
        assert.equal(await settle(await start('u2netp')), 'resolved');
        assert.equal((await snapshot()).mask.status, 'ready');
      });
    }
    await check('model discovery finishing after close cannot start inference', async () => {
      await openPhoto(photoA);
      await inPage(() => {
        photoEditorAiSubjectModelStatus = { ok: true, models: [] };
        const h = window.subjectSessionSmoke;
        h.refreshWaiter = {};
        h.refreshWaiter.promise = new Promise((resolve) => { h.refreshWaiter.resolve = resolve; });
      });
      const job = await start('u2netp');
      await inPage(() => closePhotoEditorModal());
      await inPage(() => {
        window.subjectSessionSmoke.refreshWaiter.resolve();
        window.subjectSessionSmoke.refreshWaiter = null;
      });
      await waitFor(() => inPage((index) => window.subjectSessionSmoke.jobs[index].outcome !== 'pending', job), 'cancelled discovery');
      assert.equal(await inPage(() => window.subjectSessionSmoke.pending.length), 0);
    });
  }
  assert.deepEqual(await inPage(() => window.subjectSessionSmoke.errors), []);
  report.ok = true;
}

main().catch((error) => {
  console.error(error);
}).finally(() => {
  clearTimeout(deadline);
  fs.writeFileSync(path.join(outputDir, 'result.json'), JSON.stringify(report, null, 2));
  console.log(`Subject session report: ${path.join(outputDir, 'result.json')}`);
  app.exit(report.ok ? 0 : 1);
});

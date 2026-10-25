// Exercise the actual packaged EXE, ASAR imports, preload IPC and classic Worker.
// Invoke through Node so the GUI process is awaited and cleaned up on failures.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const os = require('node:os');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { spawn } = require('node:child_process');
const asar = require('@electron/asar');
const { getCurrentFuseWire, FuseV1Options, FuseVersion } = require('@electron/fuses');
const { EXPECTED: EXPECTED_RUNTIME, PROVENANCE } = require('./check-runtime-security');

const REPO_ROOT = path.resolve(__dirname, '..');
const DEFAULT_EXE_PATH = path.join(REPO_ROOT, 'out', 'WorldShot Log-win32-x64', 'WorldShotLog.exe');
const APP_DISPLAY_NAME = 'WorldShot Log';
const RUNTIME_FILES = [
  'src/index.js', 'src/index.html', 'src/preload.js', 'src/db.js', 'src/renderer.js',
  'src/app-data-backup-service.js', 'src/photo-editor-session.js',
  'src/photo-editor-save.js', 'src/photo-editor-color-core.js', 'src/photo-editor-worker.js',
  'src/photo-list-query.js', 'src/photo-metadata.js', 'src/photo-record.js',
  'src/photo-import-service.js', 'src/concurrency.js',
  'src/ai-subject-model-catalog.js', 'src/ai-subject-model-store.js',
  'src/photo-editor-image-resources.js', 'src/thumbnail-store.js',
];
const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
const sha256 = value => createHash('sha256').update(value).digest('hex');

function getArgValue(name) {
  const prefix = `${name}=`;
  return process.argv.slice(2).find(arg => arg.startsWith(prefix))?.slice(prefix.length) || '';
}

function inspectArchive(exePath) {
  const archivePath = path.join(path.dirname(exePath), 'resources', 'app.asar');
  assert(fs.existsSync(archivePath), `Missing ASAR: ${archivePath}`);
  const entries = asar.listPackage(archivePath).map(entry => entry.replaceAll('\\', '/'));
  const excludedRoots = ['scripts', 'docs', 'out', 'test_pic', '.github'];
  for (const root of excludedRoots) {
    assert(!entries.some(entry => entry === `/${root}` || entry.startsWith(`/${root}/`)),
      `${root} must not ship in the application archive`);
  }
  const runtimeFiles = RUNTIME_FILES.map(file => {
    const packed = asar.extractFile(archivePath, path.normalize(file));
    const checkout = fs.readFileSync(path.join(REPO_ROOT, file));
    // Git's Windows checkout may convert LF/CRLF. Compare these known UTF-8
    // source files canonically; binary payloads keep exact byte/hash checks.
    assert.equal(packed.toString('utf8').replace(/\r\n/g, '\n'),
      checkout.toString('utf8').replace(/\r\n/g, '\n'), `Packaged source is missing or stale: ${file}`);
    return { file, bytes: packed.length, sha256: sha256(packed),
      checkoutSha256: sha256(checkout), comparison: 'UTF-8 with CRLF normalized to LF' };
  });
  // USERPROFILE does not override Windows known-folder home paths. Refuse old
  // builds that cannot isolate thumbnail creation and legacy cache migration.
  assert(asar.extractFile(archivePath, 'src/index.js').toString().includes('WORLDSHOT_HOME_ROOT'),
    'Build lacks WORLDSHOT_HOME_ROOT isolation; refusing to launch against real home folders');
  const nativeEntry = entries.find(entry => entry.endsWith('/better-sqlite3/prebuilds/win32-x64.node'));
  assert(nativeEntry, 'Packaged SQLite native module is missing');
  assert(asar.statFile(archivePath, path.normalize(nativeEntry.slice(1))).unpacked, 'SQLite must be unpacked from ASAR');
  assert(fs.existsSync(path.join(`${archivePath}.unpacked`, nativeEntry.slice(1))), 'Unpacked SQLite binary is missing');
  const resourceFiles = ['models/u2netp/model.onnx', 'onnxruntime-web/ort-wasm-simd-threaded.wasm'];
  for (const file of resourceFiles) {
    assert(fs.statSync(path.join(path.dirname(archivePath), file)).size > 0, `Missing offline AI resource: ${file}`);
  }
  return { archivePath, runtimeFiles, nativeEntry, resourceFiles, excludedRoots };
}

async function openProtocol(url) {
  assert(/^ws:\/\/127\.0\.0\.1:\d+\//.test(url), 'CDP must remain on loopback');
  const socket = new WebSocket(url);
  const pending = new Map();
  let nextId = 0;
  socket.addEventListener('message', event => {
    const message = JSON.parse(String(event.data));
    const request = pending.get(message.id);
    if (!request) return;
    pending.delete(message.id);
    clearTimeout(request.timer);
    if (message.error) request.reject(new Error(JSON.stringify(message.error)));
    else request.resolve(message.result);
  });
  socket.addEventListener('close', () => {
    for (const request of pending.values()) {
      clearTimeout(request.timer);
      request.reject(new Error('Packaged app closed its debugging connection'));
    }
    pending.clear();
  });
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => { socket.close(); reject(new Error('CDP connection timed out')); }, 10000);
    socket.addEventListener('open', () => { clearTimeout(timer); resolve(); }, { once: true });
    socket.addEventListener('error', () => { clearTimeout(timer); reject(new Error('CDP connection failed')); }, { once: true });
  });
  return {
    send(method, params = {}, sessionId, timeoutMs = 60000) {
      const id = ++nextId;
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => { pending.delete(id); reject(new Error(`CDP timed out: ${method}`)); }, timeoutMs);
        pending.set(id, { resolve, reject, timer });
        socket.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
      });
    },
    close: () => socket.close(),
  };
}

async function terminateOwnedProcessTree(child) {
  if (process.platform !== 'win32') {
    child.kill('SIGKILL');
    return;
  }
  // Only the child spawned by this run, including its Chromium descendants.
  await new Promise((resolve, reject) => {
    const killer = spawn('taskkill.exe', ['/PID', String(child.pid), '/T', '/F'], {
      stdio: 'ignore', windowsHide: true,
    });
    const timer = setTimeout(() => { killer.kill(); reject(new Error('Process-tree cleanup timed out')); }, 5000);
    killer.once('error', error => { clearTimeout(timer); reject(error); });
    killer.once('exit', code => {
      clearTimeout(timer);
      if (code === 0) resolve();
      else reject(new Error(`Process-tree cleanup exited with code ${code}`));
    });
  });
}

// Serialized into the packaged application's real renderer via loopback CDP.
async function verifyPackagedRenderer() {
  const check = (condition, message) => { if (!condition) throw new Error(message); };
  for (const [name, method] of [['PhotoEditorSession', 'createController'],
    ['PhotoEditorSave', 'createController'], ['PhotoEditorColorCore', 'getPhotoEditorAdjustmentRenderParams'],
    ['PhotoListQuery', 'queryPhotos'], ['PhotoEditorImageResources', 'waitForReady']]) {
    check(typeof globalThis[name]?.[method] === 'function', `Missing runtime module: ${name}`);
  }
  check(typeof photoEditorSession?.beginRequest === 'function', 'Renderer session was not initialized');
  check(typeof photoEditorSaveController?.save === 'function', 'Renderer save controller was not initialized');
  check(typeof window.electronAPI?.getApplicationDataSummary === 'function', 'Preload IPC bridge failed');
  const summary = await window.electronAPI.getApplicationDataSummary();
  for (const key of ['photoCount', 'trackedFolderCount', 'worldCacheCount', 'tagCount']) {
    check(summary[key] === 0, `Isolated database must start empty: ${key}`);
  }
  const width = 32;
  const height = 24;
  const core = globalThis.PhotoEditorColorCore;
  const sourceCanvas = document.createElement('canvas');
  sourceCanvas.width = width;
  sourceCanvas.height = height;
  const sourceCtx = sourceCanvas.getContext('2d', { willReadFrequently: true });
  const source = sourceCtx.createImageData(width, height);
  for (let i = 0; i < width * height; i += 1) source.data.set([
    (i * 13 + 31) % 256, (i * 7 + 53) % 256, (i * 3 + 79) % 256, 255,
  ], i * 4);
  sourceCtx.putImageData(source, 0, 0);
  const comparisons = [];
  for (const mode of ['rgb', 'hsv', 'combined']) {
    const values = { brightness: 7, exposure: 13, contrast: -14, highlights: -21,
      whites: 11, shadows: 24, blacks: -9, gamma: 8, temperature: 12, tint: -7,
      saturation: 16, vibrance: 22, fade: 9, grain: 4 };
    const curve = core.getDefaultPhotoEditorCurveState();
    if (mode !== 'hsv') curve.points.rgb.master = [0.03, 0.20, 0.55, 0.84, 0.98];
    if (mode !== 'rgb') curve.points.hsv.s = [0, 0.20, 0.57, 0.82, 1];
    const expected = new ImageData(new Uint8ClampedArray(source.data), width, height);
    core.applyPhotoEditorPixelAdjustmentsToImageData(expected, width,
      core.getPhotoEditorAdjustmentRenderParams(values, curve));
    check(expected.data.some((value, i) => value !== source.data[i]), 'Fixture must exercise color adjustments');
    for (const responseType of ['bitmap', 'blob']) {
      const worker = new Worker(new URL('photo-editor-worker.js', location.href));
      let bitmap;
      try {
        const sourceBitmap = await createImageBitmap(sourceCanvas);
        const requestId = comparisons.length + 1;
        const response = await new Promise((resolve, reject) => {
          const timer = setTimeout(() => reject(new Error('Packaged Worker timed out')), 10000);
          worker.onmessage = event => { clearTimeout(timer); resolve(event.data); };
          worker.onerror = event => { clearTimeout(timer); reject(new Error(event.message || 'Packaged Worker failed')); };
          worker.postMessage({ type: 'render', requestId, payload: {
            sourceBitmap, width, height, values, curve, responseType, mimeType: 'image/png',
          } }, [sourceBitmap]);
        });
        check(response?.ok === true && response.type === 'render-result' && response.requestId === requestId,
          `Packaged Worker rejected ${mode}/${responseType}: ${response?.error || 'invalid response'}`);
        check(response.width === width && response.height === height, 'Worker dimensions changed');
        if (responseType === 'bitmap') {
          check(response.bitmap instanceof ImageBitmap, 'Worker bitmap transfer failed');
          bitmap = response.bitmap;
        } else {
          check(response.mimeType === 'image/png' && response.buffer instanceof ArrayBuffer, 'Worker PNG buffer transfer failed');
          bitmap = await createImageBitmap(new Blob([response.buffer], { type: response.mimeType }));
        }
        const actualCanvas = document.createElement('canvas');
        actualCanvas.width = width;
        actualCanvas.height = height;
        const actualCtx = actualCanvas.getContext('2d');
        actualCtx.drawImage(bitmap, 0, 0);
        const actual = actualCtx.getImageData(0, 0, width, height).data;
        let differentChannels = 0;
        for (let i = 0; i < actual.length; i += 1) if (actual[i] !== expected.data[i]) differentChannels += 1;
        check(differentChannels === 0, `${mode}/${responseType}: ${differentChannels} pixel channels differ`);
        comparisons.push({ mode, responseType, differentChannels, comparedChannels: actual.length });
      } finally {
        bitmap?.close();
        worker.terminate();
      }
    }
  }
  // Real offline model/ONNX/WASM path. Only the image fixture is synthetic;
  // no inference function, model status, tensor, IPC or mask result is mocked.
  currentModalPhoto = { id: 'packaged-ai-fixture', fileName: 'packaged-ai-fixture.png',
    fileUrl: sourceCanvas.toDataURL('image/png') };
  await openPhotoEditorModal();
  check(Boolean(photoEditorState?.sourceImage), 'AI source fixture failed to load');
  await refreshPhotoEditorAiSubjectModelStatus({ silent: true });
  const model = getPhotoEditorSubjectModel('u2netp');
  check(model?.ready, `Bundled U2NetP model is unavailable: ${model?.message}`);
  check(model.modelUrl.startsWith('file:') && model.wasmBaseUrl.startsWith('file:'),
    'Bundled AI model and WASM must load from local files');
  const startedAt = performance.now();
  await generatePhotoEditorLightweightSubjectMask();
  const inferenceMs = performance.now() - startedAt;
  const mask = photoEditorState.subjectMask;
  check(mask.enabled && mask.status === 'ready', `Real U2NetP inference failed: ${mask.errorMessage}`);
  check(mask.modelId === 'u2netp' && mask.source === 'lightweight', 'Wrong model generated the mask');
  check(mask.width === width && mask.height === height, 'AI mask dimensions differ from source');
  check(mask.maskDataUrl.startsWith('data:image/png;base64,'), 'AI mask is not a PNG');
  const maskImage = new Image();
  maskImage.src = mask.maskDataUrl;
  await maskImage.decode();
  check(maskImage.naturalWidth === width && maskImage.naturalHeight === height, 'Decoded AI mask dimensions differ');
  const maskDataUrl = mask.maskDataUrl;
  undoPhotoEditorEdit();
  check(!photoEditorState.subjectMask.maskDataUrl && photoEditorState.subjectMask.status !== 'loading',
    'Undo did not remove generated AI mask');
  redoPhotoEditorEdit();
  check(photoEditorState.subjectMask.maskDataUrl === maskDataUrl && photoEditorState.subjectMask.status === 'ready',
    'Redo did not restore the generated AI mask exactly');
  const ai = { modelId: model.id, modelUrl: model.modelUrl, wasmBaseUrl: model.wasmBaseUrl,
    inferenceMs, width, height, maskDataUrlLength: maskDataUrl.length, undoRedoPassed: true };
  closePhotoEditorModal();
  return { url: location.href, title: document.title, summary, comparisons, ai };
}

async function main() {
  const exePath = path.resolve(getArgValue('--exe') || DEFAULT_EXE_PATH);
  const waitMs = Number(getArgValue('--wait-ms') || 30000);
  assert(Number.isFinite(waitMs) && waitMs >= 1000 && waitMs <= 120000, 'Invalid --wait-ms');
  assert(fs.existsSync(exePath), `Packaged app was not found: ${exePath}`);
  const archive = inspectArchive(exePath);
  const fuseWire = await getCurrentFuseWire(exePath);
  assert.equal(fuseWire.version, FuseVersion.V1, 'Unexpected packaged fuse version');
  archive.fuses = {};
  for (const [name, enabled] of Object.entries({ RunAsNode: false, EnableCookieEncryption: true,
    EnableNodeOptionsEnvironmentVariable: false, EnableNodeCliInspectArguments: false,
    EnableEmbeddedAsarIntegrityValidation: true, OnlyLoadAppFromAsar: true })) {
    // The fuse wire stores ASCII 0/1 bytes, as returned by getCurrentFuseWire.
    assert.equal(String.fromCharCode(fuseWire[FuseV1Options[name]]), enabled ? '1' : '0', `Unexpected ${name} fuse`);
    archive.fuses[name] = enabled;
  }
  const runRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-packaged-smoke-'));
  const outputDir = path.join(REPO_ROOT, 'out', 'smoke-packaged', path.basename(runRoot));
  const profile = path.join(runRoot, 'Profile');
  const appData = path.join(profile, 'AppData', 'Roaming');
  const localAppData = path.join(profile, 'AppData', 'Local');
  const temp = path.join(runRoot, 'Temp');
  for (const directory of [outputDir, appData, localAppData, temp,
    ...['Pictures', 'Documents', 'Desktop', 'Downloads'].map(name => path.join(profile, name))]) fs.mkdirSync(directory, { recursive: true });
  const blockedRequests = [];
  const proxySockets = new Set();
  const proxy = http.createServer((request, response) => {
    blockedRequests.push(request.url); response.writeHead(502); response.end('Network disabled in packaged smoke');
  });
  proxy.on('connection', socket => {
    proxySockets.add(socket);
    socket.once('close', () => proxySockets.delete(socket));
    socket.on('error', () => {});
  });
  proxy.on('connect', (request, socket) => { blockedRequests.push(request.url); socket.end('HTTP/1.1 502 Bad Gateway\r\n\r\n'); });
  await new Promise(resolve => proxy.listen(0, '127.0.0.1', resolve));
  const proxyUrl = `http://127.0.0.1:${proxy.address().port}`;
  const stderrPath = path.join(outputDir, 'stderr.log');
  const stdoutFd = fs.openSync(path.join(outputDir, 'stdout.log'), 'w');
  const stderrFd = fs.openSync(stderrPath, 'w');
  const child = spawn(exePath, ['--remote-debugging-port=0', '--remote-debugging-address=127.0.0.1',
    `--proxy-server=${proxyUrl}`, '--proxy-bypass-list=<-loopback>',
    '--host-resolver-rules=MAP * ~NOTFOUND, EXCLUDE localhost', '--disable-background-timer-throttling'], {
    env: { ...process.env, APPDATA: appData, LOCALAPPDATA: localAppData,
      USERPROFILE: profile, HOME: profile, TEMP: temp, TMP: temp,
      WORLDSHOT_HOME_ROOT: profile, WORLDSHOT_APPDATA_ROOT: appData, WORLDSHOT_LOCALAPPDATA_ROOT: localAppData,
      HTTP_PROXY: proxyUrl, HTTPS_PROXY: proxyUrl, ALL_PROXY: proxyUrl, NODE_USE_ENV_PROXY: '1', NO_PROXY: '' },
    stdio: ['ignore', stdoutFd, stderrFd], windowsHide: true,
  });
  fs.closeSync(stdoutFd);
  fs.closeSync(stderrFd);
  let exitInfo = null;
  let spawnError = null;
  child.once('error', error => { spawnError = error; });
  const childExited = new Promise(resolve => child.once('exit', (code, signal) => {
    exitInfo = { code, signal }; resolve();
  }));
  let protocol;
  let sessionId;
  let report = { ok: false, exePath, runRoot, archive, blockedRequests };
  async function until(label, predicate) {
    const deadline = Date.now() + waitMs;
    while (Date.now() < deadline) {
      if (spawnError) throw spawnError;
      if (exitInfo) throw new Error(`Packaged app exited before ${label}: ${JSON.stringify(exitInfo)}`);
      const result = await predicate();
      if (result) return result;
      await wait(100);
    }
    throw new Error(`Timed out waiting for ${label}`);
  }
  try {
    const url = await until('loopback debugging endpoint', () =>
      fs.readFileSync(stderrPath, 'utf8').match(/DevTools listening on (ws:\/\/127\.0\.0\.1:\d+\/\S+)/)?.[1]);
    protocol = await openProtocol(url);
    const runtime = await protocol.send('Browser.getVersion');
    assert(runtime.userAgent.includes(`Electron/${EXPECTED_RUNTIME.electron}`),
      `Packaged Electron differs from the reviewed security release: ${runtime.userAgent}`);
    assert(runtime.product.endsWith(`/${EXPECTED_RUNTIME.chrome}`),
      `Packaged Chromium differs from the reviewed security release: ${runtime.product}`);
    report.runtime = { ...runtime, provenance: PROVENANCE };
    const target = await until('packaged renderer', async () => (await protocol.send('Target.getTargets')).targetInfos
      .find(info => info.type === 'page' && decodeURIComponent(info.url).replaceAll('\\', '/').endsWith('/resources/app.asar/src/index.html')));
    ({ sessionId } = await protocol.send('Target.attachToTarget', { targetId: target.targetId, flatten: true }));
    await protocol.send('Network.enable', {}, sessionId);
    await protocol.send('Network.setBlockedURLs', { urls: ['http://*', 'https://*', 'ws://*', 'wss://*'] }, sessionId);
    await until('initialized renderer modules', async () => (await protocol.send('Runtime.evaluate', {
      expression: "document.readyState === 'complete' && typeof photoEditorSaveController !== 'undefined' && !document.querySelector('.app')?.classList.contains('is-app-initializing')", returnByValue: true,
    }, sessionId)).result?.value);
    const result = await protocol.send('Runtime.evaluate', {
      expression: `(${verifyPackagedRenderer.toString()})()`, awaitPromise: true, returnByValue: true,
    }, sessionId);
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
    const dbPath = path.join(appData, APP_DISPLAY_NAME, 'data', 'app.sqlite');
    const thumbnailRoot = path.join(profile, APP_DISPLAY_NAME, 'thumbnails');
    assert(fs.existsSync(dbPath), `Missing isolated database: ${dbPath}`);
    assert(fs.existsSync(thumbnailRoot), `Missing isolated thumbnail directory: ${thumbnailRoot}`);
    assert.equal(result.result.value.comparisons.length, 6);
    report = { ...report, ok: true, dbPath, thumbnailRoot, renderer: result.result.value };
  } catch (error) {
    report.error = error.stack || String(error);
    throw error;
  } finally {
    const cleanupErrors = [];
    let forcedTermination = false;
    if (!exitInfo && protocol && sessionId) {
      // Connection closure can race with this response; the actual process exit
      // below decides whether shutdown succeeded.
      await protocol.send('Runtime.evaluate', { expression: 'window.close()' }, sessionId, 3000).catch(() => {});
      await Promise.race([childExited, wait(3000)]);
    }
    try { protocol?.close(); } catch (error) { cleanupErrors.push(error.message); }
    if (!exitInfo && !spawnError) {
      forcedTermination = true;
      await terminateOwnedProcessTree(child).catch(error => cleanupErrors.push(error.message));
      await Promise.race([childExited, wait(5000)]);
    }
    try {
      for (const socket of proxySockets) socket.destroy();
      proxy.closeAllConnections();
      await new Promise((resolve, reject) => proxy.close(error => error ? reject(error) : resolve()));
    } catch (error) { cleanupErrors.push(error.message); }
    report.exitInfo = exitInfo;
    report.spawnError = spawnError?.message;
    report.forcedTermination = forcedTermination;
    report.cleanupErrors = cleanupErrors;
    if (forcedTermination || spawnError || !exitInfo || exitInfo.code !== 0 || exitInfo.signal || cleanupErrors.length) {
      report.ok = false;
      report.error ||= forcedTermination
        ? 'Packaged process did not exit normally; its owned process tree required termination'
        : `Packaged process shutdown failed: ${JSON.stringify({ exitInfo, spawnError: spawnError?.message, cleanupErrors })}`;
    }
    fs.writeFileSync(path.join(outputDir, 'result.json'), JSON.stringify(report, null, 2));
    console.log(`Packaged smoke artifacts: ${outputDir}`);
    assert(report.ok, report.error || 'Packaged smoke failed');
  }
  console.log(`Packaged app smoke passed: ${RUNTIME_FILES.length} runtime files, SQLite IPC, 6 real Worker comparisons, offline U2NetP inference`);
}

module.exports = { openProtocol, terminateOwnedProcessTree };
if (require.main === module) main().catch(error => { console.error(error); process.exitCode = 1; });

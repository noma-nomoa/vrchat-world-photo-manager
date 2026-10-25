'use strict';
// Open only the disposable seeded candidate, through its actual packaged EXE.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { spawn } = require('node:child_process');
const asar = require('@electron/asar');
const { openProtocol, terminateOwnedProcessTree } = require('./smoke-packaged-app');
const { assertInside } = require('./build-review-candidate');
const ROOT = path.resolve(__dirname, '..');
const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
const hash = file => createHash('sha256').update(fs.readFileSync(file)).digest('hex');

async function main() {
  const args = process.argv.slice(2);
  assert(args.length >= 1 && args.length <= 2 && args.every(arg => /^--(?:candidate|exe)=.+/.test(arg)),
    'Usage: node scripts/smoke-review-profile.js --candidate=<fresh review candidate> [--exe=<existing local test build>]');
  const arg = name => args.find(value => value.startsWith(`${name}=`))?.slice(name.length + 1);
  const candidateRoot = path.resolve(arg('--candidate'));
  assertInside(path.join(ROOT, 'out'), candidateRoot);
  assert(/^review-candidate-[^\\/]+$/.test(path.relative(path.join(ROOT, 'out'), candidateRoot)));
  const seed = JSON.parse(fs.readFileSync(path.join(candidateRoot, 'review-profile.json'), 'utf8'));
  const profile = path.join(candidateRoot, 'ReviewProfile');
  assert.equal(seed.profile, profile);
  assert.equal(seed.dbPath, path.join(profile, 'AppData', 'Roaming', 'WorldShot Log', 'data', 'app.sqlite'));
  assert.equal(seed.photos.length, 6);
  for (const file of [seed.dbPath, seed.appData, seed.localAppData, seed.temp]) assertInside(profile, file);
  for (const photo of seed.photos) {
    assertInside(path.join(candidateRoot, 'SamplePhotos'), photo.filePath);
    assert.equal(hash(photo.filePath), photo.fileHash, 'Seed source differs before review');
  }
  const exe = path.resolve(arg('--exe') || path.join(candidateRoot, 'build', 'WorldShot Log-win32-x64', 'WorldShotLog.exe'));
  assertInside(path.join(ROOT, 'out'), exe);
  const archive = path.join(path.dirname(exe), 'resources', 'app.asar');
  assert(asar.extractFile(archive, 'src/index.js').toString().includes('WORLDSHOT_HOME_ROOT'), 'Packaged home isolation is required');
  const output = path.join(candidateRoot, 'review-profile-smoke');
  assert(!fs.existsSync(output), 'Refusing to overwrite review profile smoke artifacts');
  fs.mkdirSync(output);
  const report = { ok: false, exe, candidateRoot, profile, dbPath: seed.dbPath, blockedRequests: [], screenshots: [], renderChecks: [],
    screenshotNote: 'External network requests are blocked. The existing online Material icon font may appear as text; screenshots verify the isolated app and fixtures, not offline icon rendering.' };
  const sockets = new Set();
  const proxy = http.createServer((request, response) => { report.blockedRequests.push(request.url); response.writeHead(502); response.end(); });
  proxy.on('connection', socket => { sockets.add(socket); socket.on('close', () => sockets.delete(socket)); socket.on('error', () => {}); });
  proxy.on('connect', (request, socket) => { report.blockedRequests.push(request.url); socket.end('HTTP/1.1 502 Bad Gateway\r\n\r\n'); });
  await new Promise(resolve => proxy.listen(0, '127.0.0.1', resolve));
  const proxyUrl = `http://127.0.0.1:${proxy.address().port}`;
  const stderrPath = path.join(output, 'stderr.log');
  const stdout = fs.openSync(path.join(output, 'stdout.log'), 'w');
  const stderr = fs.openSync(stderrPath, 'w');
  const child = spawn(exe, ['--squirrel-firstrun', '--remote-debugging-port=0', '--remote-debugging-address=127.0.0.1',
    `--proxy-server=${proxyUrl}`, '--proxy-bypass-list=<-loopback>', '--host-resolver-rules=MAP * ~NOTFOUND, EXCLUDE localhost',
    '--disable-background-timer-throttling', '--disable-backgrounding-occluded-windows',
    '--disable-features=CalculateNativeWinOcclusion'], {
    env: { ...process.env, WORLDSHOT_HOME_ROOT: profile, USERPROFILE: profile, HOME: profile,
      APPDATA: seed.appData, WORLDSHOT_APPDATA_ROOT: seed.appData, LOCALAPPDATA: seed.localAppData,
      WORLDSHOT_LOCALAPPDATA_ROOT: seed.localAppData, TEMP: seed.temp, TMP: seed.temp,
      HTTP_PROXY: proxyUrl, HTTPS_PROXY: proxyUrl, ALL_PROXY: proxyUrl, NO_PROXY: '', NODE_USE_ENV_PROXY: '1' },
    // This UI smoke needs real paint frames; a hidden Windows startup window can suspend rAF.
    stdio: ['ignore', stdout, stderr], windowsHide: false,
  });
  fs.closeSync(stdout); fs.closeSync(stderr);
  let exitInfo; let spawnError; let protocol; let sessionId;
  child.once('error', error => { spawnError = error; });
  const exited = new Promise(resolve => child.once('exit', (code, signal) => { exitInfo = { code, signal }; resolve(); }));
  async function until(label, predicate) {
    const deadline = Date.now() + 30000;
    while (Date.now() < deadline) {
      if (spawnError) throw spawnError;
      assert(!exitInfo, `Candidate exited while waiting for ${label}`);
      const value = await predicate(); if (value) return value;
      await wait(100);
    }
    throw new Error(`Timed out waiting for ${label}`);
  }
  async function evaluate(expression, timeoutMs = 10000) {
    try {
      const value = await protocol.send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true }, sessionId, timeoutMs);
      if (value.exceptionDetails) throw new Error(value.exceptionDetails.exception?.description || value.exceptionDetails.text);
      return value.result?.value;
    } catch (error) {
      const state = await protocol.send('Runtime.evaluate', {
        expression: '({ hidden: document.hidden, visibility: document.visibilityState, focused: document.hasFocus() })',
        returnByValue: true,
      }, sessionId, 2000).catch(probeError => ({ error: probeError.message }));
      report.failedEvaluation = { expression, timeoutMs, pageState: state.result?.value || state };
      throw new Error(`${error.message}\nEvaluation: ${expression}`, { cause: error });
    }
  }
  async function screenshot(name) {
    const frames = await evaluate(`new Promise(resolve => requestAnimationFrame(first => requestAnimationFrame(second =>
      resolve({ first, second, hidden: document.hidden, visibility: document.visibilityState }))))`, 5000);
    assert(frames.second > frames.first && !frames.hidden, `Screenshot paint frames did not advance visibly: ${name}`);
    report.renderChecks.push({ screenshot: name, ...frames });
    await until('settled screenshot animations', () => evaluate(`!document.getAnimations().some(animation =>
      animation.playState === 'running' && animation.effect?.getComputedTiming().iterations !== Infinity)`));
    const value = await protocol.send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: false }, sessionId);
    fs.writeFileSync(path.join(output, name), Buffer.from(value.data, 'base64'));
    report.screenshots.push(`review-profile-smoke/${name}`);
  }
  try {
    const url = await until('debugging endpoint', () => fs.readFileSync(stderrPath, 'utf8')
      .match(/DevTools listening on (ws:\/\/127\.0\.0\.1:\d+\/\S+)/)?.[1]);
    protocol = await openProtocol(url);
    const target = await until('renderer', async () => (await protocol.send('Target.getTargets')).targetInfos
      .find(info => info.type === 'page' && decodeURIComponent(info.url).replaceAll('\\', '/').endsWith('/resources/app.asar/src/index.html')));
    ({ sessionId } = await protocol.send('Target.attachToTarget', { targetId: target.targetId, flatten: true }));
    await protocol.send('Network.enable', {}, sessionId);
    await protocol.send('Network.setBlockedURLs', { urls: ['http://*', 'https://*', 'ws://*', 'wss://*'] }, sessionId);
    await until('initialized seeded gallery', () => evaluate("document.readyState === 'complete' && typeof currentPhotos !== 'undefined' && !document.querySelector('.app')?.classList.contains('is-app-initializing') && currentPhotos.length === 4"));
    const overview = await evaluate(`(async () => ({ summary: await electronAPI.getApplicationDataSummary(),
      september: await electronAPI.getPhotosByMonth(2026, 9), august: await electronAPI.getPhotosByMonth(2026, 8) }))()`);
    assert.equal(overview.summary.photoCount, 6);
    assert.equal(overview.september.length, 4); assert.equal(overview.august.length, 2);
    const photos = [...overview.september, ...overview.august];
    assert.deepEqual(photos.map(photo => photo.id).sort((a, b) => a - b), seed.photos.map(photo => photo.id));
    for (const photo of photos) {
      assert.equal(photo.filePath, seed.photos.find(value => value.id === photo.id).filePath);
      assertInside(path.join(profile, 'WorldShot Log', 'thumbnails'), photo.thumbnailPath);
      assert(fs.existsSync(photo.thumbnailPath));
    }
    report.overview = overview.summary;
    report.thumbnailRoot = path.join(profile, 'WorldShot Log', 'thumbnails');
    await evaluate(`document.querySelector('[data-year="2026"] .month-button[data-month="8"]').click()`);
    await until('August cards', () => evaluate('currentPhotos.length === 2'));
    assert.deepEqual(await evaluate('currentPhotos.map(photo => photo.id)'), [6, 5]);
    await evaluate(`document.querySelector('[data-year="2026"] .year-button').click()`);
    await until('six year cards and loaded thumbnails', () => evaluate(`currentPhotos.length === 6 &&
      document.querySelectorAll('#month-gallery-list .photo-card').length === 6 &&
      [...document.querySelectorAll('#month-gallery-list .photo-card img')].every(image => image.complete && image.naturalWidth > 0)`));
    report.cardOrder = await evaluate(`[...document.querySelectorAll('#month-gallery-list .photo-card')].map(card => Number(card.dataset.photoId))`);
    assert.deepEqual(report.cardOrder, [2, 1, 3, 4, 6, 5]);
    await screenshot('gallery.png');
    await evaluate(`document.querySelector('#month-gallery-list .photo-card[data-photo-id="2"]').click()`);
    await until('portrait modal', () => evaluate("typeof currentModalPhoto !== 'undefined' && currentModalPhoto?.id === 2"));
    await evaluate('openPhotoEditorModal()');
    const source = await until('retouch image', () => evaluate(`photoEditorState?.sourceImage && ({
      filePath: photoEditorState.sourcePhoto.filePath, width: photoEditorState.sourceImage.naturalWidth,
      height: photoEditorState.sourceImage.naturalHeight, isSaving: photoEditorState.isSaving })`));
    assert.equal(source.filePath, seed.photos[1].filePath);
    assert.equal(source.width, 1000); assert.equal(source.height, 1500); assert(!source.isSaving);
    report.retouchPreview = await until('committed retouch preview pixels', () => evaluate(`(() => {
      const canvas = document.querySelector('#photo-editor-canvas');
      if (!photoEditorPreviewCommittedValues || !canvas?.width || !canvas.height) return false;
      const alpha = canvas.getContext('2d').getImageData(Math.floor(canvas.width / 2), Math.floor(canvas.height / 2), 1, 1).data[3];
      return alpha > 0 && ({ width: canvas.width, height: canvas.height, centerAlpha: alpha });
    })()`));
    await screenshot('retouch.png');
    report.retouchSource = source;
    await evaluate('closePhotoEditorModal(); closeImageModal()');
    for (const photo of seed.photos) assert.equal(hash(photo.filePath), photo.fileHash, 'Review changed original fixture');
    report.sourcePhotosUnchanged = true;
    assert(fs.existsSync(seed.dbPath));
    assert(fs.existsSync(path.join(seed.appData, 'WorldShot Log', 'photo-editor-overlays')));
    report.ok = true;
  } catch (error) { report.error = error.stack || String(error); }
  finally {
    const cleanupErrors = []; let forcedTermination = false;
    if (!exitInfo && protocol && sessionId) {
      await protocol.send('Runtime.evaluate', { expression: 'window.close()' }, sessionId, 3000).catch(() => {});
      await Promise.race([exited, wait(3000)]);
    }
    try { protocol?.close(); } catch (error) { cleanupErrors.push(error.message); }
    if (!exitInfo && !spawnError) {
      forcedTermination = true;
      await terminateOwnedProcessTree(child).catch(error => cleanupErrors.push(error.message));
      await Promise.race([exited, wait(5000)]);
    }
    try { for (const socket of sockets) socket.destroy(); proxy.closeAllConnections();
      await new Promise((resolve, reject) => proxy.close(error => error ? reject(error) : resolve()));
    } catch (error) { cleanupErrors.push(error.message); }
    Object.assign(report, { exitInfo, spawnError: spawnError?.message, forcedTermination, cleanupErrors });
    if (spawnError || forcedTermination || exitInfo?.code !== 0 || exitInfo?.signal || cleanupErrors.length) {
      report.ok = false; report.error ||= 'Review app did not exit normally';
    }
    fs.writeFileSync(path.join(output, 'result.json'), JSON.stringify(report, null, 2));
  }
  console.log(`Review profile smoke: ${path.join(output, 'result.json')}`);
  assert(report.ok, report.error);
}
main().catch(error => { console.error(error.stack || error); process.exitCode = 1; });

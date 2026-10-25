'use strict';

// Only the disposable Windows Actions job may use the real installed profile.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { spawn } = require('node:child_process');
const { createHash } = require('node:crypto');
const { openProtocol, terminateOwnedProcessTree } = require('../smoke-packaged-app');
const wait = ms => new Promise(resolve => setTimeout(resolve, ms));

async function main() {
  assert(process.env.GITHUB_ACTIONS === 'true' && process.env.RUNNER_ENVIRONMENT === 'github-hosted' && process.platform === 'win32',
    'This helper may only run in a disposable GitHub-hosted Windows job');
  const [exe, version, mode, output] = process.argv.slice(2);
  assert(exe && /^2\.5\.[34]$/.test(version) && ['seed', 'verify'].includes(mode) && output);
  const name = `${version}-${mode}-${Date.now()}`;
  const stderr = path.join(output, `${name}.stderr.log`);
  const stdoutFd = fs.openSync(path.join(output, `${name}.stdout.log`), 'w');
  const stderrFd = fs.openSync(stderr, 'w');
  const child = spawn(exe, ['--squirrel-firstrun', '--remote-debugging-port=0', '--remote-debugging-address=127.0.0.1',
    '--host-resolver-rules=MAP * ~NOTFOUND, EXCLUDE localhost', '--disable-background-timer-throttling'],
  { stdio: ['ignore', stdoutFd, stderrFd], windowsHide: true });
  fs.closeSync(stdoutFd); fs.closeSync(stderrFd);
  let exited = null;
  let spawnError = null;
  child.once('error', error => { spawnError = error; });
  const exitPromise = new Promise(resolve => child.once('exit', (code, signal) => { exited = { code, signal }; resolve(); }));
  const report = { ok: false, exe, version, mode };
  let protocol;
  let sessionId;
  async function until(label, predicate) {
    const deadline = Date.now() + 45000;
    while (Date.now() < deadline) {
      if (spawnError) throw spawnError;
      assert(!exited, `Application exited before ${label}: ${JSON.stringify(exited)}`);
      const value = await predicate();
      if (value) return value;
      await wait(100);
    }
    throw new Error(`Timed out: ${label}`);
  }
  async function evaluate(expression) {
    const response = await protocol.send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true }, sessionId);
    assert(!response.exceptionDetails, JSON.stringify(response.exceptionDetails));
    return response.result.value;
  }
  try {
    const endpoint = await until('debugging endpoint', () => fs.readFileSync(stderr, 'utf8').match(/DevTools listening on (ws:\/\/127\.0\.0\.1:\d+\/\S+)/)?.[1]);
    protocol = await openProtocol(endpoint);
    const target = await until('installed application page', async () => (await protocol.send('Target.getTargets')).targetInfos
      .find(info => info.type === 'page' && decodeURIComponent(info.url).replaceAll('\\', '/').endsWith('/resources/app.asar/src/index.html')));
    ({ sessionId } = await protocol.send('Target.attachToTarget', { targetId: target.targetId, flatten: true }));
    await protocol.send('Network.enable', {}, sessionId);
    await protocol.send('Network.setBlockedURLs', { urls: ['http://*', 'https://*', 'ws://*', 'wss://*'] }, sessionId);
    await until('application initialization', () => evaluate("document.readyState === 'complete' && !!window.electronAPI && !document.querySelector('.app')?.classList.contains('is-app-initializing')"));
    report.title = await evaluate('document.title');
    assert(report.title.includes(`v${version}`), `Wrong installed application title: ${report.title}`);
    const fixture = path.join(output, 'synthetic-install-photo.png');
    if (mode === 'seed') {
      assert(!fs.existsSync(path.join(output, 'baseline-data.json')), 'Refusing to overwrite seeded data');
      const png = await evaluate("(() => { const c = document.createElement('canvas'); c.width = 80; c.height = 60; const x = c.getContext('2d'); x.fillStyle = '#379ab5'; x.fillRect(0, 0, 80, 60); x.fillStyle = '#f7b530'; x.fillRect(10, 10, 30, 30); return c.toDataURL('image/png').split(',')[1]; })()");
      fs.writeFileSync(fixture, Buffer.from(png, 'base64'));
      await evaluate("(() => { const input = document.createElement('input'); input.type = 'file'; input.id = 'ci-import-fixture'; document.body.appendChild(input); })()");
      const { root } = await protocol.send('DOM.getDocument', {}, sessionId);
      const { nodeId } = await protocol.send('DOM.querySelector', { nodeId: root.nodeId, selector: '#ci-import-fixture' }, sessionId);
      await protocol.send('DOM.setFileInputFiles', { nodeId, files: [fixture] }, sessionId);
      // The renderer passes an Array<File>; FileList itself is not preserved
      // when Electron copies arguments across the isolated context bridge.
      const imported = await evaluate("electronAPI.importDroppedFiles(Array.from(document.getElementById('ci-import-fixture').files))");
      assert.equal(imported.importedCount, 1, JSON.stringify(imported));
      const edits = await evaluate(`(async () => {
        const month = await electronAPI.getLatestMonth();
        const photos = await electronAPI.getPhotosByMonth(month.year, month.month);
        if (photos.length !== 1) throw new Error('Expected one synthetic photo');
        const id = photos[0].id;
        const results = [await electronAPI.updateFavoriteStatus(id, true),
          await electronAPI.updatePhotoMemo(id, 'Installed upgrade keeps this memo 日本語'),
          await electronAPI.updateWorldName(id, 'Installation fixture world'),
          await electronAPI.replacePhotoLabels(id, [{ name: 'Upgrade fixture', colorHex: '#4A8EAC' }])];
        applyTheme('light');
        return results;
      })()`);
      assert(edits.every(result => result.ok), JSON.stringify(edits));
    }
    report.data = await evaluate(`(async () => {
      const month = await electronAPI.getLatestMonth();
      const photos = await electronAPI.getPhotosByMonth(month.year, month.month);
      const records = [];
      for (const p of photos) records.push({ id: p.id, filePath: p.filePath, isFavorite: p.isFavorite,
        memoText: p.memoText, worldNameManual: p.worldNameManual, labels: await electronAPI.getPhotoLabels(p.id) });
      return { records, theme: localStorage.getItem('vrchat-world-photo-manager-theme'), appliedTheme: document.body.dataset.theme };
    })()`);
    assert.equal(report.data.records.length, 1);
    assert.equal(report.data.records[0].isFavorite, true);
    assert.equal(report.data.theme, 'light');
    assert.equal(report.data.appliedTheme, 'light');
    report.fixtureSha256 = createHash('sha256').update(fs.readFileSync(fixture)).digest('hex');
    if (mode === 'seed') fs.writeFileSync(path.join(output, 'baseline-data.json'), JSON.stringify({ data: report.data, fixtureSha256: report.fixtureSha256 }, null, 2));
    else assert.deepEqual({ data: report.data, fixtureSha256: report.fixtureSha256 }, JSON.parse(fs.readFileSync(path.join(output, 'baseline-data.json'), 'utf8')));
    report.ok = true;
  } catch (error) {
    report.error = error.stack;
  } finally {
    if (!exited && protocol && sessionId) {
      await protocol.send('Runtime.evaluate', { expression: 'window.close()' }, sessionId, 3000).catch(() => {});
      await Promise.race([exitPromise, wait(5000)]);
    }
    protocol?.close();
    if (!exited && !spawnError) {
      report.forcedTermination = true;
      await terminateOwnedProcessTree(child);
      await Promise.race([exitPromise, wait(5000)]);
    }
    report.exit = exited;
    report.ok = report.ok && !report.forcedTermination && !spawnError && exited?.code === 0 && !exited.signal;
    fs.writeFileSync(path.join(output, `${name}.json`), JSON.stringify(report, null, 2));
  }
  assert(report.ok, report.error || JSON.stringify(report));
}
main().catch(error => { console.error(error); process.exitCode = 1; });

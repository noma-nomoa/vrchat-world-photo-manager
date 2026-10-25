// Real main/preload update flow with isolated storage and native/network ports stubbed.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const Module = require('node:module');
const { app, BrowserWindow, autoUpdater, net, session } = require('electron');
const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-update-'));
const OUTPUT = path.resolve(__dirname, '..', 'out', 'smoke-app-update', path.basename(ROOT));
const PROFILE = path.join(ROOT, 'Profile');
const APPDATA = path.join(PROFILE, 'AppData', 'Roaming');
const LOCAL = path.join(PROFILE, 'AppData', 'Local');
const TEMP = path.join(ROOT, 'Temp');
for (const dir of [OUTPUT, PROFILE, APPDATA, LOCAL, TEMP]) fs.mkdirSync(dir, { recursive: true });
process.env.WORLDSHOT_HOME_ROOT = process.env.USERPROFILE = PROFILE;
process.env.WORLDSHOT_APPDATA_ROOT = process.env.APPDATA = APPDATA;
process.env.WORLDSHOT_LOCALAPPDATA_ROOT = process.env.LOCALAPPDATA = LOCAL;
process.env.TEMP = process.env.TMP = TEMP;
app.setPath('home', PROFILE);
app.setPath('appData', APPDATA);
app.setPath('userData', path.join(APPDATA, 'WorldShot Log'));
app.setPath('temp', TEMP);
Object.defineProperty(app, 'isPackaged', { value: true });
app.getVersion = () => '2.5.3';
let nextFetch = async () => ({ ok: true, status: 200, json: async () => ({ tag_name: 'v2.6.0', name: 'WorldShot Log 2.6.0', body: '## Changes\n- Fixture change' }) });
let feedError = null;
let checkError = null;
const calls = { fetch: 0, feeds: [], checks: 0, installs: 0 };
net.fetch = async () => { calls.fetch += 1; return nextFetch(); };
autoUpdater.setFeedURL = payload => { calls.feeds.push(payload); if (feedError) throw new Error(feedError); };
autoUpdater.checkForUpdates = () => { calls.checks += 1; if (checkError) throw new Error(checkError); };
autoUpdater.quitAndInstall = () => { calls.installs += 1; };
const report = { ok: false, root: ROOT, checks: [] };
let win;
const inPage = (fn, ...args) => win.webContents.executeJavaScript(`(${fn.toString()})(...${JSON.stringify(args)})`);
async function until(label, predicate) {
  const deadline = Date.now() + 10000;
  while (Date.now() < deadline) { if (await predicate()) return; await new Promise(resolve => setTimeout(resolve, 25)); }
  throw new Error(`Timed out: ${label}`);
}
async function check(name, fn) {
  await fn(); report.checks.push({ name, ok: true }); console.log(`PASS ${name}`);
}
async function main() {
  await app.whenReady();
  session.defaultSession.webRequest.onBeforeRequest({ urls: ['http://*/*', 'https://*/*'] }, (_details, callback) => callback({ cancel: true }));
  app.on('browser-window-created', (_event, window) => { window.hide(); window.webContents.setBackgroundThrottling(false); });
  const filename = path.resolve(__dirname, '..', 'src', 'index.js');
  const mainModule = new Module(filename, module);
  mainModule.filename = filename;
  mainModule.paths = Module._nodeModulePaths(path.dirname(filename));
  // The appended read-only test seam exposes an existing startup action. All
  // product functions and registered IPC handlers execute their real source.
  mainModule._compile(fs.readFileSync(filename, 'utf8') + '\nmodule.exports = { checkForAppUpdatesOnLaunch };\n', filename);
  await until('window', () => { win = BrowserWindow.getAllWindows()[0]; return win; });
  await until('preload', () => inPage(() => typeof window.electronAPI?.startAppUpdateDownload === 'function'));
  const api = mainModule.exports;
  await check('download request without a discovered release returns a failure', async () => {
    const result = await inPage(() => electronAPI.startAppUpdateDownload());
    assert.equal(result.ok, false); assert.equal(calls.checks, 0);
  });
  await check('checking a newer release only notifies and does not download/install', async () => {
    await api.checkForAppUpdatesOnLaunch();
    assert.equal(calls.fetch, 1); assert.equal(calls.checks, 0); assert.equal(calls.installs, 0);
  });
  await check('synchronous native feed failure is reported and a later retry runs', async () => {
    feedError = 'controlled feed failure';
    const failed = await inPage(() => electronAPI.startAppUpdateDownload());
    assert.equal(failed.ok, false); assert.match(failed.message, /controlled feed failure/);
    const previousFeedCalls = calls.feeds.length;
    feedError = null;
    const retry = await inPage(() => electronAPI.startAppUpdateDownload());
    assert.equal(retry.ok, true);
    assert.equal(calls.feeds.length, previousFeedCalls + 1, 'Retry must invoke native updater again after synchronous failure');
    assert.equal(calls.checks, 1);
    assert.equal(calls.installs, 0);
  });
  await check('duplicate requests during a running download are ignored', async () => {
    const previous = calls.checks;
    await inPage(() => electronAPI.startAppUpdateDownload());
    assert.equal(calls.checks, previous);
  });
  await check('native asynchronous error permits a later retry', async () => {
    autoUpdater.emit('error', new Error('controlled async failure'));
    const previous = calls.checks;
    const result = await inPage(() => electronAPI.startAppUpdateDownload());
    assert.equal(result.ok, true); assert.equal(calls.checks, previous + 1);
  });
  await check('synchronous check failure also permits a later retry', async () => {
    autoUpdater.emit('update-not-available');
    checkError = 'controlled check failure';
    const failed = await inPage(() => electronAPI.startAppUpdateDownload());
    assert.equal(failed.ok, false);
    const previous = calls.checks;
    checkError = null;
    const retry = await inPage(() => electronAPI.startAppUpdateDownload());
    assert.equal(retry.ok, true); assert.equal(calls.checks, previous + 1);
  });
  await check('download completion does not install without a separate action', async () => {
    autoUpdater.emit('update-downloaded');
    assert.equal(calls.installs, 0);
  });
  await check('network failure on startup permits another check', async () => {
    nextFetch = async () => { throw new Error('controlled network failure'); };
    const previous = calls.fetch;
    await api.checkForAppUpdatesOnLaunch();
    nextFetch = async () => ({ ok: true, status: 204 });
    await api.checkForAppUpdatesOnLaunch();
    assert.equal(calls.fetch, previous + 2);
    assert.equal(calls.installs, 0);
  });
  report.ok = true;
}
const watchdog = setTimeout(() => { report.error = 'App update smoke timed out'; fs.writeFileSync(path.join(OUTPUT, 'result.json'), JSON.stringify(report, null, 2)); app.exit(1); }, 30000);
main().catch(error => { report.error = error.stack || String(error); console.error(report.error); })
  .finally(() => {
    clearTimeout(watchdog); report.calls = calls;
    fs.writeFileSync(path.join(OUTPUT, 'result.json'), JSON.stringify(report, null, 2));
    console.log(`App update report: ${OUTPUT}`);
    app.exit(report.ok ? 0 : 1);
  });

'use strict';

// Electron is used only for native image codecs; every filesystem mutation is
// restricted to the freshly created test root by the shared scenario harness.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const { app, nativeImage } = require('electron');
const { createThumbnailStore } = require('../src/thumbnail-store');
const { runThumbnailStoreCases, runThumbnailPathSafetyCases } = require('./fixtures/thumbnail-store-cases');
const baseline = require('./fixtures/thumbnail-store-baseline.json');

const root = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-thumbnail-store-'));
const output = path.resolve(__dirname, '../out/smoke-thumbnail-store', path.basename(root));
const profile = path.join(root, 'profile');
const appData = path.join(profile, 'AppData', 'Roaming');
const localAppData = path.join(profile, 'AppData', 'Local');
const temp = path.join(root, 'temp');
for (const target of [output, profile, appData, localAppData, temp]) fs.mkdirSync(target, { recursive: true });
process.env.APPDATA = process.env.WORLDSHOT_APPDATA_ROOT = appData;
process.env.LOCALAPPDATA = process.env.WORLDSHOT_LOCALAPPDATA_ROOT = localAppData;
process.env.USERPROFILE = profile;
process.env.TEMP = process.env.TMP = temp;
app.setPath('home', profile);
app.setPath('appData', appData);
app.setPath('userData', path.join(appData, 'WorldShot Log'));
app.setPath('sessionData', app.getPath('userData'));
app.setPath('temp', temp);
for (const name of ['pictures', 'documents', 'downloads', 'desktop']) {
  const target = path.join(profile, name);
  fs.mkdirSync(target, { recursive: true });
  app.setPath(name, target);
}

const watchdog = setTimeout(() => {
  fs.writeFileSync(path.join(output, 'result.json'), JSON.stringify({ ok: false, root,
    error: 'Thumbnail store smoke exceeded 30 seconds' }, null, 2));
  app.exit(1);
}, 30000);
app.whenReady().then(async () => {
  const result = await runThumbnailStoreCases({ createStore: createThumbnailStore,
    fs: fsp, nativeImage, root });
  assert.deepEqual(result, baseline, 'Filesystem behavior and generated thumbnail bytes must match the pre-extraction baseline');
  const pathSafety = await runThumbnailPathSafetyCases({ createStore: createThumbnailStore,
    fs: fsp, nativeImage, root });
  const caseCount = result.cases.length + pathSafety.caseCount;
  fs.writeFileSync(path.join(output, 'result.json'), JSON.stringify({ ok: true, root,
    caseCount, compatibilityCaseCount: result.cases.length, pathSafety, ...result }, null, 2));
  console.log(`Thumbnail store smoke: ${caseCount} cases passed. ${output}`);
  app.quit();
}).catch(error => {
  fs.writeFileSync(path.join(output, 'result.json'), JSON.stringify({ ok: false, root,
    error: error.stack || String(error) }, null, 2));
  console.error(error.stack || error);
  console.error(`Isolated artifacts: ${root}`);
  app.exit(1);
}).finally(() => clearTimeout(watchdog));

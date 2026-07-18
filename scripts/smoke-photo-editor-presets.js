const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { pathToFileURL } = require('node:url');

const { app, BrowserWindow } = require('electron');
const { initDatabase } = require('../src/db');

const REPO_ROOT = path.resolve(__dirname, '..');
const VIEWPORT = { width: 1600, height: 1000 };
const RUN_ID = createRunId();
const SMOKE_ROOT = path.join(os.tmpdir(), 'worldshot-presets-' + RUN_ID);
const SMOKE_APPDATA = path.join(SMOKE_ROOT, 'AppData', 'Roaming');
const SMOKE_LOCALAPPDATA = path.join(SMOKE_ROOT, 'AppData', 'Local');
const OUTPUT_DIR = path.join(REPO_ROOT, 'out', 'smoke-presets', RUN_ID);
const NEW_PRESET_KEYS = [
  'airyWhite',
  'warmAiryWhite',
  'coolAiryWhite',
  'dreamyAiryWhite',
];
const EXPECTED_SEQUENCE = [
  'sweetPink',
  'pastelPink',
  'vividPink',
  ...NEW_PRESET_KEYS,
  'shadowLift',
];
const EXPECTED_LABELS = {
  ja: [
    '\u30a8\u30a2\u30ea\u30fc\u30db\u30ef\u30a4\u30c81',
    '\u30a8\u30a2\u30ea\u30fc\u30db\u30ef\u30a4\u30c82',
    '\u30a8\u30a2\u30ea\u30fc\u30db\u30ef\u30a4\u30c83',
    '\u30a8\u30a2\u30ea\u30fc\u30db\u30ef\u30a4\u30c84',
  ],
  en: ['Airy White 1', 'Airy White 2', 'Airy White 3', 'Airy White 4'],
  ko: [
    '\uc5d0\uc5b4\ub9ac \ud654\uc774\ud2b8 1',
    '\uc5d0\uc5b4\ub9ac \ud654\uc774\ud2b8 2',
    '\uc5d0\uc5b4\ub9ac \ud654\uc774\ud2b8 3',
    '\uc5d0\uc5b4\ub9ac \ud654\uc774\ud2b8 4',
  ],
};

ensureDir(SMOKE_APPDATA);
ensureDir(SMOKE_LOCALAPPDATA);
ensureDir(OUTPUT_DIR);

process.env.APPDATA = SMOKE_APPDATA;
process.env.LOCALAPPDATA = SMOKE_LOCALAPPDATA;
app.setPath('appData', SMOKE_APPDATA);
app.setPath('userData', path.join(SMOKE_APPDATA, 'WorldShot Log'));

function pad(value) {
  return String(value).padStart(2, '0');
}

function createRunId() {
  const now = new Date();
  return [
    now.getFullYear(),
    pad(now.getMonth() + 1),
    pad(now.getDate()),
    '-',
    pad(now.getHours()),
    pad(now.getMinutes()),
    pad(now.getSeconds()),
  ].join('');
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function seedIsolatedData() {
  const userData = path.join(SMOKE_APPDATA, 'WorldShot Log');
  const dataDir = path.join(userData, 'data');
  const sampleDir = path.join(SMOKE_ROOT, 'SamplePhotos');
  ensureDir(dataDir);
  ensureDir(sampleDir);

  const samplePhotoPath = path.join(
    sampleDir,
    'VRChat_2026-07-18_12-00-00.000_1920x1080.png'
  );
  fs.copyFileSync(path.join(REPO_ROOT, 'img', 'notuse.png'), samplePhotoPath);

  const db = initDatabase(path.join(dataDir, 'app.sqlite'));
  const nowIso = new Date().toISOString();
  db.insertOrUpdatePhoto({
    filePath: samplePhotoPath,
    fileName: path.basename(samplePhotoPath),
    fileHash: 'photo-editor-preset-smoke-photo',
    takenAt: '2026/07/18 12:00:00',
    takenAtTimestamp: new Date(2026, 6, 18, 12, 0, 0).getTime(),
    groupDate: '2026-07-18',
    year: 2026,
    month: 7,
    day: 18,
    worldId: 'wrld_photo_editor_preset_smoke',
    worldName: 'preset smoke',
    worldNameManual: null,
    worldUrl: 'https://vrchat.com/home/world/wrld_photo_editor_preset_smoke',
    thumbnailPath: samplePhotoPath,
    imageWidth: 1920,
    imageHeight: 1080,
    resolutionTier: 'FHD',
    orientationTier: 'landscape',
    printNoteText: '',
    memoText: '',
    createdAt: nowIso,
    updatedAt: nowIso,
  });

  db.upsertWorldMetadata({
    worldId: 'wrld_photo_editor_preset_smoke',
    sourceUrl: 'https://vrchat.com/home/world/wrld_photo_editor_preset_smoke',
    worldNameOfficial: 'preset smoke',
    worldDescription: 'Smoke test sample for built-in photo editor presets.',
    worldTags: ['debug_sample'],
    authorId: 'usr_photo_editor_preset_smoke',
    authorName: 'WorldShot Log',
    releaseStatus: 'public',
    imageUrl: pathToFileURL(samplePhotoPath).href,
    thumbnailImageUrl: pathToFileURL(samplePhotoPath).href,
    fetchStatus: 'success',
  });

  db.close();
  return { samplePhotoPath };
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForWindow() {
  const startedAt = Date.now();

  while (Date.now() - startedAt < 45000) {
    const win = BrowserWindow.getAllWindows().find(
      (candidate) => !candidate.isDestroyed()
    );

    if (win) {
      return win;
    }

    await wait(100);
  }

  throw new Error('Timed out waiting for the application window.');
}

async function runInPage(win, script) {
  return win.webContents.executeJavaScript(script);
}

async function waitForRendererReady(win) {
  const startedAt = Date.now();
  let lastState = null;

  while (Date.now() - startedAt < 20000) {
    const state = await runInPage(
      win,
      [
        '(() => {',
        '  const appEl = document.querySelector(\'.app\');',
        '  return {',
        '    readyState: document.readyState,',
        '    hasI18n: Boolean(window.WorldShotI18n),',
        '    initializing: Boolean(appEl?.classList.contains(\'is-app-initializing\')),',
        '    photoCards: document.querySelectorAll(\'.photo-card\').length,',
        '  };',
        '})()',
      ].join('\n')
    );
    lastState = state;

    if (
      state.hasI18n &&
      (!state.initializing || state.photoCards > 0) &&
      state.readyState !== 'loading'
    ) {
      return;
    }

    await wait(150);
  }

  throw new Error(
    'Timed out waiting for renderer. Last state: ' + JSON.stringify(lastState)
  );
}

async function waitForSelector(win, selector, options = {}) {
  const timeoutMs = options.timeoutMs || 10000;
  const visible = options.visible !== false;
  const startedAt = Date.now();

  while (Date.now() - startedAt < timeoutMs) {
    const found = await runInPage(
      win,
      [
        '(() => {',
        '  const el = document.querySelector(' + JSON.stringify(selector) + ');',
        '  if (!el) return false;',
        '  if (!' + JSON.stringify(visible) + ') return true;',
        '  const rect = el.getBoundingClientRect();',
        '  const style = getComputedStyle(el);',
        '  return rect.width > 0 && rect.height > 0 && style.visibility !== \'hidden\' && style.display !== \'none\';',
        '})()',
      ].join('\n')
    );

    if (found) {
      return;
    }

    await wait(100);
  }

  throw new Error('Timed out waiting for selector: ' + selector);
}

async function waitForPagePredicate(win, script, label, timeoutMs = 15000) {
  const startedAt = Date.now();
  let lastValue = null;

  while (Date.now() - startedAt < timeoutMs) {
    lastValue = await runInPage(win, script);
    const isReady =
      lastValue && typeof lastValue === 'object' && 'ready' in lastValue
        ? lastValue.ready
        : lastValue;

    if (isReady) {
      return lastValue;
    }

    await wait(100);
  }

  throw new Error(
    'Timed out waiting for ' + label + '. Last value: ' + JSON.stringify(lastValue)
  );
}

function attachWindowDiagnostics(win) {
  win.webContents.on('console-message', (event) => {
    const level = event.level ?? 'unknown';
    const message = event.message ?? '';
    console.log('[renderer console ' + level + '] ' + message);
  });
  win.webContents.on('render-process-gone', (_event, details) => {
    console.error('[renderer process gone] ' + JSON.stringify(details));
  });
}

async function openPhotoEditor(win) {
  await waitForSelector(win, '.photo-card', {
    timeoutMs: 30000,
    visible: false,
  });
  await runInPage(win, 'document.querySelector(\'.photo-card\')?.click();');
  await waitForSelector(win, '#image-modal:not(.hidden)');
  await runInPage(
    win,
    'document.getElementById(\'modal-edit-photo-btn\')?.click();'
  );
  await waitForSelector(win, '#photo-editor-modal:not(.hidden)', {
    timeoutMs: 15000,
  });
  await wait(900);
  await runInPage(
    win,
    [
      '(() => {',
      '  setPhotoEditorAccordionOpen(\'presets\', true);',
      '  return true;',
      '})()',
    ].join('\n')
  );
  await wait(250);
}

async function capturePageWithRetry(win, rect = undefined) {
  let lastError = null;

  for (let attempt = 1; attempt <= 3; attempt += 1) {
    try {
      return await win.capturePage(rect);
    } catch (error) {
      lastError = error;
      if (!String(error?.message || error).includes('UnknownVizError')) {
        throw error;
      }
      await wait(500 * attempt);
    }
  }

  throw lastError;
}

async function captureWindow(win, name) {
  await wait(200);
  const screenshotPath = path.join(OUTPUT_DIR, name + '.png');
  const image = await capturePageWithRetry(win);
  fs.writeFileSync(screenshotPath, image.toPNG());
  return screenshotPath;
}

async function captureCanvas(win, name) {
  const rect = await runInPage(
    win,
    [
      '(() => {',
      '  const rect = document.getElementById(\'photo-editor-canvas\').getBoundingClientRect();',
      '  return {',
      '    x: Math.max(0, Math.floor(rect.left)),',
      '    y: Math.max(0, Math.floor(rect.top)),',
      '    width: Math.max(1, Math.floor(rect.width)),',
      '    height: Math.max(1, Math.floor(rect.height)),',
      '  };',
      '})()',
    ].join('\n')
  );
  const screenshotPath = path.join(OUTPUT_DIR, name + '.png');
  const image = await capturePageWithRetry(win, rect);
  fs.writeFileSync(screenshotPath, image.toPNG());
  return screenshotPath;
}

async function readPresetUi(win) {
  return runInPage(
    win,
    [
      '(() => Array.from(',
      '  document.querySelectorAll(\'#photo-editor-presets .photo-editor-preset-button\')',
      ').map((button) => ({',
      '  key: button.dataset.photoEditorPreset,',
      '  label: button.textContent.trim(),',
      '  fits: button.scrollWidth <= button.clientWidth + 1,',
      '})))()',
    ].join('\n')
  );
}

function assertPresetOrder(items) {
  const keys = items.map((item) => item.key);
  const start = keys.indexOf(EXPECTED_SEQUENCE[0]);
  const actual = keys.slice(start, start + EXPECTED_SEQUENCE.length);

  if (start < 0 || JSON.stringify(actual) !== JSON.stringify(EXPECTED_SEQUENCE)) {
    throw new Error(
      'Preset order mismatch: ' + JSON.stringify({ expected: EXPECTED_SEQUENCE, actual })
    );
  }
}

async function verifyLanguageUi(win, language) {
  await runInPage(
    win,
    'window.WorldShotI18n.setLanguage(' + JSON.stringify(language) + ');'
  );
  await wait(250);
  const items = await readPresetUi(win);
  assertPresetOrder(items);
  const newItems = NEW_PRESET_KEYS.map((key) =>
    items.find((item) => item.key === key)
  );
  const labels = newItems.map((item) => item?.label || '');

  if (JSON.stringify(labels) !== JSON.stringify(EXPECTED_LABELS[language])) {
    throw new Error(
      'Preset labels mismatch for ' +
        language +
        ': ' +
        JSON.stringify({ expected: EXPECTED_LABELS[language], actual: labels })
    );
  }

  const overflow = newItems.filter((item) => !item?.fits);
  if (overflow.length > 0) {
    throw new Error(
      'Preset labels overflow for ' + language + ': ' + JSON.stringify(overflow)
    );
  }

  await runInPage(
    win,
    [
      '(() => {',
      '  const button = document.querySelector(\'[data-photo-editor-preset="dreamyAiryWhite"]\');',
      '  button?.scrollIntoView({ block: \'center\' });',
      '  return Boolean(button);',
      '})()',
    ].join('\n')
  );
  await wait(150);
  return {
    language,
    labels,
    screenshot: await captureWindow(win, language + '-preset-list'),
  };
}

async function readCanvasMetrics(win) {
  return runInPage(
    win,
    [
      '(() => {',
      '  const canvas = document.getElementById(\'photo-editor-canvas\');',
      '  const ctx = canvas.getContext(\'2d\', { willReadFrequently: true });',
      '  const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;',
      '  let count = 0;',
      '  let sum = 0;',
      '  let sumSquares = 0;',
      '  let sumRed = 0;',
      '  let sumBlue = 0;',
      '  let brightCount = 0;',
      '  let hash = 2166136261;',
      '  for (let index = 0; index < data.length; index += 16) {',
      '    const red = data[index];',
      '    const green = data[index + 1];',
      '    const blue = data[index + 2];',
      '    const luminance = 0.2126 * red + 0.7152 * green + 0.0722 * blue;',
      '    count += 1;',
      '    sum += luminance;',
      '    sumSquares += luminance * luminance;',
      '    sumRed += red;',
      '    sumBlue += blue;',
      '    if (luminance >= 220) brightCount += 1;',
      '    hash = Math.imul(hash ^ red, 16777619);',
      '    hash = Math.imul(hash ^ green, 16777619);',
      '    hash = Math.imul(hash ^ blue, 16777619);',
      '  }',
      '  const mean = sum / Math.max(1, count);',
      '  const variance = Math.max(0, sumSquares / Math.max(1, count) - mean * mean);',
      '  return {',
      '    width: canvas.width,',
      '    height: canvas.height,',
      '    mean: Number(mean.toFixed(3)),',
      '    stdDev: Number(Math.sqrt(variance).toFixed(3)),',
      '    warmth: Number(((sumRed - sumBlue) / Math.max(1, count)).toFixed(3)),',
      '    brightShare: Number((brightCount / Math.max(1, count)).toFixed(5)),',
      '    hash: (hash >>> 0).toString(16).padStart(8, \'0\'),',
      '  };',
      '})()',
    ].join('\n')
  );
}

function assertPresetLogic(key, values) {
  if (!(values.exposure < 0 && values.contrast < 0 && values.whites > 0)) {
    throw new Error(
      key +
        ' must lower exposure and contrast while raising whites: ' +
        JSON.stringify(values)
    );
  }

  if (!(values.clarity < 0 && values.texture < 0)) {
    throw new Error(
      key + ' must soften clarity and texture: ' + JSON.stringify(values)
    );
  }
}

async function applyPresetAndCapture(win, key) {
  const values = await runInPage(
    win,
    [
      '(() => {',
      '  const key = ' + JSON.stringify(key) + ';',
      '  updatePhotoEditorAutoEnhanceStrength(100);',
      '  const button = document.querySelector(',
      '    \'[data-photo-editor-preset="\' + key + \'"]\'',
      '  );',
      '  if (!button) throw new Error(\'Preset button not found: \' + key);',
      '  button.click();',
      '  return { ...photoEditorState.values };',
      '})()',
    ].join('\n')
  );
  assertPresetLogic(key, values);

  await waitForPagePredicate(
    win,
    [
      '(() => {',
      '  const expected = ' + JSON.stringify(values) + ';',
      '  const committed = photoEditorPreviewCommittedValues;',
      '  const ready = Boolean(committed && Object.entries(expected).every(',
      '    ([name, value]) => committed[name] === value',
      '  ));',
      '  return {',
      '    ready,',
      '    expected,',
      '    committed: committed ? { ...committed } : null,',
      '    current: photoEditorState?.values ? { ...photoEditorState.values } : null,',
      '    autoEnhance: photoEditorState?.autoEnhance ? { ...photoEditorState.autoEnhance } : null,',
      '  };',
      '})()',
    ].join('\n'),
    key + ' preview render'
  );
  await wait(120);

  return {
    key,
    values,
    metrics: await readCanvasMetrics(win),
    screenshot: await captureCanvas(win, key),
  };
}

function assertVariantRelationships(originalMetrics, results) {
  const byKey = Object.fromEntries(results.map((result) => [result.key, result]));
  const hashes = new Set(results.map((result) => result.metrics.hash));

  if (hashes.size !== results.length) {
    throw new Error(
      'Airy White presets did not render four distinct outputs: ' +
        JSON.stringify(results.map((result) => result.metrics))
    );
  }

  for (const result of results) {
    if (result.metrics.hash === originalMetrics.hash) {
      throw new Error(result.key + ' rendered the same pixels as the original.');
    }
    if (result.metrics.stdDev >= originalMetrics.stdDev) {
      throw new Error(
        result.key +
          ' did not soften luminance contrast: ' +
          JSON.stringify({ original: originalMetrics, preset: result.metrics })
      );
    }
  }

  if (
    byKey.warmAiryWhite.metrics.warmth <=
    byKey.coolAiryWhite.metrics.warmth
  ) {
    throw new Error(
      'Warm and cool variants are not ordered by color temperature: ' +
        JSON.stringify({
          warm: byKey.warmAiryWhite.metrics,
          cool: byKey.coolAiryWhite.metrics,
        })
    );
  }

  if (
    byKey.dreamyAiryWhite.metrics.stdDev >= byKey.airyWhite.metrics.stdDev
  ) {
    throw new Error(
      'Dreamy variant should be softer than the neutral variant: ' +
        JSON.stringify({
          neutral: byKey.airyWhite.metrics,
          dreamy: byKey.dreamyAiryWhite.metrics,
        })
    );
  }
}

async function main() {
  console.log('[preset-smoke] output: ' + OUTPUT_DIR);
  const seeded = seedIsolatedData();
  require('../src/index.js');

  const win = await waitForWindow();
  attachWindowDiagnostics(win);
  win.setSize(VIEWPORT.width, VIEWPORT.height);
  await waitForRendererReady(win);
  await openPhotoEditor(win);

  const languageChecks = [];
  for (const language of ['ja', 'en', 'ko']) {
    languageChecks.push(await verifyLanguageUi(win, language));
  }
  await runInPage(win, 'window.WorldShotI18n.setLanguage(\'ja\');');
  await wait(200);

  const originalMetrics = await readCanvasMetrics(win);
  const originalScreenshot = await captureCanvas(win, 'original');
  const presetResults = [];

  for (const key of NEW_PRESET_KEYS) {
    presetResults.push(await applyPresetAndCapture(win, key));
  }

  assertVariantRelationships(originalMetrics, presetResults);

  const result = {
    ok: true,
    runId: RUN_ID,
    smokeRoot: SMOKE_ROOT,
    samplePhotoPath: seeded.samplePhotoPath,
    original: {
      metrics: originalMetrics,
      screenshot: originalScreenshot,
    },
    presets: presetResults,
    languages: languageChecks,
  };
  const reportPath = path.join(OUTPUT_DIR, 'result.json');
  fs.writeFileSync(reportPath, JSON.stringify(result, null, 2));
  console.log('Photo editor preset smoke passed: ' + OUTPUT_DIR);
  app.quit();
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
  app.exit(1);
});

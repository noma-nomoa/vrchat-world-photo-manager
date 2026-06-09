const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { pathToFileURL } = require('node:url');

const REPO_ROOT = path.resolve(__dirname, '..');
const LANGUAGES = ['ja', 'en', 'ko'];
const SCREENSHOT_VIEWPORT = { width: 1440, height: 920 };
const HAS_JAPANESE_RE = /[\u3040-\u30ff\u3400-\u9fff]/;
const ALLOWED_UNTRANSLATED_TEXTS = ['日本語'];

const RUN_ID = createRunId();
const SMOKE_ROOT = path.join(os.tmpdir(), `worldshot-ui-screenshots-${RUN_ID}`);
const SMOKE_APPDATA = path.join(SMOKE_ROOT, 'AppData', 'Roaming');
const SMOKE_LOCALAPPDATA = path.join(SMOKE_ROOT, 'AppData', 'Local');
ensureDir(SMOKE_APPDATA);
ensureDir(SMOKE_LOCALAPPDATA);
process.env.APPDATA = SMOKE_APPDATA;
process.env.LOCALAPPDATA = SMOKE_LOCALAPPDATA;

const { app, BrowserWindow } = require('electron');
const { initDatabase } = require('../src/db');
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

function seedIsolatedData(smokeRoot) {
  const appData = SMOKE_APPDATA;
  const localAppData = SMOKE_LOCALAPPDATA;
  const userData = path.join(appData, 'WorldShot Log');
  const dataDir = path.join(userData, 'data');
  const sampleDir = path.join(smokeRoot, 'SamplePhotos');
  ensureDir(dataDir);
  ensureDir(localAppData);
  ensureDir(sampleDir);

  const samplePhotoPath = path.join(
    sampleDir,
    'VRChat_2026-06-07_22-20-00.000_3840x2160.png'
  );
  fs.copyFileSync(path.join(REPO_ROOT, 'img', 'logo.png'), samplePhotoPath);

  const db = initDatabase(path.join(dataDir, 'app.sqlite'));
  const nowIso = new Date().toISOString();
  const row = db.insertOrUpdatePhoto({
    filePath: samplePhotoPath,
    fileName: path.basename(samplePhotoPath),
    fileHash: 'visual-audit-sample-photo',
    takenAt: '2026/06/07 22:20:00',
    takenAtTimestamp: new Date(2026, 5, 7, 22, 20, 0).getTime(),
    groupDate: '2026-06-07',
    year: 2026,
    month: 6,
    day: 7,
    worldId: 'wrld_visual_audit',
    worldName: 'nagisa no machi',
    worldNameManual: null,
    worldUrl: 'https://vrchat.com/home/world/wrld_visual_audit',
    thumbnailPath: samplePhotoPath,
    imageWidth: 3840,
    imageHeight: 2160,
    resolutionTier: '4K',
    orientationTier: 'landscape',
    printNoteText: '',
    memoText: 'Visual audit memo',
    createdAt: nowIso,
    updatedAt: nowIso,
  });

  if (row?.id) {
    db.replacePhotoTags(row.id, [
      { name: 'system approved', colorHex: '#6D5EF6' },
      { name: 'photo walk', colorHex: '#14B8A6' },
    ]);
  }

  db.upsertWorldMetadata({
    worldId: 'wrld_visual_audit',
    sourceUrl: 'https://vrchat.com/home/world/wrld_visual_audit',
    worldNameOfficial: 'nagisa no machi',
    worldDescription: 'A quiet seaside town prepared for visual regression checks.',
    worldTags: ['system_approved', 'debug_sample'],
    authorId: 'usr_visual_audit',
    authorName: 'WorldShot Log',
    releaseStatus: 'public',
    imageUrl: pathToFileURL(samplePhotoPath).href,
    thumbnailImageUrl: pathToFileURL(samplePhotoPath).href,
    fetchStatus: 'success',
  });

  db.close();

  return { appData, localAppData, samplePhotoPath };
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForWindow() {
  const startedAt = Date.now();
  while (Date.now() - startedAt < 20000) {
    const win = BrowserWindow.getAllWindows().find((candidate) => !candidate.isDestroyed());
    if (win) {
      return win;
    }
    await wait(100);
  }
  throw new Error('Timed out waiting for the application window.');
}

async function waitForRendererReady(win) {
  const startedAt = Date.now();
  let lastState = null;
  while (Date.now() - startedAt < 20000) {
    const state = await win.webContents.executeJavaScript(
      `(() => {
        const appEl = document.querySelector('.app');
        return {
          readyState: document.readyState,
          hasBody: Boolean(document.body),
          hasI18n: Boolean(window.WorldShotI18n),
          initializing: Boolean(appEl?.classList.contains('is-app-initializing')),
          ariaBusy: appEl?.getAttribute('aria-busy') || null,
          photoCards: document.querySelectorAll('.photo-card').length,
          bodyText: (document.body?.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 160),
        };
      })()`
    );
    lastState = state;
    const ready =
      state.hasBody &&
      state.hasI18n &&
      !state.initializing &&
      state.readyState !== 'loading';
    if (ready) {
      return;
    }
    await wait(150);
  }
  throw new Error(
    `Timed out waiting for the renderer to finish initialization. Last state: ${JSON.stringify(lastState)}`
  );
}

function attachWindowDiagnostics(win) {
  win.webContents.on('console-message', (event) => {
    const level = event.level ?? 'unknown';
    const message = event.message ?? '';
    console.log(`[renderer console ${level}] ${message}`);
  });
  win.webContents.on('did-fail-load', (_event, errorCode, errorDescription) => {
    console.error(`[renderer did-fail-load] ${errorCode}: ${errorDescription}`);
  });
  win.webContents.on('render-process-gone', (_event, details) => {
    console.error(`[renderer process gone] ${JSON.stringify(details)}`);
  });
}

async function waitForSelector(win, selector, options = {}) {
  const timeoutMs = options.timeoutMs || 10000;
  const visible = options.visible !== false;
  const startedAt = Date.now();

  while (Date.now() - startedAt < timeoutMs) {
    const found = await win.webContents.executeJavaScript(
      `(() => {
        const el = document.querySelector(${JSON.stringify(selector)});
        if (!el) return false;
        if (!${visible ? 'true' : 'false'}) return true;
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      })()`
    );
    if (found) {
      return;
    }
    await wait(100);
  }

  throw new Error(`Timed out waiting for selector: ${selector}`);
}

async function runInPage(win, script) {
  return win.webContents.executeJavaScript(script);
}

async function capture(win, outputDir, language, name, report) {
  await wait(350);
  const screenshotPath = path.join(outputDir, `${language}-${name}.png`);
  const image = await win.capturePage();
  fs.writeFileSync(screenshotPath, image.toPNG());

  const checks = await runInPage(
    win,
    `(() => {
      const hasJapanese = ${HAS_JAPANESE_RE.toString()};
      const allowedUntranslatedTexts = new Set(${JSON.stringify(ALLOWED_UNTRANSLATED_TEXTS)});
      const isVisible = (el) => {
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      };
      const textIssues = [];
      const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const node = walker.currentNode;
        const text = (node.nodeValue || '').replace(/\\s+/g, ' ').trim();
        const parent = node.parentElement;
        if (!text || !parent || !isVisible(parent) || parent.closest('[data-i18n-ignore], .material-symbols-outlined')) {
          continue;
        }
        if (allowedUntranslatedTexts.has(text)) {
          continue;
        }
        if (hasJapanese.test(text)) {
          const rect = parent.getBoundingClientRect();
          textIssues.push({
            text,
            selector: parent.id ? '#' + parent.id : parent.className ? '.' + String(parent.className).split(/\\s+/).join('.') : parent.tagName.toLowerCase(),
            x: Math.round(rect.x),
            y: Math.round(rect.y),
          });
        }
      }

      const overflowIssues = [];
      document.querySelectorAll('button, input, select, textarea, .photo-card-world, .modal-world-link, .modal-world-description, .photo-editor-accordion-toggle').forEach((el) => {
        if (!isVisible(el)) return;
        const rect = el.getBoundingClientRect();
        if (el.scrollWidth > el.clientWidth + 2 || el.scrollHeight > el.clientHeight + 2) {
          overflowIssues.push({
            text: (el.innerText || el.value || el.getAttribute('aria-label') || '').replace(/\\s+/g, ' ').trim().slice(0, 80),
            selector: el.id ? '#' + el.id : el.className ? '.' + String(el.className).split(/\\s+/).join('.') : el.tagName.toLowerCase(),
            width: Math.round(rect.width),
            height: Math.round(rect.height),
            scrollWidth: el.scrollWidth,
            scrollHeight: el.scrollHeight,
            clientWidth: el.clientWidth,
            clientHeight: el.clientHeight,
          });
        }
      });

      return {
        url: location.href,
        language: document.body.getAttribute('data-language'),
        textIssues,
        overflowIssues,
      };
    })()`
  );

  report.screenshots.push({
    language,
    name,
    path: screenshotPath,
    textIssueCount: checks.textIssues.length,
    overflowIssueCount: checks.overflowIssues.length,
    checks,
  });
}

async function openMain(win, language) {
  await runInPage(
    win,
    `window.WorldShotI18n.setLanguage(${JSON.stringify(language)});`
  );
  await waitForSelector(win, '.photo-card');
}

async function openSettings(win) {
  await runInPage(win, `document.getElementById('settings-btn')?.click();`);
  await waitForSelector(win, '#settings-modal:not(.hidden)');
}

async function closeSettings(win) {
  await runInPage(win, `document.getElementById('settings-modal-close')?.click();`);
  await wait(250);
}

async function openImageModal(win) {
  await runInPage(win, `document.querySelector('.photo-card')?.click();`);
  await waitForSelector(win, '#image-modal:not(.hidden)');
}

async function closeImageModal(win) {
  await runInPage(win, `document.getElementById('image-modal-close')?.click();`);
  await wait(250);
}

async function openPhotoEditor(win) {
  await runInPage(win, `document.getElementById('modal-edit-photo-btn')?.click();`);
  await waitForSelector(win, '#photo-editor-modal:not(.hidden)', { timeoutMs: 15000 });
  await wait(900);
}

async function closePhotoEditor(win) {
  await runInPage(win, `document.getElementById('photo-editor-close')?.click();`);
  await wait(250);
}

async function main() {
  const outputDir = path.join(REPO_ROOT, 'out', 'ui-screenshots', RUN_ID);
  ensureDir(outputDir);

  const seeded = seedIsolatedData(SMOKE_ROOT);

  const report = {
    runId: RUN_ID,
    outputDir,
    smokeRoot: SMOKE_ROOT,
    samplePhotoPath: seeded.samplePhotoPath,
    viewport: SCREENSHOT_VIEWPORT,
    screenshots: [],
  };

  require('../src/index.js');

  const win = await waitForWindow();
  attachWindowDiagnostics(win);
  win.setSize(SCREENSHOT_VIEWPORT.width, SCREENSHOT_VIEWPORT.height);
  await waitForRendererReady(win);

  for (const language of LANGUAGES) {
    await openMain(win, language);
    await capture(win, outputDir, language, 'main', report);

    await openSettings(win);
    await capture(win, outputDir, language, 'settings', report);
    await closeSettings(win);

    await openImageModal(win);
    await capture(win, outputDir, language, 'card', report);

    await openPhotoEditor(win);
    await capture(win, outputDir, language, 'editor', report);
    await closePhotoEditor(win);
    await closeImageModal(win);
  }

  const reportPath = path.join(outputDir, 'report.json');
  fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
  console.log(`UI screenshots written to: ${outputDir}`);
  console.log(`UI screenshot report: ${reportPath}`);

  app.quit();
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
  app.quit();
});

const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { pathToFileURL } = require('node:url');

const { app, BrowserWindow } = require('electron');
const { initDatabase } = require('../src/db');

const REPO_ROOT = path.resolve(__dirname, '..');
const VIEWPORT = { width: 1440, height: 920 };
const RUN_ID = createRunId();
const SKIP_SCREENSHOTS = process.argv.includes('--skip-screenshots');
const SMOKE_ROOT = path.join(os.tmpdir(), `worldshot-coordinate-space-${RUN_ID}`);
const SMOKE_APPDATA = path.join(SMOKE_ROOT, 'AppData', 'Roaming');
const SMOKE_LOCALAPPDATA = path.join(SMOKE_ROOT, 'AppData', 'Local');
const OUTPUT_DIR = path.join(REPO_ROOT, 'out', 'smoke-coordinate-space', RUN_ID);

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
    'VRChat_2026-06-07_22-20-00.000_3840x2160.png'
  );
  fs.copyFileSync(path.join(REPO_ROOT, 'img', 'logo.png'), samplePhotoPath);

  const db = initDatabase(path.join(dataDir, 'app.sqlite'));
  const nowIso = new Date().toISOString();
  const row = db.insertOrUpdatePhoto({
    filePath: samplePhotoPath,
    fileName: path.basename(samplePhotoPath),
    fileHash: 'coordinate-space-smoke-photo',
    takenAt: '2026/06/07 22:20:00',
    takenAtTimestamp: new Date(2026, 5, 7, 22, 20, 0).getTime(),
    groupDate: '2026-06-07',
    year: 2026,
    month: 6,
    day: 7,
    worldId: 'wrld_coordinate_space_smoke',
    worldName: 'coordinate space smoke',
    worldNameManual: null,
    worldUrl: 'https://vrchat.com/home/world/wrld_coordinate_space_smoke',
    thumbnailPath: samplePhotoPath,
    imageWidth: 3840,
    imageHeight: 2160,
    resolutionTier: '4K',
    orientationTier: 'landscape',
    printNoteText: '',
    memoText: '',
    createdAt: nowIso,
    updatedAt: nowIso,
  });

  if (row?.id) {
    db.replacePhotoTags(row.id, [
      { name: 'coordinate smoke', colorHex: '#14B8A6' },
    ]);
  }

  db.upsertWorldMetadata({
    worldId: 'wrld_coordinate_space_smoke',
    sourceUrl: 'https://vrchat.com/home/world/wrld_coordinate_space_smoke',
    worldNameOfficial: 'coordinate space smoke',
    worldDescription: 'Smoke test sample for photo editor coordinate spaces.',
    worldTags: ['debug_sample'],
    authorId: 'usr_coordinate_space_smoke',
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
      `(() => {
        const appEl = document.querySelector('.app');
        return {
          readyState: document.readyState,
          hasBody: Boolean(document.body),
          hasI18n: Boolean(window.WorldShotI18n),
          initializing: Boolean(appEl?.classList.contains('is-app-initializing')),
          photoCards: document.querySelectorAll('.photo-card').length,
        };
      })()`
    );
    lastState = state;

    if (
      state.hasBody &&
      state.hasI18n &&
      (!state.initializing || state.photoCards > 0) &&
      state.readyState !== 'loading'
    ) {
      return;
    }

    await wait(150);
  }

  throw new Error(
    `Timed out waiting for renderer. Last state: ${JSON.stringify(lastState)}`
  );
}

async function waitForSelector(win, selector, options = {}) {
  const timeoutMs = options.timeoutMs || 10000;
  const visible = options.visible !== false;
  const startedAt = Date.now();

  while (Date.now() - startedAt < timeoutMs) {
    const found = await runInPage(
      win,
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

async function waitForPagePredicate(win, script, label, timeoutMs = 10000) {
  const startedAt = Date.now();
  let lastValue = null;

  while (Date.now() - startedAt < timeoutMs) {
    lastValue = await runInPage(win, script);

    if (lastValue) {
      return lastValue;
    }

    await wait(100);
  }

  throw new Error(
    `Timed out waiting for ${label}. Last value: ${JSON.stringify(lastValue)}`
  );
}

function attachWindowDiagnostics(win) {
  win.webContents.on('console-message', (event) => {
    const level = event.level ?? 'unknown';
    const message = event.message ?? '';
    console.log(`[renderer console ${level}] ${message}`);
  });
  win.webContents.on('render-process-gone', (_event, details) => {
    console.error(`[renderer process gone] ${JSON.stringify(details)}`);
  });
}

async function capture(win, name) {
  if (SKIP_SCREENSHOTS) {
    return null;
  }

  await wait(300);
  const image = await win.capturePage();
  const screenshotPath = path.join(OUTPUT_DIR, `${name}.png`);
  fs.writeFileSync(screenshotPath, image.toPNG());
  return screenshotPath;
}

async function openPhotoEditor(win) {
  await waitForSelector(win, '.photo-card', {
    timeoutMs: 30000,
    visible: false,
  });
  await runInPage(win, `document.querySelector('.photo-card')?.click();`);
  await waitForSelector(win, '#image-modal:not(.hidden)');
  await runInPage(win, `document.getElementById('modal-edit-photo-btn')?.click();`);
  await waitForSelector(win, '#photo-editor-modal:not(.hidden)', {
    timeoutMs: 15000,
  });
  await wait(900);
}

async function setupCoordinateFixture(win) {
  return runInPage(
    win,
    `new Promise((resolve) => {
      updatePhotoEditorPreviewZoom(1);
      setPhotoEditorAccordionOpen('text', true);
      const outputPoint = { x: 0.64, y: 0.48 };
      const sourcePoint = getPhotoEditorCurrentOutputPointAsSpace(outputPoint, 'source');
      const sourceScale = getPhotoEditorCurrentSourceScaleAtPoint(outputPoint);
      const text = addPhotoEditorTextOverlay({
        text: 'COORD',
        enabled: true,
        x: sourcePoint.x,
        y: sourcePoint.y,
        size: 92 / sourceScale.scale,
        strokeWidth: 5 / sourceScale.scale,
        color: '#ffffff',
        strokeColor: '#0f172a',
        space: 'source',
      });

      const canvas = document.createElement('canvas');
      canvas.width = 120;
      canvas.height = 90;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = 'rgba(20, 184, 166, 0.88)';
      ctx.fillRect(0, 0, 120, 90);
      ctx.fillStyle = 'rgba(15, 23, 42, 0.92)';
      ctx.fillRect(12, 12, 96, 66);
      ctx.fillStyle = '#f8fafc';
      ctx.font = '700 22px sans-serif';
      ctx.fillText('IMG', 36, 54);
      const dataUrl = canvas.toDataURL('image/png');
      const asset = {
        id: 'coordinate-overlay',
        fileName: 'coordinate-overlay.png',
        fileUrl: dataUrl,
        width: 120,
        height: 90,
      };
      const overlayOutputPoint = { x: 0.42, y: 0.58 };
      const overlaySourcePoint = getPhotoEditorCurrentOutputPointAsSpace(
        overlayOutputPoint,
        'source'
      );
      const overlayScale = getPhotoEditorCurrentSourceScaleAtPoint(overlayOutputPoint);
      const overlay = PhotoEditorImageOverlay.getDefaultImageOverlayState(asset, {
        x: overlaySourcePoint.x,
        y: overlaySourcePoint.y,
        width: 0.18 / overlayScale.scaleX,
        height: 0.135 / overlayScale.scaleY,
        space: 'source',
      });
      setPhotoEditorImageOverlayCollection([overlay], overlay.id);
      getPhotoEditorOverlayImageCacheEntry(overlay);

      const blurOutputPoint = { x: 0.52, y: 0.38 };
      const blurSourcePoint = getPhotoEditorCurrentOutputPointAsSpace(
        blurOutputPoint,
        'source'
      );
      const blurScale = getPhotoEditorCurrentSourceScaleAtPoint(blurOutputPoint);
      updatePhotoEditorBlur({
        mode: 'radial',
        amount: 58,
        centerX: blurSourcePoint.x,
        centerY: blurSourcePoint.y,
        radius: 0.14 / blurScale.scale,
        outerRadius: 0.27 / blurScale.scale,
        space: 'source',
      });

      requestAnimationFrame(() => {
        requestAnimationFrame(() => resolve({
          textId: text.id,
          overlayId: overlay.id,
        }));
      });
    })`
  );
}

async function readCoordinateState(win) {
  return runInPage(
    win,
    `(() => {
      const canvas = document.getElementById('photo-editor-canvas');
      const width = Math.max(1, canvas.width || 1);
      const height = Math.max(1, canvas.height || 1);
      const ctx = canvas.getContext('2d');
      const text = getPhotoEditorActiveTextOverlay();
      const imageOverlay = getPhotoEditorActiveImageOverlay();
      const textMetrics = getPhotoEditorTextCanvasMetrics(ctx, width, height, text);
      const imageMetrics = getPhotoEditorImageOverlayCanvasMetrics(
        width,
        height,
        imageOverlay
      );
      const blur = getPhotoEditorBlurGeometry(width, height, photoEditorState.blur);
      return {
        crop: { ...photoEditorState.crop },
        text: {
          space: text?.space,
          x: textMetrics.centerX / width,
          y: textMetrics.centerY / height,
          width: textMetrics.width / width,
        },
        image: {
          space: imageOverlay?.space,
          x: imageMetrics.centerX / width,
          y: imageMetrics.centerY / height,
          width: imageMetrics.width / width,
        },
        blur: {
          space: photoEditorState.blur?.space,
          x: blur.centerX / width,
          y: blur.centerY / height,
          radius: blur.radius / Math.max(1, Math.min(width, height)),
        },
      };
    })()`
  );
}

async function assertImageOverlayResizeWorks(win) {
  const result = await runInPage(
    win,
    `new Promise((resolve) => {
      setPhotoEditorAccordionOpen('imageOverlay', true);
      requestAnimationFrame(() => {
        const canvas = document.getElementById('photo-editor-canvas');
        const width = Math.max(1, canvas.width || 1);
        const height = Math.max(1, canvas.height || 1);
        const beforeOverlay = getPhotoEditorActiveImageOverlay();
        const beforeMetrics = getPhotoEditorImageOverlayCanvasMetrics(
          width,
          height,
          beforeOverlay
        );
        const handle = getPhotoEditorImageOverlayHandles(beforeMetrics).resize;
        const rect = canvas.getBoundingClientRect();
        const scaleX = rect.width / width;
        const scaleY = rect.height / height;
        const startX = rect.left + handle.x * scaleX;
        const startY = rect.top + handle.y * scaleY;
        const options = {
          bubbles: true,
          cancelable: true,
          pointerId: 91,
          pointerType: 'mouse',
          isPrimary: true,
        };
        canvas.dispatchEvent(new PointerEvent('pointerdown', {
          ...options,
          clientX: startX,
          clientY: startY,
          buttons: 1,
          button: 0,
        }));
        canvas.dispatchEvent(new PointerEvent('pointermove', {
          ...options,
          clientX: startX + 90,
          clientY: startY + 60,
          buttons: 1,
          button: 0,
        }));
        canvas.dispatchEvent(new PointerEvent('pointerup', {
          ...options,
          clientX: startX + 90,
          clientY: startY + 60,
          buttons: 0,
          button: 0,
        }));
        requestAnimationFrame(() => {
          const afterOverlay = getPhotoEditorActiveImageOverlay();
          resolve({
            before: {
              width: beforeOverlay.width,
              height: beforeOverlay.height,
            },
            after: {
              width: afterOverlay.width,
              height: afterOverlay.height,
            },
          });
        });
      });
    })`
  );

  if (
    result.after.width <= result.before.width * 1.05 ||
    result.after.height <= result.before.height * 1.05
  ) {
    throw new Error(
      `Image overlay resize did not increase size: ${JSON.stringify(result)}`
    );
  }
}

function assertMovedWithImage(before, after, key) {
  const deltaX = Math.abs(after[key].x - before[key].x);
  const deltaY = Math.abs(after[key].y - before[key].y);

  if (deltaX < 0.045 && deltaY < 0.025) {
    throw new Error(
      `${key} did not move with crop pan/zoom: ${JSON.stringify({ before: before[key], after: after[key] })}`
    );
  }
}

async function main() {
  console.log(`[coordinate-smoke] output: ${OUTPUT_DIR}`);
  const seeded = seedIsolatedData();
  console.log('[coordinate-smoke] data seeded');
  require('../src/index.js');

  const win = await waitForWindow();
  console.log('[coordinate-smoke] window ready');
  attachWindowDiagnostics(win);
  win.setSize(VIEWPORT.width, VIEWPORT.height);
  await waitForRendererReady(win);
  console.log('[coordinate-smoke] renderer ready');
  await openPhotoEditor(win);
  console.log('[coordinate-smoke] photo editor open');
  await setupCoordinateFixture(win);
  console.log('[coordinate-smoke] fixture applied');
  await waitForPagePredicate(
    win,
    `(() => {
      const canvas = document.getElementById('photo-editor-canvas');
      return Boolean(
        photoEditorState?.textOverlays?.length &&
        photoEditorState?.imageOverlays?.length &&
        canvas?.width > 0
      );
    })()`,
    'coordinate fixture'
  );

  const before = await readCoordinateState(win);
  console.log(`[coordinate-smoke] before: ${JSON.stringify(before)}`);
  const beforePath = await capture(win, 'before-crop-pan');

  await runInPage(
    win,
    `(() => {
      updatePhotoEditorCrop({
        preset: 'original',
        zoom: 68,
        offsetX: 58,
        offsetY: -34,
      });
      return true;
    })()`
  );
  await wait(900);

  const after = await readCoordinateState(win);
  console.log(`[coordinate-smoke] after: ${JSON.stringify(after)}`);
  const afterPath = await capture(win, 'after-crop-pan');

  for (const key of ['text', 'image', 'blur']) {
    if (before[key].space !== 'source' || after[key].space !== 'source') {
      throw new Error(`${key} was not kept in source space.`);
    }
    assertMovedWithImage(before, after, key);
  }
  await assertImageOverlayResizeWorks(win);
  console.log('[coordinate-smoke] image overlay resize passed');

  const result = {
    ok: true,
    runId: RUN_ID,
    smokeRoot: SMOKE_ROOT,
    samplePhotoPath: seeded.samplePhotoPath,
    before,
    after,
    screenshots: [beforePath, afterPath],
  };
  const reportPath = path.join(OUTPUT_DIR, 'result.json');
  fs.writeFileSync(reportPath, JSON.stringify(result, null, 2));
  console.log(`Photo editor coordinate-space smoke passed: ${OUTPUT_DIR}`);
  app.quit();
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
  app.quit();
});

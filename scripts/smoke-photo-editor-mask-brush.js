const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { pathToFileURL } = require('node:url');

const { app, BrowserWindow } = require('electron');
const { initDatabase } = require('../src/db');

const REPO_ROOT = path.resolve(__dirname, '..');
const VIEWPORT = { width: 1440, height: 920 };
const RUN_ID = createRunId();
const SMOKE_ROOT = path.join(os.tmpdir(), `worldshot-mask-brush-${RUN_ID}`);
const SMOKE_APPDATA = path.join(SMOKE_ROOT, 'AppData', 'Roaming');
const SMOKE_LOCALAPPDATA = path.join(SMOKE_ROOT, 'AppData', 'Local');
const OUTPUT_DIR = path.join(REPO_ROOT, 'out', 'smoke-mask-brush', RUN_ID);

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
    fileHash: 'mask-brush-smoke-photo',
    takenAt: '2026/06/07 22:20:00',
    takenAtTimestamp: new Date(2026, 5, 7, 22, 20, 0).getTime(),
    groupDate: '2026-06-07',
    year: 2026,
    month: 6,
    day: 7,
    worldId: 'wrld_mask_brush_smoke',
    worldName: 'mask brush smoke',
    worldNameManual: null,
    worldUrl: 'https://vrchat.com/home/world/wrld_mask_brush_smoke',
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
      { name: 'mask brush smoke', colorHex: '#6D5EF6' },
    ]);
  }

  db.upsertWorldMetadata({
    worldId: 'wrld_mask_brush_smoke',
    sourceUrl: 'https://vrchat.com/home/world/wrld_mask_brush_smoke',
    worldNameOfficial: 'mask brush smoke',
    worldDescription: 'Smoke test sample for subject mask brush editing.',
    worldTags: ['debug_sample'],
    authorId: 'usr_mask_brush_smoke',
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

async function openSubjectPanel(win) {
  await runInPage(
    win,
    `(() => {
      const toggle = document.querySelector('[data-photo-editor-accordion-toggle="subject"]');
      if (toggle?.getAttribute('aria-expanded') !== 'true') {
        toggle?.click();
      }
    })()`
  );
  await waitForSelector(win, '#photo-editor-subject-panel');
}

async function importTestMask(win) {
  await runInPage(
    win,
    `new Promise((resolve, reject) => {
      const canvas = document.createElement('canvas');
      canvas.width = 100;
      canvas.height = 100;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = 'rgb(0, 0, 0)';
      ctx.fillRect(0, 0, 100, 100);
      canvas.toBlob((blob) => {
        try {
          const input = document.getElementById('photo-editor-subject-file');
          const dataTransfer = new DataTransfer();
          dataTransfer.items.add(new File([blob], 'mask-smoke.png', { type: 'image/png' }));
          input.files = dataTransfer.files;
          input.dispatchEvent(new Event('change', { bubbles: true }));
          resolve(true);
        } catch (error) {
          reject(error);
        }
      }, 'image/png');
    })`
  );
  await waitForPagePredicate(
    win,
    `(() => {
      return Boolean(
        photoEditorState?.subjectMask?.enabled &&
        photoEditorState.subjectMask.maskDataUrl &&
        !document.getElementById('photo-editor-subject-brush-add')?.disabled
      );
    })()`,
    'imported subject mask'
  );
}

async function getSubjectMaskDataUrl(win) {
  return runInPage(
    win,
    `(() => photoEditorState?.subjectMask?.maskDataUrl || '')()`
  );
}

async function dispatchBrushStroke(win, mode, point) {
  await runInPage(
    win,
    `(() => {
      const button = document.getElementById(${JSON.stringify(
        mode === 'erase'
          ? 'photo-editor-subject-brush-erase'
          : 'photo-editor-subject-brush-add'
      )});
      const expectedMode = ${JSON.stringify(mode === 'erase' ? 'erase' : 'add')};
      if (photoEditorState?.subjectMaskBrushMode !== expectedMode) {
        button?.click();
      }
      const canvas = document.getElementById('photo-editor-canvas');
      const rect = canvas.getBoundingClientRect();
      const startX = rect.left + rect.width * ${point.x};
      const startY = rect.top + rect.height * ${point.y};
      const endX = rect.left + rect.width * ${point.x + 0.05};
      const endY = rect.top + rect.height * ${point.y + 0.05};
      const options = {
        bubbles: true,
        cancelable: true,
        pointerId: 77,
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
        clientX: endX,
        clientY: endY,
        buttons: 1,
        button: 0,
      }));
      canvas.dispatchEvent(new PointerEvent('pointerup', {
        ...options,
        clientX: endX,
        clientY: endY,
        buttons: 0,
        button: 0,
      }));
      return true;
    })()`
  );
  await wait(350);
}

async function dispatchBrushHover(win, point) {
  return runInPage(
    win,
    `(() => {
      const addButton = document.getElementById('photo-editor-subject-brush-add');
      if (photoEditorState?.subjectMaskBrushMode !== 'add') {
        addButton?.click();
      }
      const canvas = document.getElementById('photo-editor-canvas');
      const rect = canvas.getBoundingClientRect();
      canvas.dispatchEvent(new PointerEvent('pointermove', {
        bubbles: true,
        cancelable: true,
        pointerId: 88,
        pointerType: 'mouse',
        isPrimary: true,
        clientX: rect.left + rect.width * ${point.x},
        clientY: rect.top + rect.height * ${point.y},
        buttons: 0,
        button: 0,
      }));
      return {
        hasPoint: Boolean(photoEditorState?.subjectMaskBrushPreviewPoint),
        brushMode: photoEditorState?.subjectMaskBrushMode,
        dragMode: photoEditorState?.dragMode || '',
        point: photoEditorState?.subjectMaskBrushPreviewPoint || null,
      };
    })()`
  );
}

async function setPreviewZoomAndVerifyScroll(win) {
  await wait(150);
  const result = await runInPage(
    win,
    `(() => {
      const addButton = document.getElementById('photo-editor-subject-brush-add');
      if (photoEditorState?.subjectMaskBrushMode !== 'add') {
        addButton?.click();
      }
      const canvas = document.getElementById('photo-editor-canvas');
      const rect = canvas.getBoundingClientRect();
      for (let index = 0; index < 14; index += 1) {
        canvas.dispatchEvent(new WheelEvent('wheel', {
          bubbles: true,
          cancelable: true,
          deltaY: -120,
          clientX: rect.left + rect.width * 0.5,
          clientY: rect.top + rect.height * 0.5,
        }));
      }
      const wrap = document.getElementById('photo-editor-canvas-wrap');
      const before = {
        scrollLeft: wrap.scrollLeft,
        scrollTop: wrap.scrollTop,
        scrollWidth: wrap.scrollWidth,
        scrollHeight: wrap.scrollHeight,
        clientWidth: wrap.clientWidth,
        clientHeight: wrap.clientHeight,
        canvasWidth: canvas.getBoundingClientRect().width,
        previewZoom: photoEditorState?.previewZoom,
        brushMode: photoEditorState?.subjectMaskBrushMode,
        addDisabled: Boolean(addButton?.disabled),
        activeElement: document.activeElement?.id || '',
      };
      return before;
    })()`
  );

  if (
    result.scrollWidth <= result.clientWidth ||
    result.canvasWidth <= 0 ||
    result.previewZoom < 3
  ) {
    throw new Error(`Preview zoom did not create a scrollable canvas: ${JSON.stringify(result)}`);
  }

  await runInPage(
    win,
    `(() => {
      updatePhotoEditorPreviewZoom(1);
      const wrap = document.getElementById('photo-editor-canvas-wrap');
      wrap.scrollLeft = 0;
      wrap.scrollTop = 0;
      return true;
    })()`
  );
  await wait(120);
}

async function sampleBrushOpacity(win) {
  const result = await runInPage(
    win,
    `waitForPhotoEditorSubjectMaskToLoad().then(() => {
      const sampleButton = document.getElementById('photo-editor-subject-brush-sample');
      const opacityInput = document.getElementById('photo-editor-subject-brush-opacity');
      const canvas = document.getElementById('photo-editor-canvas');
      sampleButton.click();
      const rect = canvas.getBoundingClientRect();
      canvas.dispatchEvent(new PointerEvent('pointerdown', {
        bubbles: true,
        cancelable: true,
        pointerId: 77,
        pointerType: 'mouse',
        isPrimary: true,
        clientX: rect.left + rect.width * 0.18,
        clientY: rect.top + rect.height * 0.18,
        buttons: 1,
        button: 0,
      }));
      canvas.dispatchEvent(new PointerEvent('pointerup', {
        bubbles: true,
        cancelable: true,
        pointerId: 77,
        pointerType: 'mouse',
        isPrimary: true,
        clientX: rect.left + rect.width * 0.18,
        clientY: rect.top + rect.height * 0.18,
        buttons: 0,
        button: 0,
      }));
      return {
        opacity: Number(opacityInput.value),
        sampleMode: Boolean(photoEditorState?.subjectMaskBrushSampleMode),
      };
    })`
  );

  if (result.sampleMode || result.opacity < 1 || result.opacity > 100) {
    throw new Error(`Brush opacity sampler failed: ${JSON.stringify(result)}`);
  }
}

async function dispatchMiddleButtonScroll(win) {
  const result = await runInPage(
    win,
    `(() => {
      const wrap = document.getElementById('photo-editor-canvas-wrap');
      const canvas = document.getElementById('photo-editor-canvas');
      updatePhotoEditorPreviewZoom(4);
      wrap.scrollLeft = 20;
      wrap.scrollTop = 20;
      const rect = canvas.getBoundingClientRect();
      const startX = rect.left + Math.min(rect.width - 10, 300);
      const startY = rect.top + Math.min(rect.height - 10, 300);
      canvas.dispatchEvent(new PointerEvent('pointerdown', {
        bubbles: true,
        cancelable: true,
        pointerId: 99,
        pointerType: 'mouse',
        isPrimary: true,
        clientX: startX,
        clientY: startY,
        buttons: 4,
        button: 1,
      }));
      canvas.dispatchEvent(new PointerEvent('pointermove', {
        bubbles: true,
        cancelable: true,
        pointerId: 99,
        pointerType: 'mouse',
        isPrimary: true,
        clientX: startX - 90,
        clientY: startY - 70,
        buttons: 4,
        button: 1,
      }));
      canvas.dispatchEvent(new PointerEvent('pointerup', {
        bubbles: true,
        cancelable: true,
        pointerId: 99,
        pointerType: 'mouse',
        isPrimary: true,
        clientX: startX - 90,
        clientY: startY - 70,
        buttons: 0,
        button: 1,
      }));
      return {
        scrollLeft: wrap.scrollLeft,
        scrollTop: wrap.scrollTop,
      };
    })()`
  );

  if (result.scrollLeft <= 20 || result.scrollTop <= 20) {
    throw new Error(`Middle-button canvas scroll did not move: ${JSON.stringify(result)}`);
  }
}

async function assertTextCssOutlineFollowsPreviewZoom(win) {
  const result = await runInPage(
    win,
    `new Promise((resolve, reject) => {
      updatePhotoEditorPreviewZoom(1);
      setPhotoEditorAccordionOpen('text', true);
      const outputPoint = { x: 0.72, y: 0.46 };
      const sourcePoint = getPhotoEditorCurrentOutputPointAsSpace(outputPoint, 'source');
      const sourceScale = getPhotoEditorCurrentSourceScaleAtPoint(outputPoint);
      const text = addPhotoEditorTextOverlay({
        text: 'テキスト',
        enabled: true,
        x: sourcePoint.x,
        y: sourcePoint.y,
        size: 90 / sourceScale.scale,
        space: 'source',
      });
      updatePhotoEditorTextOverlay(
        {
          id: text.id,
          text: 'テキスト',
          enabled: true,
          x: sourcePoint.x,
          y: sourcePoint.y,
          size: 90 / sourceScale.scale,
          space: 'source',
        },
        { interactive: false }
      );
      const waitForOutline = (attempt = 0) => {
        const readCenter = () => {
          const canvas = document.getElementById('photo-editor-canvas');
          const outline = document.getElementById('photo-editor-text-outline');
          const canvasRect = canvas.getBoundingClientRect();
          const outlineRect = outline.getBoundingClientRect();
          return {
            visible: outline.classList.contains('is-visible'),
            x: (outlineRect.left + outlineRect.width / 2 - canvasRect.left) /
              Math.max(1, canvasRect.width),
            y: (outlineRect.top + outlineRect.height / 2 - canvasRect.top) /
              Math.max(1, canvasRect.height),
          };
        };
        const before = readCenter();
        if (!before.visible && attempt < 20) {
          requestAnimationFrame(() => waitForOutline(attempt + 1));
          return;
        }
        if (!before.visible) {
          reject(new Error('Text outline did not become visible before preview zoom.'));
          return;
        }
        updatePhotoEditorPreviewZoom(4);
        requestAnimationFrame(() => {
          const after = readCenter();
          resolve({ before, after });
        });
      };
      requestAnimationFrame(() => waitForOutline());
    })`
  );

  const driftX = Math.abs(result.before.x - result.after.x);
  const driftY = Math.abs(result.before.y - result.after.y);

  if (
    !result.before.visible ||
    !result.after.visible ||
    driftX > 0.015 ||
    driftY > 0.015
  ) {
    throw new Error(`Text CSS outline did not follow preview zoom: ${JSON.stringify(result)}`);
  }
}

async function assertMaskEquals(win, expected, label) {
  const actual = await getSubjectMaskDataUrl(win);

  if (actual !== expected) {
    throw new Error(`${label}: subject mask did not match expected state.`);
  }
}

async function main() {
  const seeded = seedIsolatedData();
  require('../src/index.js');

  const win = await waitForWindow();
  attachWindowDiagnostics(win);
  win.setSize(VIEWPORT.width, VIEWPORT.height);
  await waitForRendererReady(win);

  await openPhotoEditor(win);
  await openSubjectPanel(win);
  await importTestMask(win);
  await wait(300);
  await sampleBrushOpacity(win);
  await runInPage(
    win,
    `(() => {
      const opacityInput = document.getElementById('photo-editor-subject-brush-opacity');
      opacityInput.value = '100';
      opacityInput.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    })()`
  );
  await setPreviewZoomAndVerifyScroll(win);
  const importMask = await getSubjectMaskDataUrl(win);
  const beforePath = await capture(win, 'subject-panel-before-brush');

  await runInPage(win, `document.getElementById('photo-editor-subject-brush-size').value = '1';`);
  await runInPage(
    win,
    `document.getElementById('photo-editor-subject-brush-size').dispatchEvent(new Event('input', { bubbles: true }));`
  );
  await runInPage(win, `document.getElementById('photo-editor-subject-brush-size').value = '30';`);
  await runInPage(
    win,
    `document.getElementById('photo-editor-subject-brush-size').dispatchEvent(new Event('input', { bubbles: true }));`
  );
  await dispatchBrushStroke(win, 'add', { x: 0.5, y: 0.5 });
  const afterAdd = await getSubjectMaskDataUrl(win);

  if (!afterAdd || afterAdd === importMask) {
    throw new Error('Add brush did not change the subject mask.');
  }

  await runInPage(win, `waitForPhotoEditorSubjectMaskToLoad()`);
  await dispatchBrushStroke(win, 'erase', { x: 0.5, y: 0.5 });
  const afterErase = await getSubjectMaskDataUrl(win);

  if (!afterErase || afterErase === afterAdd) {
    throw new Error('Erase brush did not change the subject mask.');
  }

  const hoverVisible = await runInPage(
    win,
    `(() => {
      if (photoEditorState?.subjectMaskBrushMode !== 'add') {
        document.getElementById('photo-editor-subject-brush-add')?.click();
      }
      return true;
    })()`
  );
  const hoverResult = await dispatchBrushHover(win, { x: 0.52, y: 0.52 });
  const hasPreviewPoint = Boolean(hoverResult?.hasPoint);

  if (!hoverVisible || !hasPreviewPoint) {
    throw new Error(
      `Brush hover preview point was not updated: ${JSON.stringify(hoverResult)}`
    );
  }

  await dispatchMiddleButtonScroll(win);

  await runInPage(win, `document.getElementById('photo-editor-undo-btn')?.click();`);
  await wait(350);
  await assertMaskEquals(win, afterAdd, 'Undo after erase');

  await runInPage(win, `document.getElementById('photo-editor-redo-btn')?.click();`);
  await wait(350);
  await assertMaskEquals(win, afterErase, 'Redo after erase');
  await assertTextCssOutlineFollowsPreviewZoom(win);

  const afterPath = await capture(win, 'subject-panel-after-brush');

  const result = {
    ok: true,
    runId: RUN_ID,
    smokeRoot: SMOKE_ROOT,
    samplePhotoPath: seeded.samplePhotoPath,
    screenshots: [beforePath, afterPath],
    importMaskBytes: importMask.length,
    afterAddBytes: afterAdd.length,
    afterEraseBytes: afterErase.length,
  };

  fs.writeFileSync(
    path.join(OUTPUT_DIR, 'result.json'),
    JSON.stringify(result, null, 2)
  );
  console.log(`Photo editor subject mask brush smoke passed: ${OUTPUT_DIR}`);
  app.quit();
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
  app.quit();
});

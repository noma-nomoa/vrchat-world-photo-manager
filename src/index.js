const {
  app,
  BrowserWindow,
  ipcMain,
  dialog,
  nativeImage,
  shell,
  Menu,
  autoUpdater,
  net,
} = require('electron');
const path = require('node:path');
const fsSync = require('node:fs');
const fs = require('node:fs/promises');
const os = require('node:os');
const { createHash } = require('node:crypto');
const { performance } = require('node:perf_hooks');
const { spawn } = require('node:child_process');
const { pathToFileURL } = require('node:url');
const ExifReader = require('exifreader');
const iconv = require('iconv-lite');
const { initDatabase } = require('./db');

if (require('electron-squirrel-startup')) {
  app.quit();
}

let photoDb = null;
let thumbnailDirPath = '';
let preferencesFilePath = '';
let mainWindowRef = null;
const APP_DISPLAY_NAME = 'WorldShot Log';
const APP_TITLE = `${APP_DISPLAY_NAME} v${app.getVersion()}`;
const UNKNOWN_WORLD_DISPLAY_NAME = 'ワールド名を取得できませんでした';
const APP_WINDOW_ICON_ICO_PATH = path.join(__dirname, '..', 'img', 'logo.ico');
const APP_WINDOW_ICON_PNG_PATH = path.join(__dirname, '..', 'img', 'logo.png');
const APP_ROAMING_DATA_ROOT =
  process.env.WORLDSHOT_APPDATA_ROOT || app.getPath('appData');
const APP_USER_DATA_PATH = path.join(APP_ROAMING_DATA_ROOT, APP_DISPLAY_NAME);
const APP_LOCAL_DATA_ROOT =
  process.env.WORLDSHOT_LOCALAPPDATA_ROOT ||
  process.env.LOCALAPPDATA ||
  path.join(app.getPath('appData'), '..', 'Local');
const APP_DEV_SESSION_DATA_ROOT = path.join(
  os.tmpdir(),
  `${APP_DISPLAY_NAME}-DevSessionData`
);
const APP_SESSION_DATA_PATH = app.isPackaged
  ? path.join(APP_LOCAL_DATA_ROOT, APP_DISPLAY_NAME, 'SessionData')
  : path.join(APP_DEV_SESSION_DATA_ROOT, `${Date.now()}-${process.pid}`);
const APP_DISK_CACHE_PATH = path.join(APP_SESSION_DATA_PATH, 'Cache');

app.setName(APP_DISPLAY_NAME);
app.setPath('userData', APP_USER_DATA_PATH);

// Electron/Chromium cache data should live in a dedicated local-only path so it
// does not collide with renamed roaming app folders or fail while moving cache
// directories during startup.
for (const staleCachePath of [
  path.join(APP_USER_DATA_PATH, 'Cache'),
  path.join(APP_USER_DATA_PATH, 'GPUCache'),
  path.join(APP_USER_DATA_PATH, 'Code Cache'),
  path.join(APP_USER_DATA_PATH, 'DawnCache'),
  path.join(APP_USER_DATA_PATH, 'DawnGraphiteCache'),
]) {
  try {
    fsSync.rmSync(staleCachePath, { recursive: true, force: true });
  } catch {
    // Browser caches are disposable. If cleanup fails, Chromium can still try
    // to recreate them under the dedicated sessionData path below.
  }
}
fsSync.mkdirSync(APP_SESSION_DATA_PATH, { recursive: true });
fsSync.mkdirSync(APP_DISK_CACHE_PATH, { recursive: true });
app.setPath('sessionData', APP_SESSION_DATA_PATH);
app.commandLine.appendSwitch('disk-cache-dir', APP_DISK_CACHE_PATH);
app.commandLine.appendSwitch('disable-http-cache');
app.commandLine.appendSwitch('disable-gpu-shader-disk-cache');

const LEGACY_APP_STORAGE_DIRNAME = 'vrchat-world-photo-manager';
const APP_STORAGE_DIRNAME = APP_DISPLAY_NAME;
const THUMBNAIL_DIRNAME = 'thumbnails';
const THUMBNAIL_SIZE = 320;
const THUMBNAIL_EXTENSION = '.jpg';
const THUMBNAIL_JPEG_QUALITY = 82;
const PHOTO_EDITOR_EXPORT_FORMATS = Object.freeze({
  png: {
    mimeType: 'image/png',
    extension: '.png',
    filterName: 'PNG Image',
  },
  jpeg: {
    mimeType: 'image/jpeg',
    extension: '.jpg',
    aliases: ['.jpeg'],
    filterName: 'JPEG Image',
  },
  webp: {
    mimeType: 'image/webp',
    extension: '.webp',
    filterName: 'WebP Image',
  },
});
const PHOTO_EDITOR_DATA_URL_PATTERN = /^data:(image\/(?:png|jpeg|webp));base64,/;
const WORLD_METADATA_API_BASE_URL = 'https://api.vrchat.cloud/api/1/worlds';
const WORLD_METADATA_FETCH_TIMEOUT_MS = 10000;
const IMPORT_PREPROCESS_CONCURRENCY = Math.max(
  4,
  Math.min(
    10,
    typeof os.availableParallelism === 'function'
      ? os.availableParallelism()
      : 4
  )
);
const RENDERER_ROW_PREP_CONCURRENCY = Math.max(
  2,
  Math.min(6, IMPORT_PREPROCESS_CONCURRENCY)
);
const PHOTO_PREPARATION_CACHE_LIMIT = 5000;
const PERFORMANCE_LOG_THRESHOLD_MS = Math.max(
  80,
  Number(process.env.WORLDSHOT_PERF_LOG_THRESHOLD_MS) || 250
);
const IS_PERFORMANCE_LOG_FORCED = process.env.WORLDSHOT_PERF_LOG === '1';
const DEFAULT_PHOTO_LABEL_COLORS = [
  '#6D5EF6',
  '#4F8CFF',
  '#14B8A6',
  '#22C55E',
  '#F59E0B',
  '#F97316',
  '#EC4899',
  '#8B5CF6',
  '#EF4444',
  '#06B6D4',
];
const PHOTO_ORIENTATION_TIERS = new Set(['landscape', 'portrait', 'square']);
const VRCHAT_PRINT_PHOTO_ASPECT_RATIO = 2048 / 1440;
const VRCHAT_PRINT_PHOTO_ASPECT_TOLERANCE = 0.01;
const VRCHAT_PRINT_PHOTO_MIN_LONG_EDGE = 1400;

const SUPPORTED_IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.webp']);
const PROCESSING_PROGRESS_CHANNEL = 'processing-progress';
const WORLD_METADATA_UPDATED_CHANNEL = 'world-metadata-updated';
const APP_UPDATE_STATUS_CHANNEL = 'app-update-status';
const APP_UPDATE_ACTION_CHANNEL = 'app-update-action';
const WORLD_METADATA_SYNC_OPERATION = 'world-metadata-sync';
const WORLD_METADATA_SYNC_DELAY_MS = 1000;
const APP_UPDATE_CHECK_DELAY_MS = 3500;
const APP_UPDATE_REPOSITORY_OWNER = 'noma-nomoa';
const APP_UPDATE_REPOSITORY_NAME = 'vrchat-world-photo-manager';
const APP_UPDATE_RELEASE_API_URL = `https://api.github.com/repos/${APP_UPDATE_REPOSITORY_OWNER}/${APP_UPDATE_REPOSITORY_NAME}/releases/latest`;
const APP_UPDATE_SERVICE_BASE_URL = 'https://update.electronjs.org';
const DEFAULT_APP_PREFERENCES = Object.freeze({
  backgroundImagePath: '',
});
const APP_DATA_BACKUP_SCHEMA_VERSION = 1;
const PHOTO_CATALOG_EXPORT_COLUMNS = Object.freeze([
  ['id', 'ID'],
  ['fileName', 'File Name'],
  ['filePath', 'File Path'],
  ['takenAt', 'Taken At'],
  ['groupDate', 'Group Date'],
  ['year', 'Year'],
  ['month', 'Month'],
  ['day', 'Day'],
  ['worldId', 'World ID'],
  ['worldName', 'World Name'],
  ['worldUrl', 'World URL'],
  ['isFavorite', 'Favorite'],
  ['labels', 'Labels'],
  ['memoText', 'Memo'],
  ['printNoteText', 'Print Note'],
  ['imageWidth', 'Image Width'],
  ['imageHeight', 'Image Height'],
  ['resolutionTier', 'Resolution'],
  ['orientationTier', 'Orientation'],
]);

const pendingWorldMetadataSyncTargets = new Map();
const worldMetadataSyncSubscribers = new Set();
let isWorldMetadataSyncRunning = false;
let activeWorldMetadataSyncWorldId = null;
let isAutoUpdaterConfigured = false;
let isAutoUpdateCheckRunning = false;
let isAutoUpdateDownloadRunning = false;
let latestAvailableAppUpdateRelease = null;
const photoOrientationPreparationCache = new Map();
const photoPrintNotePreparationCache = new Map();

function roundDuration(durationMs) {
  return Math.round((Number(durationMs) || 0) * 10) / 10;
}

function shouldLogPerformance(durationMs) {
  return (
    IS_PERFORMANCE_LOG_FORCED ||
    (Number.isFinite(durationMs) && durationMs >= PERFORMANCE_LOG_THRESHOLD_MS)
  );
}

function logPerformanceTiming(label, durationMs, details = {}) {
  if (!shouldLogPerformance(durationMs)) {
    return;
  }

  const normalizedDetails = Object.fromEntries(
    Object.entries(details || {}).map(([key, value]) => [
      key,
      typeof value === 'number' ? roundDuration(value) : value,
    ])
  );

  console.info('[performance]', label, {
    durationMs: roundDuration(durationMs),
    ...normalizedDetails,
  });
}

// Lightweight app preferences are stored outside the renderer so they survive
// reloads/restarts even if the file:// localStorage origin changes.
function normalizeBackgroundImagePreferencePath(filePath) {
  return typeof filePath === 'string' ? filePath.trim() : '';
}

async function readAppPreferences() {
  if (!preferencesFilePath) {
    return { ...DEFAULT_APP_PREFERENCES };
  }

  try {
    const raw = await fs.readFile(preferencesFilePath, 'utf8');
    const parsed = JSON.parse(raw);
    return {
      ...DEFAULT_APP_PREFERENCES,
      ...(parsed && typeof parsed === 'object' ? parsed : {}),
      backgroundImagePath: normalizeBackgroundImagePreferencePath(
        parsed?.backgroundImagePath
      ),
    };
  } catch {
    return { ...DEFAULT_APP_PREFERENCES };
  }
}

async function writeAppPreferences(nextPreferences) {
  if (!preferencesFilePath) {
    return { ...DEFAULT_APP_PREFERENCES };
  }

  const normalizedPreferences = {
    ...DEFAULT_APP_PREFERENCES,
    ...(nextPreferences && typeof nextPreferences === 'object'
      ? nextPreferences
      : {}),
    backgroundImagePath: normalizeBackgroundImagePreferencePath(
      nextPreferences?.backgroundImagePath
    ),
  };

  await fs.writeFile(
    preferencesFilePath,
    JSON.stringify(normalizedPreferences, null, 2),
    'utf8'
  );

  return normalizedPreferences;
}

async function getBackgroundImagePreference() {
  const preferences = await readAppPreferences();
  return preferences.backgroundImagePath;
}

async function setBackgroundImagePreference(filePath) {
  const preferences = await readAppPreferences();
  preferences.backgroundImagePath = normalizeBackgroundImagePreferencePath(
    filePath
  );
  const savedPreferences = await writeAppPreferences(preferences);
  return savedPreferences.backgroundImagePath;
}

function getDatabaseDirectoryPath() {
  return preferencesFilePath ? path.dirname(preferencesFilePath) : '';
}

function formatFileTimestamp(date = new Date()) {
  const pad = (value) => String(value).padStart(2, '0');

  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate()),
    '-',
    pad(date.getHours()),
    pad(date.getMinutes()),
    pad(date.getSeconds()),
  ].join('');
}

function buildExportDefaultPath(fileName) {
  return path.join(app.getPath('documents'), fileName);
}

function buildAppDataBackupPayload() {
  const dbSnapshot = photoDb.createBackupSnapshot();

  return {
    appName: APP_DISPLAY_NAME,
    appVersion: app.getVersion(),
    backupSchemaVersion: APP_DATA_BACKUP_SCHEMA_VERSION,
    exportedAt: new Date().toISOString(),
    data: dbSnapshot,
  };
}

async function buildAppDataBackupPayloadWithPreferences() {
  return {
    ...buildAppDataBackupPayload(),
    preferences: await readAppPreferences(),
  };
}

async function createAutomaticAppDataBackupFile(reason = 'maintenance') {
  const normalizedReason =
    typeof reason === 'string' && reason.trim().length > 0
      ? reason.trim().replace(/[^a-z0-9_-]+/gi, '-').replace(/^-+|-+$/g, '')
      : 'maintenance';
  const backupDirectoryPath = path.join(getDatabaseDirectoryPath(), 'backups');
  const filePath = path.join(
    backupDirectoryPath,
    `worldshot-log-auto-backup-${normalizedReason}-${formatFileTimestamp()}.json`
  );
  const payload = await buildAppDataBackupPayloadWithPreferences();

  await ensureDir(backupDirectoryPath);
  await fs.writeFile(filePath, JSON.stringify(payload, null, 2), 'utf8');

  return {
    ok: true,
    filePath,
    fileName: path.basename(filePath),
    ...getBackupCounts(payload),
  };
}

function getBackupCounts(snapshot) {
  const counts = snapshot?.data?.counts || snapshot?.counts || {};
  const tables = snapshot?.data?.tables || snapshot?.tables || {};

  return {
    photoCount:
      Number(counts.photoCount) ||
      (Array.isArray(tables.photos) ? tables.photos.length : 0),
    trackedFolderCount:
      Number(counts.trackedFolderCount) ||
      (Array.isArray(tables.tracked_folders)
        ? tables.tracked_folders.length
        : 0),
    worldCacheCount:
      Number(counts.worldCacheCount) ||
      (Array.isArray(tables.world_metadata_cache)
        ? tables.world_metadata_cache.length
        : 0),
    tagCount:
      Number(counts.tagCount) ||
      (Array.isArray(tables.tags) ? tables.tags.length : 0),
  };
}

function csvEscape(value) {
  if (value === null || value === undefined) {
    return '';
  }

  const text = String(value);

  if (/[",\r\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }

  return text;
}

function buildCsvContent(rows, columns) {
  const header = columns.map(([, label]) => csvEscape(label)).join(',');
  const body = rows.map((row) =>
    columns.map(([key]) => csvEscape(row[key])).join(',')
  );

  return [header, ...body].join('\r\n');
}

function createPhotoCatalogExportRows() {
  const rows = attachPhotoLabelsToRows(photoDb.getAllPhotosWithWorldInfo());

  return rows.map((row) => {
    const worldMetadata = row.world_id
      ? photoDb.getWorldMetadataByWorldId(row.world_id)
      : null;
    const labels = (Array.isArray(row.photo_labels) ? row.photo_labels : [])
      .map(toRendererPhotoLabel)
      .filter(Boolean);
    const labelNames = labels.map((label) => label.name).filter(Boolean);
    const worldName =
      normalizeDisplayWorldName(row.world_name_manual || row.world_name) ||
      normalizeOfficialWorldText(worldMetadata?.world_name_official) ||
      '';

    return {
      id: row.id,
      fileName: row.file_name || '',
      filePath: row.file_path || '',
      takenAt: row.taken_at || '',
      groupDate: row.group_date || '',
      year: row.year || '',
      month: row.month || '',
      day: row.day || '',
      worldId: row.world_id || '',
      worldName,
      worldUrl: row.world_url || buildWorldUrlFromId(row.world_id) || '',
      isFavorite: row.is_favorite ? 'true' : 'false',
      labels: labelNames.join('; '),
      memoText: row.memo_text || '',
      printNoteText: normalizePhotoPrintNoteText(row.print_note_text) || '',
      imageWidth: row.image_width || '',
      imageHeight: row.image_height || '',
      resolutionTier: row.resolution_tier || '',
      orientationTier: row.orientation_tier || '',
    };
  });
}

function createHealthIssue(row, message, extra = {}) {
  return {
    photoId: row.id,
    fileName: row.file_name || '',
    filePath: row.file_path || '',
    worldId: row.world_id || '',
    worldName:
      normalizeDisplayWorldName(row.world_name_manual || row.world_name) || '',
    message,
    ...extra,
  };
}

function pushHealthIssue(target, issue, limit = 50) {
  if (target.length < limit) {
    target.push(issue);
  }
}

function getWorldMetadataHealthIssue(row, metadataCache = null) {
  const normalizedWorldId =
    normalizeWorldId(row.world_id) || parseWorldIdFromUrl(row.world_url);

  if (!normalizedWorldId) {
    return null;
  }

  let metadataRow;

  if (metadataCache instanceof Map) {
    if (!metadataCache.has(normalizedWorldId)) {
      metadataCache.set(
        normalizedWorldId,
        photoDb.getWorldMetadataByWorldId(normalizedWorldId)
      );
    }

    metadataRow = metadataCache.get(normalizedWorldId);
  } else {
    metadataRow = photoDb.getWorldMetadataByWorldId(normalizedWorldId);
  }

  if (metadataRow && metadataRow.fetch_status === 'success') {
    return null;
  }

  return {
    worldId: normalizedWorldId,
    fetchStatus: metadataRow?.fetch_status || 'missing',
    message:
      metadataRow?.fetch_error ||
      (metadataRow
        ? `Worldメタデータ取得状態: ${metadataRow.fetch_status || 'unknown'}`
        : 'Worldメタデータが未取得です'),
  };
}

async function analyzePhotoHealthRow(row, metadataCache = null) {
  const originalPath = normalizeKnownFilePath(row.file_path);
  const thumbnailPath =
    typeof row.thumbnail_path === 'string' ? row.thumbnail_path.trim() : '';
  const [originalExists, thumbnailExists] = await Promise.all([
    originalPath ? pathExists(originalPath) : Promise.resolve(false),
    thumbnailPath ? pathExists(thumbnailPath) : Promise.resolve(false),
  ]);
  const normalizedWorldId =
    normalizeWorldId(row.world_id) || parseWorldIdFromUrl(row.world_url);
  const displayWorldName = normalizeDisplayWorldName(
    row.world_name_manual || row.world_name
  );
  const missingWorldInfo =
    !normalizedWorldId && !displayWorldName && !row.world_url;
  const worldMetadataIssue = missingWorldInfo
    ? null
    : getWorldMetadataHealthIssue(row, metadataCache);

  return {
    row,
    originalPath,
    thumbnailPath,
    originalExists,
    thumbnailExists,
    missingOriginal: !originalExists,
    missingThumbnail: !thumbnailPath || !thumbnailExists,
    missingWorldInfo,
    worldMetadataIssue,
  };
}

function getHealthIssueRowsFromAnalyses(analyses, issueKind) {
  const normalizedIssueKind = normalizeHealthIssueKind(issueKind);

  return (Array.isArray(analyses) ? analyses : [])
    .filter((analysis) => {
      if (!analysis?.row) {
        return false;
      }

      if (normalizedIssueKind === 'missing-original') {
        return analysis.missingOriginal;
      }

      if (normalizedIssueKind === 'missing-thumbnail') {
        return analysis.missingThumbnail;
      }

      if (normalizedIssueKind === 'missing-world-info') {
        return analysis.missingWorldInfo;
      }

      if (normalizedIssueKind === 'world-metadata') {
        return Boolean(analysis.worldMetadataIssue);
      }

      return false;
    })
    .map((analysis) => analysis.row);
}

function normalizeHealthIssueKind(issueKind) {
  const normalizedKind =
    typeof issueKind === 'string' ? issueKind.trim() : '';

  if (normalizedKind === 'world-metadata-issues') {
    return 'world-metadata';
  }

  return [
    'missing-original',
    'missing-thumbnail',
    'missing-world-info',
    'world-metadata',
  ].includes(normalizedKind)
    ? normalizedKind
    : null;
}

async function analyzeApplicationDataHealthRows(rows) {
  const metadataCache = new Map();

  return mapWithConcurrencyLimit(
    Array.isArray(rows) ? rows : [],
    IMPORT_PREPROCESS_CONCURRENCY,
    (row) => analyzePhotoHealthRow(row, metadataCache)
  );
}

async function checkApplicationDataHealth() {
  const rows = photoDb.getAllPhotosWithWorldInfo();
  const analyses = await analyzeApplicationDataHealthRows(rows);
  const issues = {
    missingOriginalFiles: [],
    missingThumbnails: [],
    missingWorldInfo: [],
    worldMetadataIssues: [],
  };
  const summary = {
    totalPhotoCount: rows.length,
    missingOriginalCount: 0,
    missingThumbnailCount: 0,
    missingWorldInfoCount: 0,
    worldMetadataIssueCount: 0,
  };

  for (const analysis of analyses) {
    const row = analysis.row;

    if (analysis.missingOriginal) {
      summary.missingOriginalCount += 1;
      pushHealthIssue(
        issues.missingOriginalFiles,
        createHealthIssue(row, '元画像ファイルが見つかりません')
      );
    }

    if (analysis.missingThumbnail) {
      summary.missingThumbnailCount += 1;
      pushHealthIssue(
        issues.missingThumbnails,
        createHealthIssue(
          row,
          analysis.thumbnailPath
            ? 'サムネイルファイルが見つかりません'
            : 'サムネイルが未生成です',
          { thumbnailPath: analysis.thumbnailPath }
        )
      );
    }

    if (analysis.missingWorldInfo) {
      summary.missingWorldInfoCount += 1;
      pushHealthIssue(
        issues.missingWorldInfo,
        createHealthIssue(row, 'World情報が未取得です')
      );
      continue;
    }

    if (analysis.worldMetadataIssue) {
      summary.worldMetadataIssueCount += 1;
      pushHealthIssue(
        issues.worldMetadataIssues,
        createHealthIssue(row, analysis.worldMetadataIssue.message, {
          worldId: analysis.worldMetadataIssue.worldId,
          fetchStatus: analysis.worldMetadataIssue.fetchStatus,
        })
      );
    }
  }

  return {
    ok: true,
    checkedAt: new Date().toISOString(),
    ...summary,
    healthy:
      summary.missingOriginalCount === 0 &&
      summary.missingThumbnailCount === 0 &&
      summary.missingWorldInfoCount === 0 &&
      summary.worldMetadataIssueCount === 0,
    issues,
  };
}

async function getWorldMetadataIssuePhotos() {
  return getHealthIssuePhotos('world-metadata');
}

async function getHealthIssuePhotos(issueKind) {
  const normalizedIssueKind = normalizeHealthIssueKind(issueKind);

  if (!normalizedIssueKind) {
    throw new Error('状態チェックカテゴリが不正です');
  }

  const rows = photoDb.getAllPhotosWithWorldInfo();
  const analyses = await analyzeApplicationDataHealthRows(rows);
  const issueRows = getHealthIssueRowsFromAnalyses(
    analyses,
    normalizedIssueKind
  );
  const preparedRows = await prepareRowsForRenderer(issueRows);

  return {
    ok: true,
    issueKind: normalizedIssueKind,
    photoCount: preparedRows.length,
    photos: preparedRows.map(toRendererPhoto),
  };
}

function buildWorldMetadataSyncTargetsFromPhotoRows(rows) {
  const targets = new Map();

  for (const row of Array.isArray(rows) ? rows : []) {
    const syncTarget = createWorldMetadataSyncTarget({
      worldId: normalizeWorldId(row?.world_id) || parseWorldIdFromUrl(row?.world_url),
      worldUrl: row?.world_url,
    });

    if (!syncTarget) {
      continue;
    }

    if (!targets.has(syncTarget.worldId)) {
      targets.set(syncTarget.worldId, syncTarget);
    }
  }

  return Array.from(targets.values());
}

async function refreshWorldMetadataIssuePhotos(webContents) {
  const rows = photoDb.getAllPhotosWithWorldInfo();
  const analyses = await analyzeApplicationDataHealthRows(rows);
  const issueRows = getHealthIssueRowsFromAnalyses(analyses, 'world-metadata');
  const targets = buildWorldMetadataSyncTargetsFromPhotoRows(issueRows);

  if (targets.length === 0) {
    return {
      ok: true,
      photoCount: issueRows.length,
      targetCount: 0,
      queuedCount: 0,
      pendingCount: getQueuedWorldMetadataSyncCount(),
    };
  }

  const syncResult = startWorldMetadataSyncForSubscriber(targets, webContents);

  return {
    ...syncResult,
    photoCount: issueRows.length,
    targetCount: targets.length,
  };
}

async function regenerateMissingThumbnails(progressReporter = null) {
  const rows = photoDb.getAllPhotosWithWorldInfo();
  const analyses = await analyzeApplicationDataHealthRows(rows);
  const missingThumbnailRows = getHealthIssueRowsFromAnalyses(
    analyses,
    'missing-thumbnail'
  );

  return regenerateThumbnailRows(missingThumbnailRows, null, progressReporter, {
    missingOnly: true,
    progressMessage: '欠損サムネイルを再生成中...',
  });
}

async function createAppDataBackupFile() {
  const timestamp = formatFileTimestamp();
  const result = await dialog.showSaveDialog({
    title: 'WorldShot Log のバックアップを保存',
    defaultPath: buildExportDefaultPath(
      `worldshot-log-backup-${timestamp}.json`
    ),
    filters: [{ name: 'WorldShot Log Backup', extensions: ['json'] }],
  });

  if (result.canceled || !result.filePath) {
    return { ok: true, canceled: true };
  }

  const payload = await buildAppDataBackupPayloadWithPreferences();

  await fs.writeFile(result.filePath, JSON.stringify(payload, null, 2), 'utf8');

  return {
    ok: true,
    canceled: false,
    filePath: result.filePath,
    fileName: path.basename(result.filePath),
    ...getBackupCounts(payload),
  };
}

async function restoreAppDataBackupFile() {
  const result = await dialog.showOpenDialog({
    title: 'WorldShot Log のバックアップを選択',
    properties: ['openFile'],
    filters: [{ name: 'WorldShot Log Backup', extensions: ['json'] }],
  });

  if (result.canceled || result.filePaths.length === 0) {
    return { ok: true, canceled: true };
  }

  const filePath = result.filePaths[0];
  const raw = await fs.readFile(filePath, 'utf8');
  const parsed = JSON.parse(raw);

  if (!parsed || typeof parsed !== 'object') {
    throw new Error('バックアップファイルを読み込めませんでした');
  }

  if (parsed.appName && parsed.appName !== APP_DISPLAY_NAME) {
    throw new Error('WorldShot Log のバックアップではありません');
  }

  const automaticBackup = await createAutomaticAppDataBackupFile('before-restore');
  const restoreResult = photoDb.restoreBackupSnapshot(parsed.data || parsed);
  clearPhotoPreparationCaches();

  if (parsed.preferences && typeof parsed.preferences === 'object') {
    await writeAppPreferences(parsed.preferences);
  }

  return {
    ok: true,
    canceled: false,
    filePath,
    fileName: path.basename(filePath),
    automaticBackupFilePath: automaticBackup.filePath,
    automaticBackupFileName: automaticBackup.fileName,
    ...restoreResult.afterCounts,
    restoredPhotoCount: restoreResult.restoredPhotoCount,
    restoredTrackedFolderCount: restoreResult.restoredTrackedFolderCount,
    restoredWorldCacheCount: restoreResult.restoredWorldCacheCount,
    restoredTagCount: restoreResult.restoredTagCount,
  };
}

async function exportPhotoCatalogFile(format = 'csv') {
  const normalizedFormat = format === 'json' ? 'json' : 'csv';
  const timestamp = formatFileTimestamp();
  const result = await dialog.showSaveDialog({
    title:
      normalizedFormat === 'json'
        ? '写真一覧をJSONでエクスポート'
        : '写真一覧をCSVでエクスポート',
    defaultPath: buildExportDefaultPath(
      `worldshot-log-photos-${timestamp}.${normalizedFormat}`
    ),
    filters: [
      normalizedFormat === 'json'
        ? { name: 'JSON Files', extensions: ['json'] }
        : { name: 'CSV Files', extensions: ['csv'] },
    ],
  });

  if (result.canceled || !result.filePath) {
    return { ok: true, canceled: true, format: normalizedFormat };
  }

  const photos = createPhotoCatalogExportRows();

  if (normalizedFormat === 'json') {
    await fs.writeFile(
      result.filePath,
      JSON.stringify(
        {
          appName: APP_DISPLAY_NAME,
          appVersion: app.getVersion(),
          exportedAt: new Date().toISOString(),
          photos,
        },
        null,
        2
      ),
      'utf8'
    );
  } else {
    await fs.writeFile(
      result.filePath,
      buildCsvContent(photos, PHOTO_CATALOG_EXPORT_COLUMNS),
      'utf8'
    );
  }

  return {
    ok: true,
    canceled: false,
    format: normalizedFormat,
    filePath: result.filePath,
    fileName: path.basename(result.filePath),
    photoCount: photos.length,
  };
}

function getThumbnailStorageRootPath(dirname = APP_STORAGE_DIRNAME) {
  return path.join(app.getPath('home'), dirname);
}

function getManagedThumbnailRootPaths() {
  return Array.from(
    new Set(
      [APP_STORAGE_DIRNAME, LEGACY_APP_STORAGE_DIRNAME]
        .filter(Boolean)
        .map((dirname) => getThumbnailStorageRootPath(dirname))
    )
  );
}

function getManagedThumbnailDirectoryPaths() {
  return getManagedThumbnailRootPaths().map((rootPath) =>
    path.join(rootPath, THUMBNAIL_DIRNAME)
  );
}

async function migrateLegacyThumbnailStorage() {
  const legacyRootPath = getThumbnailStorageRootPath(LEGACY_APP_STORAGE_DIRNAME);
  const nextRootPath = getThumbnailStorageRootPath(APP_STORAGE_DIRNAME);

  if (legacyRootPath === nextRootPath) {
    return;
  }

  try {
    await fs.access(legacyRootPath);
  } catch {
    return;
  }

  try {
    await fs.access(nextRootPath);
    // If the new root already exists, prefer keeping it and just clean up any
    // orphaned legacy thumbnails later via regeneration/maintenance actions.
    return;
  } catch {
    // New storage root does not exist yet, so we can migrate the legacy cache.
  }

  try {
    await fs.rename(legacyRootPath, nextRootPath);
  } catch {
    await ensureDir(nextRootPath);

    const legacyThumbnailPath = path.join(legacyRootPath, THUMBNAIL_DIRNAME);
    const nextThumbnailPath = path.join(nextRootPath, THUMBNAIL_DIRNAME);

    try {
      await ensureDir(nextThumbnailPath);
      const entries = await fs.readdir(legacyThumbnailPath, { withFileTypes: true });

      for (const entry of entries) {
        if (!entry.isFile()) {
          continue;
        }

        const sourcePath = path.join(legacyThumbnailPath, entry.name);
        const destinationPath = path.join(nextThumbnailPath, entry.name);

        try {
          await fs.access(destinationPath);
        } catch {
          await fs.rename(sourcePath, destinationPath);
        }
      }

      await fs.rm(legacyRootPath, {
        recursive: true,
        force: true,
        maxRetries: 2,
        retryDelay: 120,
      });
    } catch {
      // Thumbnail cache is disposable. If migration fails, startup continues and
      // thumbnails can be recreated later.
    }
  }
}

function rebaseManagedThumbnailPath(thumbnailPath) {
  if (typeof thumbnailPath !== 'string' || thumbnailPath.trim().length === 0) {
    return null;
  }

  const normalizedTargetPath = path.resolve(thumbnailPath);
  const legacyThumbnailDirectoryPath = path.resolve(
    path.join(getThumbnailStorageRootPath(LEGACY_APP_STORAGE_DIRNAME), THUMBNAIL_DIRNAME)
  );

  if (
    normalizedTargetPath !== legacyThumbnailDirectoryPath &&
    !normalizedTargetPath.startsWith(`${legacyThumbnailDirectoryPath}${path.sep}`)
  ) {
    return null;
  }

  const currentThumbnailDirectoryPath = path.resolve(
    path.join(getThumbnailStorageRootPath(APP_STORAGE_DIRNAME), THUMBNAIL_DIRNAME)
  );
  const relativeThumbnailPath = path.relative(
    legacyThumbnailDirectoryPath,
    normalizedTargetPath
  );

  return path.join(currentThumbnailDirectoryPath, relativeThumbnailPath);
}

async function reconcileManagedThumbnailPaths() {
  if (!photoDb?.getAllPhotos || !photoDb?.updateThumbnailPath) {
    return;
  }

  const rows = photoDb.getAllPhotos();

  for (const row of rows) {
    const rebasedThumbnailPath = rebaseManagedThumbnailPath(row?.thumbnail_path);

    if (!rebasedThumbnailPath || rebasedThumbnailPath === row.thumbnail_path) {
      continue;
    }

    try {
      await fs.access(rebasedThumbnailPath);
      photoDb.updateThumbnailPath(row.id, rebasedThumbnailPath);
    } catch {
      // Thumbnail cache is disposable. If a migrated file is missing, the row can
      // be refreshed later by regeneration or maintenance actions.
    }
  }
}

function getSquirrelUpdateExecutablePath() {
  return path.resolve(path.dirname(process.execPath), '..', 'Update.exe');
}

function isInternalUninstallSupportedRuntime() {
  return app.isPackaged && process.platform === 'win32';
}

async function closePhotoDatabaseForUninstall() {
  if (!photoDb?.close) {
    return;
  }

  photoDb.close();
  photoDb = null;
}

async function deleteAppDataForUninstall() {
  const targets = [
    getDatabaseDirectoryPath(),
    ...getManagedThumbnailRootPaths(),
  ].filter(Boolean);

  for (const targetPath of targets) {
    await fs.rm(targetPath, {
      recursive: true,
      force: true,
      maxRetries: 2,
      retryDelay: 120,
    });
  }
}

async function startInternalUninstall({ deleteData = false } = {}) {
  if (!isInternalUninstallSupportedRuntime()) {
    return {
      ok: false,
      message: 'インストーラー版の Windows アプリでのみ利用できます。',
    };
  }

  const updateExecutablePath = getSquirrelUpdateExecutablePath();

  try {
    await fs.access(updateExecutablePath);
  } catch {
    return {
      ok: false,
      message: 'アンインストーラーが見つかりませんでした。',
    };
  }

  try {
    if (deleteData) {
      await closePhotoDatabaseForUninstall();
      await deleteAppDataForUninstall();
    }

    const uninstallProcess = spawn(
      updateExecutablePath,
      ['--uninstall'],
      {
        detached: true,
        stdio: 'ignore',
        cwd: path.dirname(updateExecutablePath),
      }
    );
    uninstallProcess.unref();
    setTimeout(() => app.quit(), 120);

    return { ok: true };
  } catch (error) {
    return {
      ok: false,
      message: error?.message || 'アンインストールを開始できませんでした。',
    };
  }
}

// App update helpers stay in main so packaged builds can decide whether to
// download via Squirrel without exposing network logic to the renderer.
function isAutoUpdateSupportedRuntime() {
  return process.platform === 'win32' && app.isPackaged;
}

function isSquirrelFirstRunLaunch() {
  return process.argv.includes('--squirrel-firstrun');
}

function normalizeReleaseVersionTag(versionTag) {
  return typeof versionTag === 'string' ? versionTag.trim().replace(/^v/i, '') : '';
}

function getComparableVersionParts(version) {
  return normalizeReleaseVersionTag(version)
    .split('.')
    .map((part) => Number.parseInt(part, 10))
    .map((part) => (Number.isFinite(part) ? part : 0));
}

function compareComparableVersions(leftVersion, rightVersion) {
  const leftParts = getComparableVersionParts(leftVersion);
  const rightParts = getComparableVersionParts(rightVersion);
  const maxLength = Math.max(leftParts.length, rightParts.length, 3);

  for (let index = 0; index < maxLength; index += 1) {
    const leftPart = leftParts[index] || 0;
    const rightPart = rightParts[index] || 0;

    if (leftPart > rightPart) {
      return 1;
    }

    if (leftPart < rightPart) {
      return -1;
    }
  }

  return 0;
}

function sendAppUpdateStatusToRenderer(message) {
  if (
    !message ||
    !mainWindowRef ||
    mainWindowRef.isDestroyed() ||
    mainWindowRef.webContents.isDestroyed()
  ) {
    return;
  }

  mainWindowRef.webContents.send(APP_UPDATE_STATUS_CHANNEL, {
    message,
  });
}

function sendAppUpdateActionToRenderer(payload) {
  if (
    !payload ||
    typeof payload.kind !== 'string' ||
    !mainWindowRef ||
    mainWindowRef.isDestroyed() ||
    mainWindowRef.webContents.isDestroyed()
  ) {
    return;
  }

  mainWindowRef.webContents.send(APP_UPDATE_ACTION_CHANNEL, payload);
}

function buildAutoUpdateFeedUrl() {
  return `${APP_UPDATE_SERVICE_BASE_URL}/${APP_UPDATE_REPOSITORY_OWNER}/${APP_UPDATE_REPOSITORY_NAME}/${process.platform}-${process.arch}/${app.getVersion()}`;
}

async function fetchLatestGitHubReleaseInfo() {
  const response = await net.fetch(APP_UPDATE_RELEASE_API_URL, {
    headers: {
      accept: 'application/vnd.github+json',
      'user-agent': `${APP_UPDATE_REPOSITORY_NAME}/${app.getVersion()}`,
    },
  });

  if (response.status === 404 || response.status === 204) {
    return null;
  }

  if (!response.ok) {
    throw new Error(`GitHub Releases returned ${response.status}`);
  }

  const payload = await response.json();
  const version = normalizeReleaseVersionTag(
    payload?.tag_name || payload?.name || ''
  );

  if (!version) {
    return null;
  }

  return {
    version,
    title:
      typeof payload?.name === 'string' && payload.name.trim()
        ? payload.name.trim()
        : `v${version}`,
    htmlUrl:
      typeof payload?.html_url === 'string' ? payload.html_url.trim() : '',
  };
}

function setupAutoUpdater() {
  if (isAutoUpdaterConfigured || !isAutoUpdateSupportedRuntime()) {
    return;
  }

  isAutoUpdaterConfigured = true;
  autoUpdater.on('update-available', () => {
    sendAppUpdateStatusToRenderer('アップデートをダウンロードしています...');
  });

  autoUpdater.on('update-not-available', () => {
    isAutoUpdateDownloadRunning = false;
  });

  autoUpdater.on('error', (error) => {
    isAutoUpdateDownloadRunning = false;
    const message =
      error instanceof Error && error.message
        ? error.message
        : '不明なエラー';
    sendAppUpdateStatusToRenderer(`アップデートに失敗しました: ${message}`);
  });

  autoUpdater.on('update-downloaded', () => {
    isAutoUpdateDownloadRunning = false;
    sendAppUpdateStatusToRenderer('アップデートの準備ができました');
    sendAppUpdateActionToRenderer({
      kind: 'downloaded',
      version: latestAvailableAppUpdateRelease?.version || '',
    });
  });
  // Legacy native-dialog flow is intentionally left below as quarantine only.
  return;

  autoUpdater.on('update-available', () => {
    sendAppUpdateStatusToRenderer('アップデートをダウンロードしています...');
  });

  autoUpdater.on('update-not-available', () => {
    isAutoUpdateDownloadRunning = false;
  });

  autoUpdater.on('error', (error) => {
    isAutoUpdateDownloadRunning = false;
    const message =
      error instanceof Error && error.message
        ? error.message
        : '不明なエラー';
    sendAppUpdateStatusToRenderer(
      `アップデートに失敗しました: ${message}`
    );
  });

  autoUpdater.on('update-downloaded', async () => {
    isAutoUpdateDownloadRunning = false;
    sendAppUpdateStatusToRenderer('アップデートの準備ができました');

    const targetWindow =
      mainWindowRef && !mainWindowRef.isDestroyed() ? mainWindowRef : null;
    const promptResult = await dialog.showMessageBox(targetWindow, {
      type: 'info',
      buttons: ['再起動して更新', 'あとで'],
      defaultId: 0,
      cancelId: 1,
      title: 'アップデートの準備ができました',
      message: 'ダウンロードが完了しました。',
      detail: '再起動して最新バージョンを適用しますか？',
      noLink: true,
    });

    if (promptResult.response === 0) {
      autoUpdater.quitAndInstall();
    }
  });
}

async function promptForAvailableUpdate(releaseInfo) {
  if (!releaseInfo || isAutoUpdatePromptOpen) {
    return false;
  }

  isAutoUpdatePromptOpen = true;

  try {
    const targetWindow =
      mainWindowRef && !mainWindowRef.isDestroyed() ? mainWindowRef : null;
    const promptResult = await dialog.showMessageBox(targetWindow, {
      type: 'info',
      buttons: ['今すぐ更新', 'あとで'],
      defaultId: 0,
      cancelId: 1,
      title: 'アップデートがあります',
      message: `新しいバージョン ${releaseInfo.version} が利用できます。`,
      detail: 'ダウンロードして適用しますか？',
      noLink: true,
    });

    return promptResult.response === 0;
  } finally {
    isAutoUpdatePromptOpen = false;
  }
}

async function startAppUpdateDownload(releaseInfo) {
  const activeReleaseInfo = releaseInfo || latestAvailableAppUpdateRelease;

  if (
    !activeReleaseInfo ||
    isAutoUpdateDownloadRunning ||
    !isAutoUpdateSupportedRuntime()
  ) {
    return;
  }

  setupAutoUpdater();
  isAutoUpdateDownloadRunning = true;
  latestAvailableAppUpdateRelease = activeReleaseInfo;
  sendAppUpdateStatusToRenderer(
    `アップデート ${activeReleaseInfo.version} をダウンロードしています...`
  );
  autoUpdater.setFeedURL({
    url: buildAutoUpdateFeedUrl(),
  });
  autoUpdater.checkForUpdates();
  // Legacy native-dialog flow is intentionally left below as quarantine only.
  return;

  if (!releaseInfo || isAutoUpdateDownloadRunning || !isAutoUpdateSupportedRuntime()) {
    return;
  }

  setupAutoUpdater();
  isAutoUpdateDownloadRunning = true;
  sendAppUpdateStatusToRenderer(
    `アップデート ${releaseInfo.version} をダウンロードしています...`
  );
  autoUpdater.setFeedURL({
    url: buildAutoUpdateFeedUrl(),
  });
  autoUpdater.checkForUpdates();
}

async function checkForAppUpdatesOnLaunch() {
  if (
    !isAutoUpdateSupportedRuntime() ||
    isSquirrelFirstRunLaunch() ||
    isAutoUpdateCheckRunning ||
    isAutoUpdateDownloadRunning
  ) {
    return;
  }

  isAutoUpdateCheckRunning = true;

  try {
    const latestRelease = await fetchLatestGitHubReleaseInfo();

    if (!latestRelease) {
      return;
    }

    if (compareComparableVersions(latestRelease.version, app.getVersion()) <= 0) {
      return;
    }

    latestAvailableAppUpdateRelease = latestRelease;
    sendAppUpdateStatusToRenderer(
      `新しいバージョン ${latestRelease.version} が見つかりました`
    );
    sendAppUpdateActionToRenderer({
      kind: 'available',
      version: latestRelease.version,
    });
    // Legacy native-dialog prompt flow is intentionally left below as quarantine only.
    return;
  } catch (error) {
    const message =
      error instanceof Error && error.message ? error.message : '不明なエラー';
    sendAppUpdateStatusToRenderer(`アップデート確認に失敗しました: ${message}`);
    return;
  } finally {
    isAutoUpdateCheckRunning = false;
  }

  try {
    const latestRelease = await fetchLatestGitHubReleaseInfo();

    if (!latestRelease) {
      return;
    }

    if (compareComparableVersions(latestRelease.version, app.getVersion()) <= 0) {
      return;
    }

    sendAppUpdateStatusToRenderer(
      `新しいバージョン ${latestRelease.version} が見つかりました`
    );
    const shouldDownload = await promptForAvailableUpdate(latestRelease);

    if (!shouldDownload) {
      sendAppUpdateStatusToRenderer('アップデートは保留しました');
      return;
    }

    await startAppUpdateDownload(latestRelease);
  } catch (error) {
    const message =
      error instanceof Error && error.message ? error.message : '不明なエラー';
    sendAppUpdateStatusToRenderer(
      `アップデート確認に失敗しました: ${message}`
    );
  } finally {
    isAutoUpdateCheckRunning = false;
  }
}

function scheduleStartupAutoUpdateCheck(mainWindow) {
  if (!mainWindow || !isAutoUpdateSupportedRuntime()) {
    return;
  }

  mainWindow.webContents.once('did-finish-load', () => {
    setTimeout(() => {
      void checkForAppUpdatesOnLaunch();
    }, APP_UPDATE_CHECK_DELAY_MS);
  });
}

function isSupportedImageFile(filePath) {
  return SUPPORTED_IMAGE_EXTENSIONS.has(path.extname(filePath).toLowerCase());
}

function countJapaneseChars(value) {
  return (value.match(/[\u3040-\u30ff\u3400-\u9fff]/g) || []).length;
}

function countReplacementChars(value) {
  return (value.match(/�/g) || []).length;
}

function countMojibakeMarkers(value) {
  return (
    value.match(/[ÃÂâåãæçèéêëìíîïðñòóôõöøùúûüýþÿ竄繝縺繧鬘蜿榊]/g) || []
  ).length;
}

function countControlChars(value) {
  return (value.match(/[\u0000-\u001f\u007f]/g) || []).length;
}

function countQuestionMarks(value) {
  return (value.match(/\?/g) || []).length;
}

function textScore(value) {
  if (typeof value !== 'string' || value.length === 0) {
    return -9999;
  }

  return (
    countJapaneseChars(value) * 4 -
    countMojibakeMarkers(value) * 3 -
    countReplacementChars(value) * 8 -
    countControlChars(value) * 10 -
    countQuestionMarks(value)
  );
}

function tryLatin1ToUtf8(value) {
  try {
    return Buffer.from(value, 'latin1').toString('utf8');
  } catch {
    return value;
  }
}

function tryShiftJisToUtf8(value) {
  try {
    return iconv.encode(value, 'shift_jis').toString('utf8');
  } catch {
    return value;
  }
}

function repairUtf8Mojibake(value) {
  if (typeof value !== 'string' || value.length === 0) {
    return value;
  }

  const candidates = new Set([
    value,
    tryLatin1ToUtf8(value),
    tryShiftJisToUtf8(value),
    tryLatin1ToUtf8(tryShiftJisToUtf8(value)),
    tryShiftJisToUtf8(tryLatin1ToUtf8(value)),
  ]);

  let best = value;
  let bestScore = textScore(value);

  for (const candidate of candidates) {
    const score = textScore(candidate);
    if (score > bestScore) {
      best = candidate;
      bestScore = score;
    }
  }

  return best;
}

function isExifReaderTagEnvelope(rawValue) {
  return (
    rawValue &&
    typeof rawValue === 'object' &&
    !Array.isArray(rawValue) &&
    (Object.prototype.hasOwnProperty.call(rawValue, 'value') ||
      Object.prototype.hasOwnProperty.call(rawValue, 'description') ||
      Object.prototype.hasOwnProperty.call(rawValue, 'attributes'))
  );
}

function isScalarTagTextValue(rawValue) {
  return (
    typeof rawValue === 'string' ||
    typeof rawValue === 'number' ||
    typeof rawValue === 'boolean' ||
    typeof rawValue === 'bigint'
  );
}

function collectTagTextCandidates(rawValue, visited = new Set()) {
  if (rawValue == null) {
    return [];
  }

  if (isScalarTagTextValue(rawValue)) {
    return [String(rawValue)];
  }

  if (typeof rawValue !== 'object') {
    return [String(rawValue)];
  }

  if (visited.has(rawValue)) {
    return [];
  }

  visited.add(rawValue);

  if (isExifReaderTagEnvelope(rawValue)) {
    const valueCandidates = collectTagTextCandidates(rawValue.value, visited);

    if (valueCandidates.length > 0) {
      if (isScalarTagTextValue(rawValue.value)) {
        return [
          ...valueCandidates,
          ...collectTagTextCandidates(rawValue.description, visited),
        ];
      }

      return valueCandidates;
    }

    return collectTagTextCandidates(rawValue.description, visited);
  }

  if (Array.isArray(rawValue)) {
    return rawValue.flatMap((item) => collectTagTextCandidates(item, visited));
  }

  const nested = Object.values(rawValue).flatMap((item) =>
    collectTagTextCandidates(item, visited)
  );

  if (nested.length > 0) {
    return nested;
  }

  try {
    return [JSON.stringify(rawValue)];
  } catch {
    return [];
  }
}

function sanitizeExtractedText(value) {
  if (typeof value !== 'string') {
    return null;
  }

  const sanitized = value
    .replace(/\u0000/g, ' ')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();

  if (/^\[object\s+[^\]]+\]$/i.test(sanitized)) {
    return null;
  }

  return sanitized.length > 0 ? sanitized : null;
}

function isEmptyPrintNotePlaceholder(value) {
  if (typeof value !== 'string') {
    return false;
  }

  const trimmed = value.trim();

  if (trimmed === '{}' || trimmed === '[]') {
    return true;
  }

  try {
    const parsed = JSON.parse(trimmed);

    if (Array.isArray(parsed)) {
      return parsed.length === 0;
    }

    if (parsed && typeof parsed === 'object') {
      return Object.values(parsed).every((entry) => {
        if (entry == null) {
          return true;
        }

        if (typeof entry === 'string') {
          return entry.trim().length === 0;
        }

        if (Array.isArray(entry)) {
          return entry.length === 0;
        }

        return false;
      });
    }
  } catch {
    // Not JSON-like, so it can still be a real print note.
  }

  return false;
}

function pickBestTextCandidate(rawValues) {
  const candidates = new Set();

  for (const rawValue of rawValues) {
    for (const stringValue of collectTagTextCandidates(rawValue)) {
      if (!stringValue) {
        continue;
      }

      candidates.add(stringValue);
      candidates.add(repairUtf8Mojibake(stringValue));

      const strippedNulls = stringValue.replace(/\u0000+/g, '');
      candidates.add(strippedNulls);
      candidates.add(repairUtf8Mojibake(strippedNulls));

      const strippedControls = stringValue.replace(/[\u0000-\u001f\u007f]+/g, ' ');
      candidates.add(strippedControls);
      candidates.add(repairUtf8Mojibake(strippedControls));

      const removedControls = stringValue.replace(/[\u0000-\u001f\u007f]+/g, '');
      candidates.add(removedControls);
      candidates.add(repairUtf8Mojibake(removedControls));
    }
  }

  let best = null;
  let bestScore = -9999;

  for (const candidate of candidates) {
    const sanitized = sanitizeExtractedText(candidate);

    if (!sanitized) {
      continue;
    }

    const score = textScore(sanitized);

    if (score > bestScore) {
      best = sanitized;
      bestScore = score;
    }
  }

  return best;
}

function normalizeWorldId(value) {
  if (typeof value !== 'string') {
    return null;
  }

  const trimmed = value.trim();
  return /^wrld_[0-9a-f-]+$/i.test(trimmed) ? trimmed : null;
}

function isNumericTagValueArray(value) {
  return (
    Array.isArray(value) &&
    value.length > 0 &&
    value.every((item) => typeof item === 'number' && Number.isFinite(item))
  );
}

function normalizeTagValue(tag) {
  if (!tag) return null;

  const valueCandidate = isNumericTagValueArray(tag.value)
    ? null
    : pickBestTextCandidate([tag.value]);

  return valueCandidate || pickBestTextCandidate([tag.description]);
}

function normalizePhotoPrintNoteText(value) {
  const sanitized = sanitizeExtractedText(repairUtf8Mojibake(value));

  if (!sanitized || isEmptyPrintNotePlaceholder(sanitized)) {
    return null;
  }

  return sanitized
    .replace(/\u2044/g, '/')
    .replace(/[\u201c\u201d\uFF02]/g, '"')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/\uFFFD/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim()
    .normalize('NFC');
}

function shouldRepairStoredPrintNote(value) {
  if (typeof value !== 'string') {
    return false;
  }

  const trimmed = value.trim();

  if (!trimmed) {
    return false;
  }

  if (/^\[object\s+[^\]]+\]$/i.test(trimmed)) {
    return true;
  }

  if (isEmptyPrintNotePlaceholder(trimmed)) {
    return true;
  }

  const commaSeparatedSegments = trimmed
    .split(/\s*,\s*/)
    .map((segment) => segment.trim())
    .filter(Boolean);

  if (
    commaSeparatedSegments.some((segment) =>
      /^(?:x-default|[a-z]{2,3}(?:-[a-z0-9]{2,8})+)$/i.test(segment)
    )
  ) {
    return true;
  }

  return /^\{.*(?:x-default|xml:lang|lang=).*\}$/i.test(trimmed);
}

// Only treat the exact "{}" placeholder as missing so unusual but real names still render.
function hasMeaningfulWorldNameContent(value) {
  if (typeof value !== 'string') {
    return false;
  }

  const trimmed = value.trim();

  if (!trimmed) {
    return false;
  }

  return trimmed !== '{}';
}

function normalizeDisplayWorldName(value) {
  const sanitized = sanitizeExtractedText(value);

  if (!sanitized) {
    return null;
  }

  const normalized = sanitized
    .replace(/[⁄∕／]/g, '/')
    .replace(/ǃ/g, '!')
    .replace(/‚/g, ',')
    .replace(/․/g, '.')
    .replace(/［/g, '[')
    .replace(/］/g, ']')
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .replace(/＂/g, '"')
    .replace(/\uFFFD/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim()
    .normalize('NFC');

  return hasMeaningfulWorldNameContent(normalized) ? normalized : null;
}

function normalizeOfficialWorldText(value) {
  const sanitized = sanitizeExtractedText(repairUtf8Mojibake(value));

  if (!sanitized) {
    return null;
  }

  return sanitized
    .replace(/\u2044/g, '/')
    .replace(/\u201a/g, ',')
    .replace(/\u2024/g, '.')
    .replace(/[\u201c\u201d\uFF02]/g, '"')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/\uFFFD/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim()
    .normalize('NFC');
}

function normalizeOfficialWorldTags(tags) {
  if (!Array.isArray(tags)) {
    return [];
  }

  return Array.from(
    new Set(
      tags
        .map((tag) => normalizeOfficialWorldText(tag))
        .filter(Boolean)
    )
  );
}

function pickDefaultPhotoLabelColor(normalizedName = '') {
  const seed = typeof normalizedName === 'string' ? normalizedName : '';
  let hash = 0;

  for (let index = 0; index < seed.length; index += 1) {
    hash = (hash * 31 + seed.charCodeAt(index)) >>> 0;
  }

  return DEFAULT_PHOTO_LABEL_COLORS[
    hash % DEFAULT_PHOTO_LABEL_COLORS.length
  ];
}

function normalizePhotoLabelColor(colorHex, normalizedName = '') {
  if (typeof colorHex === 'string') {
    const trimmed = colorHex.trim();
    const match = trimmed.match(/^#?([0-9a-fA-F]{6})$/);

    if (match) {
      return `#${match[1].toUpperCase()}`;
    }
  }

  return pickDefaultPhotoLabelColor(normalizedName);
}

function normalizePhotoLabelPayload(labels) {
  return Array.from(
    new Map(
      (Array.isArray(labels) ? labels : [])
        .map((label) => {
          const rawName =
            typeof label === 'string'
              ? label
              : typeof label?.name === 'string'
                ? label.name
                : '';

          const name = rawName
            .normalize('NFC')
            .replace(/\s+/g, ' ')
            .trim();

          if (!name) {
            return null;
          }

          const normalizedName = name.toLowerCase();

          return [
            normalizedName,
            {
              name,
              normalizedName,
              colorHex: normalizePhotoLabelColor(
                typeof label === 'object'
                  ? label.colorHex || label.color_hex
                  : null,
                normalizedName
              ),
            },
          ];
        })
        .filter(Boolean)
    ).values()
  );
}

function buildWorldUrlFromId(worldId) {
  return normalizeWorldId(worldId)
    ? `https://vrchat.com/home/world/${worldId}`
    : null;
}

function parseWorldIdFromUrl(worldUrl) {
  if (typeof worldUrl !== 'string' || worldUrl.trim().length === 0) {
    return null;
  }

  const matchedWorldId = worldUrl.match(/\/world\/(wrld_[0-9a-f-]+)/i);
  return normalizeWorldId(matchedWorldId?.[1] || null);
}

function normalizeManualWorldUrl(worldUrl, fallbackWorldUrl = null) {
  const rawValue =
    typeof worldUrl === 'string' ? worldUrl.trim() : '';

  if (rawValue.length === 0) {
    const fallbackWorldId = parseWorldIdFromUrl(fallbackWorldUrl);
    return {
      worldId: fallbackWorldId,
      worldUrl:
        buildWorldUrlFromId(fallbackWorldId) ||
        (typeof fallbackWorldUrl === 'string' && fallbackWorldUrl.trim().length > 0
          ? fallbackWorldUrl.trim()
          : null),
    };
  }

  const normalizedWorldId =
    normalizeWorldId(rawValue) || parseWorldIdFromUrl(rawValue);

  if (!normalizedWorldId) {
    throw new Error('VRChatのWorld URLまたはWorld IDを入力してください');
  }

  return {
    worldId: normalizedWorldId,
    worldUrl: buildWorldUrlFromId(normalizedWorldId),
  };
}

function decodeHtmlEntities(value) {
  if (typeof value !== 'string' || value.length === 0) {
    return value;
  }

  return value
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
}

function extractMetaTagContent(html, keys) {
  if (typeof html !== 'string' || html.length === 0) {
    return null;
  }

  const normalizedKeys = Array.isArray(keys) ? keys : [keys];

  for (const key of normalizedKeys) {
    if (typeof key !== 'string' || key.length === 0) {
      continue;
    }

    const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const patterns = [
      new RegExp(
        `<meta[^>]+(?:name|property)=["']${escapedKey}["'][^>]+content=["']([^"']*)["'][^>]*>`,
        'i'
      ),
      new RegExp(
        `<meta[^>]+content=["']([^"']*)["'][^>]+(?:name|property)=["']${escapedKey}["'][^>]*>`,
        'i'
      ),
    ];

    for (const pattern of patterns) {
      const matched = html.match(pattern);

      if (matched?.[1]) {
        return decodeHtmlEntities(matched[1]);
      }
    }
  }

  return null;
}

function pad2(value) {
  return String(value).padStart(2, '0');
}

function formatDateTime(date) {
  return `${date.getFullYear()}/${pad2(date.getMonth() + 1)}/${pad2(
    date.getDate()
  )} ${pad2(date.getHours())}:${pad2(date.getMinutes())}:${pad2(
    date.getSeconds()
  )}`;
}

function formatGroupDate(date) {
  return `${date.getFullYear()}/${pad2(date.getMonth() + 1)}/${pad2(
    date.getDate()
  )}`;
}

function parseExifDateString(value) {
  if (!value || typeof value !== 'string') return null;

  const normalized = value
    .trim()
    .replace(/^(\d{4}):(\d{2}):(\d{2})/, '$1-$2-$3')
    .replace(/\.\d+$/, '');

  const parsed = new Date(normalized);

  if (Number.isNaN(parsed.getTime())) {
    return null;
  }

  return parsed;
}

function parseVrchatFilenameDate(filePath) {
  const fileName = path.basename(filePath);

  const match = fileName.match(
    /VRChat_(\d{4})-(\d{2})-(\d{2})_(\d{2})-(\d{2})-(\d{2})(?:\.\d+)?/i
  );

  if (!match) {
    return null;
  }

  const [, year, month, day, hour, minute, second] = match;

  return new Date(
    Number(year),
    Number(month) - 1,
    Number(day),
    Number(hour),
    Number(minute),
    Number(second)
  );
}

function parseImageDimensionsFromFilename(filePath) {
  const fileName = path.basename(filePath);
  const patterns = [
    /(?:^|[_\-. \[(])(\d{2,5})[xX](\d{2,5})(?:[_\-. \])]|$)/,
    /(?:width|w)(\d{2,5})[^0-9]{0,4}(?:height|h)(\d{2,5})/i,
  ];

  for (const pattern of patterns) {
    const match = fileName.match(pattern);

    if (!match) {
      continue;
    }

    const imageWidth = Number(match[1]);
    const imageHeight = Number(match[2]);

    if (
      !Number.isFinite(imageWidth) ||
      !Number.isFinite(imageHeight) ||
      imageWidth <= 0 ||
      imageHeight <= 0
    ) {
      continue;
    }

    return {
      imageWidth,
      imageHeight,
      resolutionTier: getResolutionTier(imageWidth, imageHeight),
      orientationTier: getOrientationTier(imageWidth, imageHeight),
    };
  }

  return null;
}

function extractTakenAtFromTags(tags) {
  const candidateKeys = [
    'DateTimeOriginal',
    'CreateDate',
    'DateCreated',
    'ModifyDate',
    'File Modification Date/Time',
  ];

  for (const key of candidateKeys) {
    const tag = tags[key];
    if (!tag) continue;

    const value = normalizeTagValue(tag);
    const parsed = parseExifDateString(value);

    if (parsed) {
      return parsed;
    }
  }

  return null;
}

function pickPreferredTag(flatTags, preferredKeys, matcher) {
  for (const key of preferredKeys) {
    if (flatTags[key]) {
      return { key, value: flatTags[key] };
    }
  }

  const fallbackKey = Object.keys(flatTags).find((key) =>
    matcher.test(key.toLowerCase())
  );

  if (!fallbackKey) {
    return null;
  }

  return {
    key: fallbackKey,
    value: flatTags[fallbackKey],
  };
}

function decodeXmlEntities(value) {
  if (typeof value !== 'string') {
    return value;
  }

  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) =>
      String.fromCodePoint(parseInt(hex, 16))
    )
    .replace(/&#([0-9]+);/g, (_, num) =>
      String.fromCodePoint(parseInt(num, 10))
    );
}

function extractXmpAttributeFromText(text, attrName) {
  if (typeof text !== 'string' || text.length === 0) {
    return null;
  }

  const patterns = [
    new RegExp(`(?:vrc:)?${attrName}="([^"]*)"`, 'i'),
    new RegExp(`(?:vrc:)?${attrName}='([^']*)'`, 'i'),
    new RegExp(
      `<(?:[a-z0-9_-]+:)?${attrName}>([\\s\\S]*?)<\\/(?:[a-z0-9_-]+:)?${attrName}>`,
      'i'
    ),
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match && match[1] != null) {
      return decodeXmlEntities(match[1]);
    }
  }

  return null;
}

function extractRelevantXmpTextFromBuffer(fileBuffer) {
  const xmpStartMarker = Buffer.from('<x:xmpmeta', 'utf8');
  const xmpEndMarker = Buffer.from('</x:xmpmeta>', 'utf8');
  const worldIdMarker = Buffer.from('WorldID', 'utf8');
  const worldNameMarker = Buffer.from('WorldDisplayName', 'utf8');

  const xmpStartIndex = fileBuffer.indexOf(xmpStartMarker);

  if (xmpStartIndex !== -1) {
    const xmpEndIndex = fileBuffer.indexOf(xmpEndMarker, xmpStartIndex);

    if (xmpEndIndex !== -1) {
      return fileBuffer.toString(
        'utf8',
        xmpStartIndex,
        xmpEndIndex + xmpEndMarker.length
      );
    }
  }

  const anchorIndices = [
    fileBuffer.indexOf(worldIdMarker),
    fileBuffer.indexOf(worldNameMarker),
  ].filter((index) => index !== -1);

  if (anchorIndices.length === 0) {
    return null;
  }

  const anchorIndex = Math.min(...anchorIndices);
  const sliceStart = Math.max(0, anchorIndex - 4096);
  const sliceEnd = Math.min(fileBuffer.length, anchorIndex + 4096);

  return fileBuffer.toString('utf8', sliceStart, sliceEnd);
}

function extractWorldInfoFromRawBuffer(fileBuffer) {
  const xmpText = extractRelevantXmpTextFromBuffer(fileBuffer);

  if (!xmpText) {
    return {
      worldId: null,
      worldName: null,
    };
  }

  const worldId = extractXmpAttributeFromText(xmpText, 'WorldID');
  const worldName = extractXmpAttributeFromText(xmpText, 'WorldDisplayName');

  return {
    worldId,
    worldName,
  };
}

function extractWorldInfo(tags, fileBuffer, rawBufferWorldInfo = null) {
  const flatTags = {};

  for (const [key, tag] of Object.entries(tags || {})) {
    flatTags[key] = normalizeTagValue(tag);
  }

  const worldIdEntry = pickPreferredTag(
    flatTags,
    ['XMP-vrc:WorldID', 'vrc:WorldID', 'World ID', 'WorldID'],
    /worldid|world id/
  );

  const worldNameEntry = pickPreferredTag(
    flatTags,
    [
      'XMP-vrc:WorldDisplayName',
      'vrc:WorldDisplayName',
      'World Display Name',
      'WorldDisplayName',
    ],
    /worlddisplayname|world display name/
  );

  const resolvedRawWorldInfo =
    rawBufferWorldInfo || extractWorldInfoFromRawBuffer(fileBuffer);

  const worldId = normalizeWorldId(
    resolvedRawWorldInfo.worldId ||
      (worldIdEntry ? worldIdEntry.value : null) ||
      null
  );

  const worldName = normalizeDisplayWorldName(
    pickBestTextCandidate([
      resolvedRawWorldInfo.worldName,
      worldNameEntry ? worldNameEntry.value : null,
    ])
  );

  const worldUrl = worldId
    ? `https://vrchat.com/home/world/${worldId}`
    : null;

  return {
    worldId,
    worldName,
    worldUrl,
  };
}

function extractPhotoPrintNote(tags) {
  const flatTags = {};
  const printNoteCandidates = [];
  const preferredKeys = [
    'XMP-dc:Title',
    'dc:Title',
    'Title',
    'XPTitle',
    'Windows XP Title',
    'IPTC:ObjectName',
    'ObjectName',
    'Object Name',
    'XMP-photoshop:Headline',
    'Headline',
  ];
  const fallbackMatcher =
    /(?:^|[: _-])(title|xptitle|objectname|object name|headline)(?:$|[: _-])/;

  for (const [key, tag] of Object.entries(tags || {})) {
    const normalizedValue = normalizeTagValue(tag);
    flatTags[key] = normalizedValue;

    if (
      normalizedValue &&
      (preferredKeys.includes(key) || fallbackMatcher.test(key.toLowerCase()))
    ) {
      printNoteCandidates.push(normalizedValue);
    }
  }

  const printNoteEntry = pickPreferredTag(flatTags, preferredKeys, fallbackMatcher);
  const preferredPrintNoteText = pickBestTextCandidate(
    printNoteCandidates.length > 0
      ? printNoteCandidates
      : [printNoteEntry?.value || null]
  );

  return normalizePhotoPrintNoteText(preferredPrintNoteText);
}

function loadExifTagsSafely(fileBuffer) {
  try {
    return ExifReader.load(fileBuffer);
  } catch {
    return {};
  }
}


async function ensureDir(dirPath) {
  await fs.mkdir(dirPath, { recursive: true });
}

function calculateThumbnailFitSize(width, height) {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return null;
  }

  const longEdge = Math.max(width, height);
  const scale = Math.min(1, THUMBNAIL_SIZE / longEdge);

  return {
    width: Math.max(1, Math.round(width * scale)),
    height: Math.max(1, Math.round(height * scale)),
  };
}

function getManagedThumbnailPath(fileHash, extension = THUMBNAIL_EXTENSION) {
  return path.join(thumbnailDirPath, `${fileHash}${extension}`);
}

async function createThumbnailFromNativeImage(image, fileHash, options = {}) {
  const { force = false } = options;
  const thumbnailPath = getManagedThumbnailPath(fileHash);

  if (!force) {
    try {
      await fs.access(thumbnailPath);
      return thumbnailPath;
    } catch {
      // no existing thumbnail
    }
  }

  if (!image || image.isEmpty()) {
    return null;
  }

  const sourceSize = image.getSize();
  const fitSize = calculateThumbnailFitSize(
    Number(sourceSize.width) || 0,
    Number(sourceSize.height) || 0
  );

  if (!fitSize) {
    return null;
  }

  try {
    const resizedImage =
      fitSize.width === sourceSize.width && fitSize.height === sourceSize.height
        ? image
        : image.resize({
            width: fitSize.width,
            height: fitSize.height,
            quality: 'good',
          });

    const jpegBuffer = resizedImage.toJPEG(THUMBNAIL_JPEG_QUALITY);

    if (!jpegBuffer || jpegBuffer.length === 0) {
      return null;
    }

    await fs.writeFile(thumbnailPath, jpegBuffer);
    return thumbnailPath;
  } catch {
    return null;
  }
}

async function createThumbnail(filePath, fileHash, options = {}) {
  const { force = false, sourceImage = null } = options;

  if (sourceImage && !sourceImage.isEmpty()) {
    const thumbnailPath = await createThumbnailFromNativeImage(
      sourceImage,
      fileHash,
      options
    );

    if (thumbnailPath) {
      return thumbnailPath;
    }
  }

  const thumbnailPath = getManagedThumbnailPath(fileHash);

  if (!force) {
    try {
      await fs.access(thumbnailPath);
      return thumbnailPath;
    } catch {
      // なければ生成
    }
  }

  try {
    const image = await nativeImage.createThumbnailFromPath(filePath, {
      width: THUMBNAIL_SIZE,
      height: THUMBNAIL_SIZE,
    });

    const jpegBuffer = image.toJPEG(THUMBNAIL_JPEG_QUALITY);

    if (!jpegBuffer || jpegBuffer.length === 0) {
      return null;
    }

    await fs.writeFile(thumbnailPath, jpegBuffer);
    return thumbnailPath;
  } catch {
    return null;
  }
}

async function ensureManagedThumbnailForRow(row) {
  if (
    !row ||
    !Number.isInteger(row.id) ||
    typeof row.file_path !== 'string' ||
    row.file_path.trim().length === 0 ||
    typeof row.file_hash !== 'string' ||
    row.file_hash.trim().length === 0
  ) {
    return row;
  }

  const hasStoredThumbnail =
    typeof row.thumbnail_path === 'string' && row.thumbnail_path.trim().length > 0;

  if (hasStoredThumbnail && (await pathExists(row.thumbnail_path))) {
    return row;
  }

  const nextThumbnailPath = await createThumbnail(row.file_path, row.file_hash, {
    force: true,
  });

  if (!nextThumbnailPath) {
    return row;
  }

  return photoDb.updateThumbnailPath(row.id, nextThumbnailPath) || {
    ...row,
    thumbnail_path: nextThumbnailPath,
  };
}

function getResolutionTier(width, height) {
  if (!Number.isFinite(width) || !Number.isFinite(height)) {
    return null;
  }

  const longEdge = Math.max(width, height);
  const shortEdge = Math.min(width, height);

  if (longEdge >= 7680 && shortEdge >= 4320) {
    return '8K';
  }

  if (longEdge >= 3840 && shortEdge >= 2160) {
    return '4K';
  }

  if (longEdge >= 2560 && shortEdge >= 1440) {
    return 'QHD';
  }

  if (longEdge >= 1920 && shortEdge >= 1080) {
    return 'FHD';
  }

  if (longEdge >= 1280 && shortEdge >= 720) {
    return 'HD';
  }

  return null;
}

function getOrientationTier(width, height) {
  if (!Number.isFinite(width) || !Number.isFinite(height)) {
    return null;
  }

  if (width === height) {
    return 'square';
  }

  return width > height ? 'landscape' : 'portrait';
}

function isKnownPhotoOrientationTier(value) {
  return typeof value === 'string' && PHOTO_ORIENTATION_TIERS.has(value);
}

function getStoredPhotoOrientationTier(row) {
  if (!row) {
    return null;
  }

  if (isKnownPhotoOrientationTier(row.orientation_tier)) {
    return row.orientation_tier;
  }

  return getOrientationTier(row.image_width, row.image_height);
}

function isLikelyVrchatPrintPhotoDimensions(width, height) {
  const imageWidth = Number(width);
  const imageHeight = Number(height);

  if (
    !Number.isFinite(imageWidth) ||
    !Number.isFinite(imageHeight) ||
    imageWidth <= 0 ||
    imageHeight <= 0
  ) {
    return false;
  }

  const longEdge = Math.max(imageWidth, imageHeight);
  const shortEdge = Math.min(imageWidth, imageHeight);

  if (longEdge < VRCHAT_PRINT_PHOTO_MIN_LONG_EDGE || shortEdge <= 0) {
    return false;
  }

  const aspectRatio = longEdge / shortEdge;
  return (
    Math.abs(aspectRatio - VRCHAT_PRINT_PHOTO_ASPECT_RATIO) <=
    VRCHAT_PRINT_PHOTO_ASPECT_TOLERANCE
  );
}

function isLikelyVrchatPrintPhoto(row) {
  return isLikelyVrchatPrintPhotoDimensions(row?.image_width, row?.image_height);
}

function extractImageMetadataFromNativeImage(image) {
  if (!image || image.isEmpty()) {
    return {
      imageWidth: null,
      imageHeight: null,
      resolutionTier: null,
      orientationTier: null,
    };
  }

  const size = image.getSize();
  const imageWidth = Number(size.width) || 0;
  const imageHeight = Number(size.height) || 0;

  if (imageWidth <= 0 || imageHeight <= 0) {
    return {
      imageWidth: null,
      imageHeight: null,
      resolutionTier: null,
      orientationTier: null,
    };
  }

  return {
    imageWidth,
    imageHeight,
    resolutionTier: getResolutionTier(imageWidth, imageHeight),
    orientationTier: getOrientationTier(imageWidth, imageHeight),
  };
}

function extractImageDetailsFromBuffer(fileBuffer) {
  try {
    const image = nativeImage.createFromBuffer(fileBuffer);
    const metadata = extractImageMetadataFromNativeImage(image);

    return {
      image,
      ...metadata,
    };
  } catch {
    return {
      image: null,
      imageWidth: null,
      imageHeight: null,
      resolutionTier: null,
      orientationTier: null,
    };
  }
}

function extractImageMetadataFromBuffer(fileBuffer) {
  const { image: _image, ...metadata } = extractImageDetailsFromBuffer(fileBuffer);
  return metadata;
}

function extractImageMetadataFromPath(filePath) {
  try {
    const image = nativeImage.createFromPath(filePath);
    return extractImageMetadataFromNativeImage(image);
  } catch {
    return {
      imageWidth: null,
      imageHeight: null,
      resolutionTier: null,
      orientationTier: null,
    };
  }
}

function ensurePhotoResolutionMetadata(row) {
  const hasWidth = Number.isFinite(row.image_width) && row.image_width > 0;
  const hasHeight = Number.isFinite(row.image_height) && row.image_height > 0;

  if (hasWidth && hasHeight) {
    return row;
  }

  const imageMetadata = extractImageMetadataFromPath(row.file_path);

  if (!imageMetadata.imageWidth || !imageMetadata.imageHeight) {
    return row;
  }

  photoDb.updateImageMetadata(row.id, imageMetadata);

  return {
    ...row,
    image_width: imageMetadata.imageWidth,
    image_height: imageMetadata.imageHeight,
    resolution_tier: imageMetadata.resolutionTier,
    orientation_tier: imageMetadata.orientationTier,
  };
}

function setBoundedCacheEntry(
  cache,
  key,
  value,
  limit = PHOTO_PREPARATION_CACHE_LIMIT
) {
  if (!(cache instanceof Map) || !key) {
    return;
  }

  if (!cache.has(key) && cache.size >= limit) {
    const oldestKey = cache.keys().next().value;
    cache.delete(oldestKey);
  }

  cache.set(key, value);
}

function clearPhotoPreparationCaches() {
  photoOrientationPreparationCache.clear();
  photoPrintNotePreparationCache.clear();
}

function buildPhotoOrientationPreparationCacheKey(row) {
  if (!row || !Number.isInteger(row.id)) {
    return null;
  }

  return [
    row.id,
    row.file_path || '',
    row.thumbnail_path || '',
    row.image_width || '',
    row.image_height || '',
    row.resolution_tier || '',
    row.orientation_tier || '',
  ].join('\u001f');
}

function buildPhotoPrintNotePreparationCacheKey(row) {
  if (!row || !Number.isInteger(row.id)) {
    return null;
  }

  return [
    row.id,
    row.file_path || '',
    normalizePhotoPrintNoteText(row.print_note_text) || '',
  ].join('\u001f');
}

async function pathExistsCached(targetPath, cache = null) {
  const normalizedTargetPath =
    typeof targetPath === 'string' ? targetPath.trim() : '';

  if (!normalizedTargetPath) {
    return false;
  }

  if (!(cache instanceof Map)) {
    return pathExists(normalizedTargetPath);
  }

  if (!cache.has(normalizedTargetPath)) {
    cache.set(normalizedTargetPath, pathExists(normalizedTargetPath));
  }

  return cache.get(normalizedTargetPath);
}

async function derivePhotoOrientationTierFromStoredAssets(
  row,
  pathExistsCache = null,
  stats = null
) {
  const thumbnailPath =
    typeof row?.thumbnail_path === 'string' ? row.thumbnail_path.trim() : '';

  if (thumbnailPath && (await pathExistsCached(thumbnailPath, pathExistsCache))) {
    if (stats) {
      stats.orientationAssetReads += 1;
    }
    const thumbnailMetadata = extractImageMetadataFromPath(thumbnailPath);

    if (thumbnailMetadata.orientationTier) {
      return thumbnailMetadata.orientationTier;
    }
  }

  const filePath = typeof row?.file_path === 'string' ? row.file_path.trim() : '';

  if (filePath && (await pathExistsCached(filePath, pathExistsCache))) {
    if (stats) {
      stats.orientationAssetReads += 1;
    }
    const imageMetadata = extractImageMetadataFromPath(filePath);

    if (imageMetadata.orientationTier) {
      return imageMetadata.orientationTier;
    }
  }

  return null;
}

async function ensurePhotoOrientationMetadata(
  row,
  pathExistsCache = null,
  stats = null
) {
  if (!row || !Number.isInteger(row.id)) {
    return row;
  }

  const storedOrientationTier = getStoredPhotoOrientationTier(row);

  if (storedOrientationTier) {
    if (stats) {
      stats.orientationStoredHits += 1;
    }

    const cacheKey = buildPhotoOrientationPreparationCacheKey(row);
    setBoundedCacheEntry(
      photoOrientationPreparationCache,
      cacheKey,
      storedOrientationTier
    );

    return row.orientation_tier === storedOrientationTier
      ? row
      : {
          ...row,
          orientation_tier: storedOrientationTier,
        };
  }

  const cacheKey = buildPhotoOrientationPreparationCacheKey(row);
  const hasCachedOrientationTier =
    cacheKey && photoOrientationPreparationCache.has(cacheKey);
  const cachedOrientationTier = hasCachedOrientationTier
    ? photoOrientationPreparationCache.get(cacheKey)
    : null;

  if (hasCachedOrientationTier) {
    if (stats) {
      stats.orientationCacheHits += 1;
    }
    return cachedOrientationTier === row.orientation_tier
      ? row
      : {
          ...row,
          orientation_tier: cachedOrientationTier,
        };
  }

  const actualOrientationTier = await derivePhotoOrientationTierFromStoredAssets(
    row,
    pathExistsCache,
    stats
  );

  if (!actualOrientationTier || row.orientation_tier === actualOrientationTier) {
    setBoundedCacheEntry(
      photoOrientationPreparationCache,
      cacheKey,
      row.orientation_tier || actualOrientationTier || ''
    );
    return row;
  }

  photoDb.updateImageMetadata(row.id, {
    imageWidth: row.image_width,
    imageHeight: row.image_height,
    resolutionTier: row.resolution_tier,
    orientationTier: actualOrientationTier,
  });

  return {
    ...row,
    orientation_tier: actualOrientationTier,
  };
}

async function ensurePhotoPrintNoteMetadata(
  row,
  pathExistsCache = null,
  stats = null
) {
  const cacheKey = buildPhotoPrintNotePreparationCacheKey(row);
  const cachedPrintNoteText = cacheKey
    ? photoPrintNotePreparationCache.get(cacheKey)
    : undefined;

  if (cachedPrintNoteText !== undefined) {
    if (stats) {
      stats.printNoteCacheHits += 1;
    }
    return cachedPrintNoteText === row.print_note_text
      ? row
      : {
          ...row,
          print_note_text: cachedPrintNoteText,
        };
  }

  if (
    row &&
    Number.isInteger(row.id) &&
    isEmptyPrintNotePlaceholder(row.print_note_text)
  ) {
    const savedRow = photoDb.updatePhotoPrintNote(row.id, null);
    setBoundedCacheEntry(photoPrintNotePreparationCache, cacheKey, '');
    return savedRow || {
      ...row,
      print_note_text: null,
    };
  }

  if (
    !row ||
    !Number.isInteger(row.id) ||
    !shouldRepairStoredPrintNote(row.print_note_text) ||
    !(await pathExistsCached(row.file_path, pathExistsCache))
  ) {
    if (cacheKey) {
      setBoundedCacheEntry(
        photoPrintNotePreparationCache,
        cacheKey,
        row?.print_note_text || ''
      );
    }
    return row;
  }

  try {
    if (stats) {
      stats.printNoteFileReads += 1;
    }
    const fileBuffer = await fs.readFile(row.file_path);
    const nextPrintNoteText = extractPhotoPrintNote(loadExifTagsSafely(fileBuffer));
    const savedRow = photoDb.updatePhotoPrintNote(row.id, nextPrintNoteText);
    setBoundedCacheEntry(
      photoPrintNotePreparationCache,
      cacheKey,
      nextPrintNoteText || ''
    );
    return savedRow || {
      ...row,
      print_note_text: nextPrintNoteText,
    };
  } catch {
    setBoundedCacheEntry(
      photoPrintNotePreparationCache,
      cacheKey,
      row.print_note_text || ''
    );
    return row;
  }
}

async function pathExists(targetPath) {
  if (typeof targetPath !== 'string' || targetPath.trim().length === 0) {
    return false;
  }

  try {
    await fs.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

function isManagedThumbnailPath(thumbnailPath) {
  if (typeof thumbnailPath !== 'string' || thumbnailPath.trim().length === 0) {
    return false;
  }

  const normalizedTarget = path.resolve(thumbnailPath);
  return getManagedThumbnailDirectoryPaths().some((managedDirectoryPath) => {
    const normalizedManagedDirectory = path.resolve(managedDirectoryPath);
    return (
      normalizedTarget === normalizedManagedDirectory ||
      normalizedTarget.startsWith(`${normalizedManagedDirectory}${path.sep}`)
    );
  });
}

async function resetManagedThumbnailDirectories() {
  for (const managedThumbnailDirectoryPath of getManagedThumbnailDirectoryPaths()) {
    await fs.rm(managedThumbnailDirectoryPath, {
      recursive: true,
      force: true,
      maxRetries: 2,
      retryDelay: 120,
    });
  }

  await ensureDir(thumbnailDirPath);
}

async function regenerateThumbnailRows(
  rows,
  targetMonth = null,
  progressReporter = null,
  {
    missingOnly = false,
    progressMessage = 'サムネイルを再生成中...',
  } = {}
) {
  const normalizedRows = Array.isArray(rows) ? rows : [];
  let regeneratedCount = 0;
  let skippedCount = 0;
  const failedFiles = [];
  let processedCount = 0;

  progressReporter?.({
    phase: 'process',
    current: 0,
    total: normalizedRows.length,
    message: progressMessage,
  });

  for (const row of normalizedRows) {
    try {
      const originalExists = await pathExists(row.file_path);

      if (!originalExists) {
        failedFiles.push({
          filePath: row.file_path,
          message: '元画像が見つかりません',
        });
        continue;
      }

      const nextThumbnailPath = await createThumbnail(
        row.file_path,
        row.file_hash,
        { force: true }
      );

      if (!nextThumbnailPath) {
        failedFiles.push({
          filePath: row.file_path,
          message: 'サムネイル生成に失敗しました',
        });
        continue;
      }

      photoDb.updateThumbnailPath(row.id, nextThumbnailPath);
      regeneratedCount += 1;
    } catch (error) {
      failedFiles.push({
        filePath: row.file_path,
        message: error.message,
      });
    } finally {
      processedCount += 1;
      progressReporter?.({
        phase: 'process',
        current: processedCount,
        total: normalizedRows.length,
        message: progressMessage,
      });
    }
  }

  return {
    ok: true,
    targetMonth,
    missingOnly,
    totalCount: normalizedRows.length,
    regeneratedCount,
    skippedCount,
    failedCount: failedFiles.length,
    failedFiles,
  };
}

async function regenerateManagedThumbnails(targetSelection = null, progressReporter = null) {
  const hasTargetMonth =
    Number.isInteger(targetSelection?.year) &&
    Number.isInteger(targetSelection?.month);
  const rows = hasTargetMonth
    ? photoDb.getPhotosByMonth(targetSelection.year, targetSelection.month)
    : photoDb.getAllPhotos();
  const targetMonth = hasTargetMonth
    ? {
        year: targetSelection.year,
        month: targetSelection.month,
      }
    : null;

  return regenerateThumbnailRows(rows, targetMonth, progressReporter);
}

async function clearManagedThumbnailCache(progressReporter = null) {
  if (!photoDb) {
    throw new Error('データベースが初期化されていません');
  }

  const rows = photoDb
    .getAllPhotos()
    .filter(
      (row) =>
        typeof row?.thumbnail_path === 'string' &&
        row.thumbnail_path.trim().length > 0 &&
        isManagedThumbnailPath(row.thumbnail_path)
    );

  progressReporter?.({
    phase: 'process',
    current: 0,
    total: rows.length,
    message: 'サムネイルキャッシュを削除しています...',
  });

  if (rows.length === 0) {
    return {
      ok: true,
      totalCount: 0,
      clearedCount: 0,
      failedCount: 0,
      failedFiles: [],
    };
  }

  const failedFiles = [];
  let processedCount = 0;
  let clearedCount = 0;

  try {
    await resetManagedThumbnailDirectories();
  } catch (error) {
    throw new Error(`サムネイル保存先を初期化できませんでした: ${error.message}`);
  }

  try {
    photoDb.clearThumbnailPaths(rows.map((row) => row.id));
    clearedCount = rows.length;
  } catch (error) {
    throw new Error(`サムネイル参照を更新できませんでした: ${error.message}`);
  } finally {
    processedCount = rows.length;
    progressReporter?.({
      phase: 'process',
      current: processedCount,
      total: rows.length,
      message: 'サムネイルキャッシュを削除しています...',
    });
  }

  return {
    ok: true,
    totalCount: rows.length,
    clearedCount,
    failedCount: failedFiles.length,
    failedFiles,
  };
}

async function deletePhotoRegistration(photoId) {
  if (!photoDb) {
    throw new Error('データベースが初期化されていません');
  }

  const deletedRow = photoDb.deletePhotoById(photoId);

  if (!deletedRow) {
    return {
      ok: false,
      message: '削除対象が見つかりませんでした',
    };
  }

  if (
    typeof deletedRow.thumbnail_path === 'string' &&
    deletedRow.thumbnail_path.trim().length > 0 &&
    isManagedThumbnailPath(deletedRow.thumbnail_path)
  ) {
    try {
      await fs.unlink(deletedRow.thumbnail_path);
    } catch {
      // サムネイルが既に無い場合は無視
    }
  }

  return {
    ok: true,
    deletedPhotoId: photoId,
  };
}

async function deletePhotoRegistrations(photoIds) {
  const normalizedIds = Array.from(
    new Set(
      (Array.isArray(photoIds) ? photoIds : []).filter(
        (photoId) => Number.isInteger(photoId) && photoId > 0
      )
    )
  );

  const deletedPhotoIds = [];
  const failed = [];

  for (const photoId of normalizedIds) {
    try {
      const result = await deletePhotoRegistration(photoId);

      if (result?.ok) {
        deletedPhotoIds.push(photoId);
      } else {
        failed.push({
          photoId,
          message: result?.message || '荳肴・縺ｪ繧ｨ繝ｩ繝ｼ',
        });
      }
    } catch (error) {
      failed.push({
        photoId,
        message: error.message,
      });
    }
  }

  return {
    ok: true,
    deletedPhotoIds,
    deletedCount: deletedPhotoIds.length,
    failed,
    failedCount: failed.length,
  };
}

async function deletePhotoRegistrationsByMonth(year, month) {
  const normalizedYear = Number.parseInt(year, 10);
  const normalizedMonth = Number.parseInt(month, 10);

  if (!Number.isInteger(normalizedYear) || !Number.isInteger(normalizedMonth)) {
    throw new Error('有効な年月が指定されていません');
  }

  const rows = photoDb.getPhotosByMonth(normalizedYear, normalizedMonth);
  const photoIds = rows
    .map((row) => row?.id)
    .filter((photoId) => Number.isInteger(photoId) && photoId > 0);

  const result = await deletePhotoRegistrations(photoIds);
  return {
    ...result,
    targetMonth: {
      year: normalizedYear,
      month: normalizedMonth,
    },
    totalCount: photoIds.length,
  };
}

async function deleteAllPhotoRegistrations() {
  const rows = photoDb.getAllPhotos();
  const photoIds = rows
    .map((row) => row?.id)
    .filter((photoId) => Number.isInteger(photoId) && photoId > 0);

  const result = await deletePhotoRegistrations(photoIds);
  return {
    ...result,
    totalCount: photoIds.length,
  };
}

// Full reset is reserved for maintenance/testing, so it clears both managed
// thumbnail files and every persisted table that belongs to this app.
async function resetApplicationData(progressReporter = null) {
  if (!photoDb) {
    throw new Error('データベースが初期化されていません');
  }

  progressReporter?.({
    phase: 'process',
    current: 0,
    total: 1,
    message: 'アプリデータを初期化しています...',
  });

  try {
    await resetManagedThumbnailDirectories();
  } catch (error) {
    throw new Error(`サムネイル保存先を初期化できませんでした: ${error.message}`);
  }

  const counts = photoDb.resetApplicationData();
  clearPhotoPreparationCaches();

  progressReporter?.({
    phase: 'process',
    current: 1,
    total: 1,
    message: 'アプリデータを初期化しています...',
  });

  return {
    ok: true,
    ...counts,
  };
}

function toRendererPhoto(row) {
  const resolvedPhotoLabels = Array.isArray(row?.photo_labels)
    ? row.photo_labels
    : photoDb && Number.isInteger(row?.id)
      ? photoDb.getPhotoTags(row.id)
      : [];
  const displayWorldName =
    normalizeDisplayWorldName(row.world_name_manual || row.world_name) ||
    UNKNOWN_WORLD_DISPLAY_NAME;
  const derivedOrientationTier =
    row.orientation_tier || getOrientationTier(row.image_width, row.image_height);

  const hasThumbnail =
    typeof row.thumbnail_path === 'string' && row.thumbnail_path.trim().length > 0;

  return {
    id: row.id,
    filePath: row.file_path,
    fileName: row.file_name,
    fileUrl: pathToFileURL(row.file_path).href,
    thumbnailPath: row.thumbnail_path,
    thumbnailUrl: hasThumbnail ? pathToFileURL(row.thumbnail_path).href : null,
    hasThumbnail,
    takenAt: row.taken_at,
    takenAtTimestamp: row.taken_at_timestamp,
    groupDate: row.group_date,
    year: row.year,
    month: row.month,
    day: row.day,
    worldId: row.world_id,
    worldName: displayWorldName,
    rawWorldName: row.world_name,
    worldNameManual: row.world_name_manual,
    worldUrl: row.world_url,
    isFavorite: Boolean(row.is_favorite),
    imageWidth: row.image_width,
    imageHeight: row.image_height,
    resolutionTier: row.resolution_tier,
    orientationTier: derivedOrientationTier,
    printNoteText: normalizePhotoPrintNoteText(row.print_note_text) || '',
    isPrintPhoto: isLikelyVrchatPrintPhoto(row),
    memoText: row.memo_text || '',
    photoLabels: resolvedPhotoLabels
      .map(toRendererPhotoLabel)
      .filter(Boolean),
  };
}

function attachPhotoLabelsToRows(rows) {
  const normalizedRows = Array.isArray(rows) ? rows : [];
  const photoIds = normalizedRows
    .map((row) => row?.id)
    .filter((photoId) => Number.isInteger(photoId) && photoId > 0);

  if (photoIds.length === 0) {
    return normalizedRows;
  }

  const labelsByPhotoId = new Map();
  const labelRows = photoDb.getPhotoTagsByPhotoIds(photoIds);

  for (const labelRow of labelRows) {
    const photoId = Number(labelRow.photo_id);

    if (!Number.isInteger(photoId) || photoId <= 0) {
      continue;
    }

    if (!labelsByPhotoId.has(photoId)) {
      labelsByPhotoId.set(photoId, []);
    }

    labelsByPhotoId.get(photoId).push(labelRow);
  }

  return normalizedRows.map((row) => ({
    ...row,
    photo_labels: labelsByPhotoId.get(row.id) || [],
  }));
}

async function prepareRowsForRenderer(rows) {
  const normalizedRows = Array.isArray(rows) ? rows : [];
  const pathExistsCache = new Map();
  const stats = {
    orientationStoredHits: 0,
    orientationCacheHits: 0,
    orientationAssetReads: 0,
    printNoteCacheHits: 0,
    printNoteFileReads: 0,
  };
  const totalStartedAt = performance.now();
  const orientationStartedAt = performance.now();
  const rowsWithAccurateOrientation = await mapWithConcurrencyLimit(
    normalizedRows,
    RENDERER_ROW_PREP_CONCURRENCY,
    (row) => ensurePhotoOrientationMetadata(row, pathExistsCache, stats)
  );
  const orientationMs = performance.now() - orientationStartedAt;
  const printNoteStartedAt = performance.now();
  const rowsWithResolvedPrintNotes = await mapWithConcurrencyLimit(
    rowsWithAccurateOrientation,
    RENDERER_ROW_PREP_CONCURRENCY,
    (row) => ensurePhotoPrintNoteMetadata(row, pathExistsCache, stats)
  );
  const printNoteMs = performance.now() - printNoteStartedAt;
  const labelStartedAt = performance.now();
  const rowsWithLabels = attachPhotoLabelsToRows(rowsWithResolvedPrintNotes);
  const labelsMs = performance.now() - labelStartedAt;
  const totalMs = performance.now() - totalStartedAt;

  logPerformanceTiming('prepareRowsForRenderer', totalMs, {
    rowCount: normalizedRows.length,
    pathCheckCount: pathExistsCache.size,
    orientationMs,
    printNoteMs,
    labelsMs,
    ...stats,
  });

  return rowsWithLabels;
}

async function buildRendererPhotosFromRows(rowProvider, label, details = {}) {
  const totalStartedAt = performance.now();
  const dbStartedAt = performance.now();
  const rows = typeof rowProvider === 'function' ? rowProvider() : [];
  const dbMs = performance.now() - dbStartedAt;
  const prepareStartedAt = performance.now();
  const preparedRows = await prepareRowsForRenderer(rows);
  const prepareMs = performance.now() - prepareStartedAt;
  const mapStartedAt = performance.now();
  const photos = preparedRows.map(toRendererPhoto);
  const mapMs = performance.now() - mapStartedAt;
  const totalMs = performance.now() - totalStartedAt;

  logPerformanceTiming(`renderer-photos:${label}`, totalMs, {
    ...details,
    rowCount: Array.isArray(rows) ? rows.length : 0,
    dbMs,
    prepareMs,
    mapMs,
  });

  return photos;
}

function buildWorldSidebarData(sortMode = 'count') {
  const rows = photoDb.getAllPhotosWithWorldInfo();
  const worldMap = new Map();

  for (const row of rows) {
    const worldId =
      normalizeWorldId(row.world_id) || parseWorldIdFromUrl(row.world_url);
    const worldName =
      normalizeDisplayWorldName(row.world_name_manual || row.world_name) ||
      UNKNOWN_WORLD_DISPLAY_NAME;
    const worldKey = worldId ? `id:${worldId}` : `name:${worldName}`;
    const existingEntry = worldMap.get(worldKey);

    if (existingEntry) {
      existingEntry.count += 1;
      continue;
    }

    worldMap.set(worldKey, {
      worldKey,
      worldId: worldId || null,
      worldName,
      count: 1,
    });
  }

  const entries = Array.from(worldMap.values());
  entries.sort((leftEntry, rightEntry) => {
    const leftName = leftEntry.worldName || '';
    const rightName = rightEntry.worldName || '';

    if (sortMode === 'name') {
      return leftName.localeCompare(rightName, 'ja');
    }

    const countDelta = rightEntry.count - leftEntry.count;
    return countDelta !== 0
      ? countDelta
      : leftName.localeCompare(rightName, 'ja');
  });

  return entries;
}

function toRendererWorldMetadata(row) {
  if (!row) {
    return null;
  }

  let tags = [];

  try {
    const parsedTags = JSON.parse(row.world_tags_json || '[]');
    tags = Array.isArray(parsedTags) ? parsedTags : [];
  } catch {
    tags = [];
  }

  return {
    worldId: row.world_id,
    sourceUrl: row.source_url,
    worldNameOfficial: normalizeOfficialWorldText(row.world_name_official),
    worldDescription: normalizeOfficialWorldText(row.world_description),
    worldTags: normalizeOfficialWorldTags(tags),
    authorId: row.author_id || null,
    authorName: normalizeOfficialWorldText(row.author_name),
    releaseStatus: row.release_status || null,
    imageUrl: row.image_url || null,
    thumbnailImageUrl: row.thumbnail_image_url || null,
    fetchStatus: row.fetch_status || null,
    fetchError: row.fetch_error || null,
    fetchedAt: row.fetched_at || null,
    lastAttemptedAt: row.last_attempted_at || null,
  };
}

function toRendererPhotoLabel(row) {
  if (!row) {
    return null;
  }

  const normalizedName =
    typeof row.normalized_name === 'string'
      ? row.normalized_name
      : typeof row.normalizedName === 'string'
        ? row.normalizedName
        : '';

  return {
    id: row.id,
    name: row.name || '',
    normalizedName,
    colorHex: normalizePhotoLabelColor(
      row.color_hex || row.colorHex,
      normalizedName
    ),
    photoCount: Number(row.photo_count ?? row.photoCount ?? 0) || 0,
  };
}

async function buildPhotoRecord(filePath) {
  const fileBuffer = await fs.readFile(filePath);
  const imageDetails = extractImageDetailsFromBuffer(fileBuffer);
  const rawWorldInfo = extractWorldInfoFromRawBuffer(fileBuffer);
  const takenAtFromFileName = parseVrchatFilenameDate(filePath);
  const imageDetailsFromFileName = parseImageDimensionsFromFilename(filePath);
  const tags = loadExifTagsSafely(fileBuffer);

  const fileHash = createHash('sha256').update(fileBuffer).digest('hex');
  const thumbnailPathPromise = createThumbnail(filePath, fileHash, {
    sourceImage: imageDetails.image,
  });
  const takenAtDate =
    takenAtFromFileName ||
    extractTakenAtFromTags(tags) ||
    (await fs.stat(filePath)).mtime;
  const worldInfo = extractWorldInfo(tags, fileBuffer, rawWorldInfo);
  const printNoteText = extractPhotoPrintNote(tags);
  const resolvedImageDetails = imageDetailsFromFileName
    ? {
        imageWidth: imageDetailsFromFileName.imageWidth,
        imageHeight: imageDetailsFromFileName.imageHeight,
        resolutionTier: imageDetailsFromFileName.resolutionTier,
        // Keep width / height filename-first for speed, but prefer the decoded
        // image orientation so portrait captures filter correctly.
        orientationTier:
          imageDetails.orientationTier || imageDetailsFromFileName.orientationTier,
      }
    : {
        imageWidth: imageDetails.imageWidth,
        imageHeight: imageDetails.imageHeight,
        resolutionTier: imageDetails.resolutionTier,
        orientationTier: imageDetails.orientationTier,
      };
  const nowIso = new Date().toISOString();
  const thumbnailPath = await thumbnailPathPromise;

  return {
    filePath,
    fileName: path.basename(filePath),
    fileHash,
    takenAt: formatDateTime(takenAtDate),
    takenAtTimestamp: takenAtDate.getTime(),
    groupDate: formatGroupDate(takenAtDate),
    year: takenAtDate.getFullYear(),
    month: takenAtDate.getMonth() + 1,
    day: takenAtDate.getDate(),
    worldId: worldInfo.worldId,
    worldName: worldInfo.worldName,
    worldNameManual: null,
    worldUrl: worldInfo.worldUrl,
    thumbnailPath,
    imageWidth: resolvedImageDetails.imageWidth,
    imageHeight: resolvedImageDetails.imageHeight,
    resolutionTier: resolvedImageDetails.resolutionTier,
    orientationTier: resolvedImageDetails.orientationTier,
    printNoteText,
    memoText: null,
    createdAt: nowIso,
    updatedAt: nowIso,
  };
}

async function mapWithConcurrencyLimit(
  items,
  concurrencyLimit,
  mapper,
  onProgress = null
) {
  const normalizedItems = Array.isArray(items) ? items : [];

  if (normalizedItems.length === 0) {
    return [];
  }

  const results = new Array(normalizedItems.length);
  const workerCount = Math.max(
    1,
    Math.min(concurrencyLimit || 1, normalizedItems.length)
  );
  let nextIndex = 0;
  let completedCount = 0;

  async function worker() {
    while (true) {
      const currentIndex = nextIndex;
      nextIndex += 1;

      if (currentIndex >= normalizedItems.length) {
        return;
      }

      results[currentIndex] = await mapper(normalizedItems[currentIndex], currentIndex);
      completedCount += 1;

      if (typeof onProgress === 'function') {
        onProgress({
          completedCount,
          totalCount: normalizedItems.length,
          item: normalizedItems[currentIndex],
          index: currentIndex,
          result: results[currentIndex],
        });
      }
    }
  }

  await Promise.all(Array.from({ length: workerCount }, () => worker()));

  return results;
}

function findLatestPhotoRecord(photoRecords) {
  let latestRecord = null;

  for (const photoRecord of photoRecords) {
    if (
      !latestRecord ||
      (photoRecord.takenAtTimestamp || 0) > (latestRecord.takenAtTimestamp || 0)
    ) {
      latestRecord = photoRecord;
    }
  }

  return latestRecord;
}

function buildSelectedMonthFromPhotoRecord(photoRecord) {
  if (
    !photoRecord ||
    !Number.isInteger(photoRecord.year) ||
    !Number.isInteger(photoRecord.month)
  ) {
    return null;
  }

  return {
    year: photoRecord.year,
    month: photoRecord.month,
  };
}

// Import and refresh flows share the same summary shape so renderer-side
// result handling can stay predictable.
function createImportSummaryResult(overrides = {}) {
  return {
    canceled: false,
    totalSelected: 0,
    importedCount: 0,
    newCount: 0,
    updatedCount: 0,
    relocatedCount: 0,
    worldMetadataTargets: [],
    failedCount: 0,
    failedFiles: [],
    selectedMonth: null,
    ...overrides,
  };
}

function createTrackedFolderRefreshResult(overrides = {}) {
  return {
    ...createImportSummaryResult(),
    noTrackedFolders: false,
    emptyRefresh: false,
    trackedFolderCount: 0,
    scannedFolderCount: 0,
    missingFolderPaths: [],
    skippedKnownCount: 0,
    backfilledTrackedFolderCount: 0,
    ...overrides,
  };
}

function createRegisteredPhotoReimportResult(overrides = {}) {
  return {
    ok: true,
    emptyReimport: false,
    registeredPhotoCount: 0,
    ...createImportSummaryResult(),
    ...overrides,
  };
}

// File access and path recovery return the same payload shape so open/show
// callers can share the same success handling.
function createResolvedPhotoAccessResult(
  resolvedPhoto,
  { includeFilePath = false } = {}
) {
  const payload = {
    ok: true,
    recovered: Boolean(resolvedPhoto?.recovered),
    photo: resolvedPhoto?.photo || null,
  };

  if (includeFilePath) {
    payload.filePath = resolvedPhoto?.filePath || null;
  }

  return payload;
}

function sanitizeEditedPhotoFileNamePart(value, fallback = 'photo') {
  const normalizedValue =
    typeof value === 'string' && value.trim().length > 0
      ? value.trim()
      : fallback;
  const sanitizedValue = normalizedValue
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();

  return sanitizedValue || fallback;
}

function formatEditedPhotoTimestamp(date = new Date()) {
  const safeDate = date instanceof Date && !Number.isNaN(date.getTime())
    ? date
    : new Date();

  return [
    safeDate.getFullYear(),
    pad2(safeDate.getMonth() + 1),
    pad2(safeDate.getDate()),
  ].join('') + '-' + [
    pad2(safeDate.getHours()),
    pad2(safeDate.getMinutes()),
    pad2(safeDate.getSeconds()),
  ].join('');
}

function normalizePhotoEditorOutputFormat(format) {
  return PHOTO_EDITOR_EXPORT_FORMATS[format]
    ? format
    : 'png';
}

function getPhotoEditorOutputFormatMeta(format) {
  return PHOTO_EDITOR_EXPORT_FORMATS[normalizePhotoEditorOutputFormat(format)];
}

function buildEditedPhotoDefaultPath(payload = {}) {
  const sourceFilePath = normalizeKnownFilePath(payload.sourceFilePath);
  const outputFormat = getPhotoEditorOutputFormatMeta(payload.outputFormat);
  const sourceFileName =
    typeof payload.sourceFileName === 'string' && payload.sourceFileName.trim()
      ? payload.sourceFileName.trim()
      : sourceFilePath
        ? path.basename(sourceFilePath)
        : 'photo.png';
  const parsedName = path.parse(sourceFileName);
  const baseName = sanitizeEditedPhotoFileNamePart(parsedName.name, 'photo');
  const outputFileName = `${baseName}_edited_${formatEditedPhotoTimestamp()}${outputFormat.extension}`;
  const outputDirectory = sourceFilePath
    ? path.dirname(sourceFilePath)
    : app.getPath('pictures');

  return path.join(outputDirectory, outputFileName);
}

function parseEditedPhotoDataUrl(dataUrl) {
  if (typeof dataUrl !== 'string') {
    throw new Error('編集済み画像データが不正です');
  }

  const match = PHOTO_EDITOR_DATA_URL_PATTERN.exec(dataUrl);

  if (!match) {
    throw new Error('編集済み画像データが不正です');
  }

  const mimeType = match[1];
  const formatKey = Object.entries(PHOTO_EDITOR_EXPORT_FORMATS).find(
    ([, meta]) => meta.mimeType === mimeType
  )?.[0];
  const formatMeta = getPhotoEditorOutputFormatMeta(formatKey);
  const base64Data = dataUrl.slice(match[0].length);
  const imageBuffer = Buffer.from(base64Data, 'base64');

  if (!imageBuffer || imageBuffer.length === 0) {
    throw new Error('編集済み画像データが空です');
  }

  return {
    imageBuffer,
    formatKey: formatKey || 'png',
    formatMeta,
  };
}

function isPngBuffer(buffer) {
  return (
    Buffer.isBuffer(buffer) &&
    buffer.length > 8 &&
    buffer.readUInt32BE(0) === 0x89504e47 &&
    buffer.readUInt32BE(4) === 0x0d0a1a0a
  );
}

function getPngChunks(buffer) {
  if (!isPngBuffer(buffer)) {
    return [];
  }

  const chunks = [];
  let offset = 8;

  while (offset + 12 <= buffer.length) {
    const length = buffer.readUInt32BE(offset);
    const chunkEnd = offset + 12 + length;

    if (chunkEnd > buffer.length) {
      return [];
    }

    const type = buffer.toString('ascii', offset + 4, offset + 8);
    chunks.push({
      type,
      buffer: buffer.subarray(offset, chunkEnd),
    });
    offset = chunkEnd;

    if (type === 'IEND') {
      break;
    }
  }

  return chunks;
}

function preservePngMetadataChunks(outputBuffer, sourceBuffer) {
  const metadataTypes = new Set(['tEXt', 'zTXt', 'iTXt', 'eXIf', 'tIME', 'pHYs']);
  const sourceChunks = getPngChunks(sourceBuffer).filter((chunk) =>
    metadataTypes.has(chunk.type)
  );

  if (sourceChunks.length === 0 || !isPngBuffer(outputBuffer)) {
    return outputBuffer;
  }

  const outputChunks = getPngChunks(outputBuffer).filter(
    (chunk) => !metadataTypes.has(chunk.type)
  );
  const firstIdatIndex = outputChunks.findIndex((chunk) => chunk.type === 'IDAT');
  const insertIndex = firstIdatIndex >= 0 ? firstIdatIndex : 1;

  return Buffer.concat([
    outputBuffer.subarray(0, 8),
    ...outputChunks.slice(0, insertIndex).map((chunk) => chunk.buffer),
    ...sourceChunks.map((chunk) => chunk.buffer),
    ...outputChunks.slice(insertIndex).map((chunk) => chunk.buffer),
  ]);
}

function isJpegBuffer(buffer) {
  return (
    Buffer.isBuffer(buffer) &&
    buffer.length > 4 &&
    buffer[0] === 0xff &&
    buffer[1] === 0xd8
  );
}

function getJpegExifSegments(buffer) {
  if (!isJpegBuffer(buffer)) {
    return [];
  }

  const segments = [];
  let offset = 2;

  while (offset + 4 <= buffer.length) {
    if (buffer[offset] !== 0xff) {
      break;
    }

    const marker = buffer[offset + 1];

    if (marker === 0xda || marker === 0xd9) {
      break;
    }

    const length = buffer.readUInt16BE(offset + 2);
    const segmentEnd = offset + 2 + length;

    if (length < 2 || segmentEnd > buffer.length) {
      break;
    }

    if (
      marker === 0xe1 &&
      buffer.toString('ascii', offset + 4, offset + 10) === 'Exif\u0000\u0000'
    ) {
      segments.push(buffer.subarray(offset, segmentEnd));
    }

    offset = segmentEnd;
  }

  return segments;
}

function preserveJpegExifSegments(outputBuffer, sourceBuffer) {
  const exifSegments = getJpegExifSegments(sourceBuffer);

  if (exifSegments.length === 0 || !isJpegBuffer(outputBuffer)) {
    return outputBuffer;
  }

  return Buffer.concat([
    outputBuffer.subarray(0, 2),
    ...exifSegments,
    outputBuffer.subarray(2),
  ]);
}

async function preserveEditedPhotoEmbeddedMetadata(
  imageBuffer,
  { sourceFilePath, formatKey } = {}
) {
  if (!sourceFilePath) {
    return imageBuffer;
  }

  try {
    const sourceBuffer = await fs.readFile(sourceFilePath);

    if (formatKey === 'png') {
      return preservePngMetadataChunks(imageBuffer, sourceBuffer);
    }

    if (formatKey === 'jpeg') {
      return preserveJpegExifSegments(imageBuffer, sourceBuffer);
    }
  } catch {
    // Embedded metadata is best-effort; saving the edited image takes priority.
  }

  return imageBuffer;
}

function isSameFilePath(leftPath, rightPath) {
  if (!leftPath || !rightPath) {
    return false;
  }

  return path.resolve(leftPath).toLowerCase() === path.resolve(rightPath).toLowerCase();
}

async function applyEditedPhotoSourceTimestamp(savePath, timestamp) {
  const numericTimestamp = Number(timestamp);

  if (!Number.isFinite(numericTimestamp) || numericTimestamp <= 0) {
    return;
  }

  const sourceDate = new Date(numericTimestamp);

  if (Number.isNaN(sourceDate.getTime())) {
    return;
  }

  try {
    await fs.utimes(savePath, sourceDate, sourceDate);
  } catch {
    // Keeping the image saved is more important than preserving file times.
  }
}

function normalizeEditedPhotoSourceMetadata(metadata = {}) {
  const timestamp = Number(metadata.takenAtTimestamp);
  const date = Number.isFinite(timestamp) && timestamp > 0
    ? new Date(timestamp)
    : null;
  const hasValidDate = date && !Number.isNaN(date.getTime());
  const toOptionalInteger = (value, fallback = null) => {
    const numericValue = Number(value);
    return Number.isInteger(numericValue) && numericValue > 0
      ? numericValue
      : fallback;
  };

  return {
    takenAt:
      typeof metadata.takenAt === 'string' && metadata.takenAt.trim()
        ? metadata.takenAt.trim()
        : hasValidDate
          ? formatDateTime(date)
          : null,
    takenAtTimestamp: hasValidDate ? timestamp : null,
    groupDate:
      typeof metadata.groupDate === 'string' && metadata.groupDate.trim()
        ? metadata.groupDate.trim()
        : hasValidDate
          ? formatGroupDate(date)
          : null,
    year: toOptionalInteger(metadata.year, hasValidDate ? date.getFullYear() : null),
    month: toOptionalInteger(metadata.month, hasValidDate ? date.getMonth() + 1 : null),
    day: toOptionalInteger(metadata.day, hasValidDate ? date.getDate() : null),
    worldId:
      typeof metadata.worldId === 'string' && metadata.worldId.trim()
        ? metadata.worldId.trim()
        : null,
    worldName:
      typeof metadata.worldName === 'string' && metadata.worldName.trim()
        ? metadata.worldName.trim()
        : null,
    worldNameManual:
      typeof metadata.worldNameManual === 'string' && metadata.worldNameManual.trim()
        ? metadata.worldNameManual.trim()
        : null,
    worldUrl:
      typeof metadata.worldUrl === 'string' && metadata.worldUrl.trim()
        ? metadata.worldUrl.trim()
        : null,
    isFavorite: Boolean(metadata.isFavorite),
    printNoteText:
      typeof metadata.printNoteText === 'string' && metadata.printNoteText.trim()
        ? metadata.printNoteText.trim()
        : null,
    memoText:
      typeof metadata.memoText === 'string' && metadata.memoText.trim()
        ? metadata.memoText.trim()
        : null,
    labels: normalizePhotoLabelPayload(metadata.labels),
  };
}

async function saveEditedPhotoFile(payload = {}) {
  const parsedImage = parseEditedPhotoDataUrl(payload.dataUrl);
  const { formatMeta } = parsedImage;
  const sourceFilePath = normalizeKnownFilePath(payload.sourceFilePath);
  const dialogResult = await dialog.showSaveDialog({
    title: '編集済み画像を保存',
    defaultPath: buildEditedPhotoDefaultPath({
      ...payload,
      outputFormat: parsedImage.formatKey,
    }),
    filters: [
      {
        name: formatMeta.filterName,
        extensions: [
          formatMeta.extension,
          ...(Array.isArray(formatMeta.aliases) ? formatMeta.aliases : []),
        ].map((extension) => extension.slice(1)),
      },
      {
        name: 'Images',
        extensions: Object.values(PHOTO_EDITOR_EXPORT_FORMATS).map((meta) =>
          meta.extension.slice(1)
        ),
      },
    ],
  });

  if (dialogResult.canceled || !dialogResult.filePath) {
    return {
      ok: true,
      canceled: true,
    };
  }

  let savePath = dialogResult.filePath;
  const saveExtension = path.extname(savePath).toLowerCase();
  const allowedExtensions = [
    formatMeta.extension,
    ...(Array.isArray(formatMeta.aliases) ? formatMeta.aliases : []),
  ];

  if (!allowedExtensions.includes(saveExtension)) {
    savePath += formatMeta.extension;
  }

  if (sourceFilePath && isSameFilePath(savePath, sourceFilePath)) {
    return {
      ok: false,
      message: '元画像は上書きできません。別名で保存してください。',
    };
  }

  await ensureDir(path.dirname(savePath));
  const imageBuffer = await preserveEditedPhotoEmbeddedMetadata(
    parsedImage.imageBuffer,
    {
      sourceFilePath,
      formatKey: parsedImage.formatKey,
    }
  );
  await fs.writeFile(savePath, imageBuffer);
  await applyEditedPhotoSourceTimestamp(savePath, payload.sourceTakenAtTimestamp);

  const imageHash = createHash('sha256').update(imageBuffer).digest('hex');
  let importResult = null;
  let importedPhoto = null;
  let importFailed = false;
  let importMessage = '';

  try {
    importResult = await importManyFiles([savePath]);
    let savedRow = photoDb.getPhotoByHash(imageHash);

    if (savedRow && payload.sourcePhotoMetadata) {
      const sourceMetadata = normalizeEditedPhotoSourceMetadata(
        payload.sourcePhotoMetadata
      );
      savedRow = photoDb.updateEditedPhotoSourceMetadata(
        savedRow.id,
        sourceMetadata
      ) || savedRow;

      if (Array.isArray(sourceMetadata.labels) && sourceMetadata.labels.length > 0) {
        photoDb.replacePhotoTags(savedRow.id, sourceMetadata.labels);
      }
    }

    importedPhoto = savedRow ? toRendererPhoto(savedRow) : null;
  } catch (error) {
    importFailed = true;
    importMessage = error.message;
  }

  return {
    ok: true,
    canceled: false,
    filePath: savePath,
    fileName: path.basename(savePath),
    importResult,
    importFailed,
    importMessage,
    photo: importedPhoto,
  };
}

// ------------------------------
// World metadata sync: queue target normalization
// ------------------------------
function createWorldMetadataSyncTarget(target) {
  const worldId =
    normalizeWorldId(target?.worldId) ||
    parseWorldIdFromUrl(target?.worldUrl);

  // Private / non-public captures can be stored without any world metadata in
  // the local file. VRChat 2025.3.2 notes that private worlds do not write
  // their ID/name into local photo metadata, so we intentionally skip queueing
  // here and let the UI use its explicit fallback label instead.
  if (!worldId) {
    return null;
  }

  return {
    worldId,
    worldUrl: target?.worldUrl || buildWorldUrlFromId(worldId),
  };
}

// ------------------------------
// Quarantine: mojibake-affected compatibility helpers
// ------------------------------
// These helpers remain only because this file previously went through an
// encoding-corruption incident. Active code paths delegate to the clean
// implementations further below so the live app does not rely on this block.
function createMissingAccessiblePhotoFileErrorLegacyCorrupt() {
  return createMissingAccessiblePhotoFileError();

  return new Error(
    '画像ファイルが見つかりません。保存先を開いて場所を再確認してください。'
  );

  return new Error(
    '画像ファイルが見つかりません。保存先を開いて場所を再確認してください。'
  );
}

function getWorldMetadataSyncProgressMessageLegacyCorrupt(phase) {
  return getWorldMetadataSyncProgressMessage(phase);

  return phase === 'complete'
    ? 'World情報の自動取得が完了しました'
    : 'World情報を自動で取得しています...';

  return phase === 'complete'
    ? 'World情報の自動取得が完了しました'
    : 'World情報を自動で取得しています...';
}

// Quarantined compatibility helper kept in place because the original
// mojibake-adjacent block is still preserved below. Active code now calls the
// *Active helper further down instead of routing through this function.
function createMissingAccessiblePhotoFileError() {
  return new Error(
    '画像ファイルが見つかりません。保存先を開いて場所を再確認してください。'
  );

  return new Error(
    '画像ファイルが見つかりません。保存先を開いて場所を再確認してください。'
  );
}

// Quarantined compatibility helper kept in place because the original
// mojibake-adjacent block is still preserved below. Active code now calls the
// *Active helper further down instead of routing through this function.
function getWorldMetadataSyncProgressMessage(phase) {
  return phase === 'complete'
    ? 'World情報の自動取得が完了しました'
    : 'World情報を自動で取得しています...';

  return phase === 'complete'
    ? 'World情報の自動取得が完了しました'
    : 'World情報を自動で取得しています...';
}

// ------------------------------
// Active world metadata sync helpers
// ------------------------------
// Active call sites use these UTF-8-safe helpers. The original function names
// above remain only because the file still contains quarantined mojibake
// blocks that are kept for reference and safe fallback only.
function createMissingAccessiblePhotoFileErrorActive() {
  return new Error(
    '画像ファイルが見つかりません。保存先を開いて場所を再確認してください。'
  );
}

function getWorldMetadataSyncProgressMessageActive(phase) {
  return phase === 'complete'
    ? 'World情報の自動取得が完了しました'
    : 'World情報を自動で取得しています...';
}

function broadcastQueuedWorldMetadataSyncProgress(phase, processedCount) {
  const totalCount = processedCount + pendingWorldMetadataSyncTargets.size;

  broadcastWorldMetadataSyncProgress({
    phase,
    current: processedCount,
    total: totalCount,
    message: getWorldMetadataSyncProgressMessageActive(phase),
  });
}

function queueWorldMetadataSyncTarget(target) {
  const normalizedTarget = createWorldMetadataSyncTarget(target);

  if (
    !normalizedTarget ||
    activeWorldMetadataSyncWorldId === normalizedTarget.worldId
  ) {
    return false;
  }

  if (pendingWorldMetadataSyncTargets.has(normalizedTarget.worldId)) {
    const existingTarget = pendingWorldMetadataSyncTargets.get(
      normalizedTarget.worldId
    );

    if (!existingTarget.worldUrl && normalizedTarget.worldUrl) {
      existingTarget.worldUrl = normalizedTarget.worldUrl;
    }

    return false;
  }

  pendingWorldMetadataSyncTargets.set(
    normalizedTarget.worldId,
    normalizedTarget
  );
  return true;
}

async function collectImageFilesFromDirectory(dirPath) {
  const results = [];
  const entries = await fs.readdir(dirPath, { withFileTypes: true });

  for (const entry of entries) {
    const fullPath = path.join(dirPath, entry.name);

    if (entry.isDirectory()) {
      const nestedFiles = await collectImageFilesFromDirectory(fullPath);
      results.push(...nestedFiles);
      continue;
    }

    if (entry.isFile() && isSupportedImageFile(fullPath)) {
      results.push(fullPath);
    }
  }

  return results;
}

function normalizeTrackedFolderPath(folderPath) {
  if (typeof folderPath !== 'string' || folderPath.trim().length === 0) {
    return null;
  }

  try {
    return path.resolve(folderPath.trim());
  } catch {
    return null;
  }
}

function normalizeKnownFilePath(filePath) {
  if (typeof filePath !== 'string' || filePath.trim().length === 0) {
    return null;
  }

  try {
    return path.resolve(filePath.trim());
  } catch {
    return null;
  }
}

function collapseTrackedFolderPaths(folderPaths) {
  const normalizedFolderPaths = Array.from(
    new Set(
      (Array.isArray(folderPaths) ? folderPaths : [])
        .map(normalizeTrackedFolderPath)
        .filter(Boolean)
    )
  ).sort((a, b) => a.length - b.length || a.localeCompare(b));

  const collapsedPaths = [];

  for (const folderPath of normalizedFolderPaths) {
    const isCoveredByAncestor = collapsedPaths.some(
      (existingPath) =>
        folderPath === existingPath ||
        folderPath.startsWith(`${existingPath}${path.sep}`)
    );

    if (!isCoveredByAncestor) {
      collapsedPaths.push(folderPath);
    }
  }

  return collapsedPaths;
}

function registerTrackedFolders(folderPaths) {
  if (!photoDb) {
    return;
  }

  const normalizedFolderPaths = collapseTrackedFolderPaths(folderPaths);

  for (const folderPath of normalizedFolderPaths) {
    photoDb.upsertTrackedFolder(folderPath);
  }
}

function getTrackedFolderPaths(targetFolderPaths) {
  if (!photoDb) {
    return [];
  }

  const sourcePaths =
    Array.isArray(targetFolderPaths) && targetFolderPaths.length > 0
      ? targetFolderPaths
      : photoDb.getTrackedFolders().map((row) => row.folder_path);

  return Array.from(
    new Set(sourcePaths.map(normalizeTrackedFolderPath).filter(Boolean))
  );
}

function deriveTrackedFoldersFromExistingPhotos() {
  if (!photoDb) {
    return [];
  }

  const candidateFolderPaths = photoDb
    .getAllPhotos()
    .map((row) => normalizeKnownFilePath(row.file_path))
    .filter(Boolean)
    .map((filePath) => path.dirname(filePath));

  return collapseTrackedFolderPaths(candidateFolderPaths);
}

function getKnownPhotoPathSet() {
  if (!photoDb) {
    return new Set();
  }

  return new Set(
    photoDb
      .getAllPhotos()
      .map((row) => normalizeKnownFilePath(row.file_path))
      .filter(Boolean)
  );
}

function normalizeFileNameForComparison(fileName) {
  if (typeof fileName !== 'string' || fileName.trim().length === 0) {
    return null;
  }

  return fileName.trim().toLowerCase();
}

function buildPhotoRecoverySearchRoots(row) {
  const trackedFolderPaths = getTrackedFolderPaths();
  const fallbackRoots = [];
  let currentPath = normalizeKnownFilePath(path.dirname(row?.file_path || ''));

  for (let depth = 0; currentPath && depth < 3; depth += 1) {
    fallbackRoots.push(currentPath);

    const nextPath = normalizeKnownFilePath(path.dirname(currentPath));
    if (!nextPath || nextPath === currentPath) {
      break;
    }

    currentPath = nextPath;
  }

  return collapseTrackedFolderPaths([...trackedFolderPaths, ...fallbackRoots]);
}

async function findMatchingPhotoPathInDirectory(
  dirPath,
  targetFileName,
  targetFileHash,
  excludedFilePath
) {
  let entries = [];

  try {
    entries = await fs.readdir(dirPath, { withFileTypes: true });
  } catch {
    return null;
  }

  const normalizedTargetFileName = normalizeFileNameForComparison(targetFileName);

  for (const entry of entries) {
    const fullPath = path.join(dirPath, entry.name);

    if (entry.isDirectory()) {
      const nestedMatch = await findMatchingPhotoPathInDirectory(
        fullPath,
        targetFileName,
        targetFileHash,
        excludedFilePath
      );

      if (nestedMatch) {
        return nestedMatch;
      }

      continue;
    }

    if (!entry.isFile() || !isSupportedImageFile(fullPath)) {
      continue;
    }

    if (
      normalizeFileNameForComparison(entry.name) !== normalizedTargetFileName
    ) {
      continue;
    }

    const normalizedFullPath = normalizeKnownFilePath(fullPath);

    if (!normalizedFullPath || normalizedFullPath === excludedFilePath) {
      continue;
    }

    try {
      const fileBuffer = await fs.readFile(normalizedFullPath);
      const fileHash = createHash('sha256').update(fileBuffer).digest('hex');

      if (fileHash === targetFileHash) {
        return normalizedFullPath;
      }
    } catch {
      // 候補ファイルが読み取れない場合は次を試す
    }
  }

  return null;
}

async function tryRecoverMovedPhoto(photoId) {
  if (!photoDb || !Number.isInteger(photoId) || photoId <= 0) {
    return null;
  }

  let row = photoDb.getPhotoById(photoId);

  if (
    !row ||
    typeof row.file_hash !== 'string' ||
    row.file_hash.trim().length === 0 ||
    typeof row.file_name !== 'string' ||
    row.file_name.trim().length === 0
  ) {
    return null;
  }

  const excludedFilePath = normalizeKnownFilePath(row.file_path);
  const searchRoots = buildPhotoRecoverySearchRoots(row);

  for (const rootPath of searchRoots) {
    const recoveredFilePath = await findMatchingPhotoPathInDirectory(
      rootPath,
      row.file_name,
      row.file_hash,
      excludedFilePath
    );

    if (!recoveredFilePath) {
      continue;
    }

    registerTrackedFolders([path.dirname(recoveredFilePath)]);

    const savedRow = photoDb.updatePhotoFileLocation(
      photoId,
      recoveredFilePath,
      path.basename(recoveredFilePath)
    );

    return savedRow ? toRendererPhoto(savedRow) : null;
  }

  return null;
}

function normalizePhotoAccessPayload(payload) {
  if (typeof payload === 'string') {
    return {
      photoId: null,
      filePath: normalizeKnownFilePath(payload),
    };
  }

  if (!payload || typeof payload !== 'object') {
    return {
      photoId: null,
      filePath: null,
    };
  }

  return {
    photoId:
      Number.isInteger(payload.photoId) && payload.photoId > 0
        ? payload.photoId
        : null,
    filePath: normalizeKnownFilePath(payload.filePath),
  };
}

async function ensureAccessiblePhotoFile(payload) {
  const normalizedPayload = normalizePhotoAccessPayload(payload);

  if (normalizedPayload.filePath && (await pathExists(normalizedPayload.filePath))) {
    return {
      filePath: normalizedPayload.filePath,
      photo: null,
      recovered: false,
    };
  }

  const recoveredPhoto = await tryRecoverMovedPhoto(normalizedPayload.photoId);

  if (
    recoveredPhoto?.filePath &&
    (await pathExists(recoveredPhoto.filePath))
  ) {
    return {
      filePath: recoveredPhoto.filePath,
      photo: recoveredPhoto,
      recovered: true,
    };
  }

  throw createMissingAccessiblePhotoFileErrorActive();
}

// Disk actions all share the same recovery flow: recover the moved file if
// needed, verify the final path exists, then run the specific shell action.
// This keeps "open image", "open folder", and future shell actions aligned.
async function withAccessiblePhotoFile(payload, action) {
  const resolvedPhoto = await ensureAccessiblePhotoFile(payload);
  const filePath = resolvedPhoto.filePath;

  if (typeof filePath !== 'string' || filePath.trim().length === 0) {
    throw createMissingAccessiblePhotoFileErrorActive();
  }

  await fs.access(filePath);
  await action(filePath);

  return createResolvedPhotoAccessResult(resolvedPhoto);
}

async function collectNewFilesFromTrackedFolders(
  targetFolderPaths,
  progressReporter = null
) {
  let trackedFolderPaths = getTrackedFolderPaths(targetFolderPaths);
  const knownPhotoPathSet = getKnownPhotoPathSet();
  const newFilePaths = [];
  const scannedFolderPaths = [];
  const missingFolderPaths = [];
  let skippedKnownCount = 0;
  let backfilledTrackedFolderCount = 0;

  if (
    trackedFolderPaths.length === 0 &&
    (!Array.isArray(targetFolderPaths) || targetFolderPaths.length === 0)
  ) {
    trackedFolderPaths = deriveTrackedFoldersFromExistingPhotos();
    backfilledTrackedFolderCount = trackedFolderPaths.length;

    if (trackedFolderPaths.length > 0) {
      registerTrackedFolders(trackedFolderPaths);
    }
  }

  progressReporter?.({
    phase: 'scan',
    current: 0,
    total: trackedFolderPaths.length,
    message: '追跡フォルダを確認中...',
  });

  for (const folderPath of trackedFolderPaths) {
    try {
      const stats = await fs.stat(folderPath);

      if (!stats.isDirectory()) {
        missingFolderPaths.push(folderPath);
        continue;
      }

      scannedFolderPaths.push(folderPath);

      const nestedFiles = await collectImageFilesFromDirectory(folderPath);

      for (const filePath of nestedFiles) {
        const normalizedFilePath = normalizeKnownFilePath(filePath);

        if (!normalizedFilePath) {
          continue;
        }

        if (knownPhotoPathSet.has(normalizedFilePath)) {
          skippedKnownCount += 1;
          continue;
        }

        knownPhotoPathSet.add(normalizedFilePath);
        newFilePaths.push(normalizedFilePath);
      }
    } catch {
      missingFolderPaths.push(folderPath);
    }

    progressReporter?.({
      phase: 'scan',
      current: scannedFolderPaths.length + missingFolderPaths.length,
      total: trackedFolderPaths.length,
      message: '追跡フォルダを確認中...',
    });
  }

  return {
    trackedFolderPaths,
    scannedFolderPaths,
    missingFolderPaths,
    newFilePaths,
    skippedKnownCount,
    backfilledTrackedFolderCount,
  };
}

async function refreshTrackedFolderImports(
  targetFolderPaths,
  progressReporter = null
) {
  const {
    trackedFolderPaths,
    scannedFolderPaths,
    missingFolderPaths,
    newFilePaths,
    skippedKnownCount,
    backfilledTrackedFolderCount,
  } = await collectNewFilesFromTrackedFolders(targetFolderPaths, progressReporter);

  if (scannedFolderPaths.length > 0) {
    registerTrackedFolders(scannedFolderPaths);
  }

  if (trackedFolderPaths.length === 0) {
    return createTrackedFolderRefreshResult({
      noTrackedFolders: true,
    });
  }

  if (newFilePaths.length === 0) {
    return createTrackedFolderRefreshResult({
      emptyRefresh: true,
      trackedFolderCount: trackedFolderPaths.length,
      scannedFolderCount: scannedFolderPaths.length,
      missingFolderPaths,
      skippedKnownCount,
      backfilledTrackedFolderCount,
    });
  }

  const importResult = await importManyFiles(newFilePaths, progressReporter);

  return createTrackedFolderRefreshResult({
    ...importResult,
    trackedFolderCount: trackedFolderPaths.length,
    scannedFolderCount: scannedFolderPaths.length,
    missingFolderPaths,
    skippedKnownCount,
    backfilledTrackedFolderCount,
  });
}

async function reimportRegisteredPhotos(targetSelection, progressReporter = null) {
  const normalizedTargetSelection =
    Number.isInteger(targetSelection?.year) && Number.isInteger(targetSelection?.month)
      ? {
          year: targetSelection.year,
          month: targetSelection.month,
        }
      : null;

  if (!normalizedTargetSelection) {
    return createRegisteredPhotoReimportResult({
      ok: false,
      message: '再取り込みする年月を指定してください',
    });
  }

  const registeredPhotoPaths = Array.from(
    new Set(
      photoDb
        .getPhotosByMonth(
          normalizedTargetSelection.year,
          normalizedTargetSelection.month
        )
        .map((row) => normalizeKnownFilePath(row?.file_path))
        .filter(
          (filePath) =>
            typeof filePath === 'string' &&
            filePath.trim().length > 0 &&
            isSupportedImageFile(filePath)
        )
    )
  );

  if (registeredPhotoPaths.length === 0) {
    return createRegisteredPhotoReimportResult({
      emptyReimport: true,
      selectedMonth: normalizedTargetSelection,
      targetMonth: normalizedTargetSelection,
    });
  }

  const importResult = await importManyFiles(registeredPhotoPaths, progressReporter);

  return createRegisteredPhotoReimportResult({
    ...importResult,
    selectedMonth: normalizedTargetSelection,
    targetMonth: normalizedTargetSelection,
    registeredPhotoCount: registeredPhotoPaths.length,
  });
}

async function importManyFiles(filePaths, progressReporter = null) {
  const uniquePaths = Array.from(
    new Set(filePaths.filter((filePath) => isSupportedImageFile(filePath)))
  );

  const photoRecords = [];
  const failedFiles = [];

  progressReporter?.({
    phase: 'process',
    current: 0,
    total: uniquePaths.length,
    message: '画像を取り込み中...',
  });

  const buildResults = await mapWithConcurrencyLimit(
    uniquePaths,
    IMPORT_PREPROCESS_CONCURRENCY,
    async (filePath) => {
      try {
        return {
          ok: true,
          photoRecord: await buildPhotoRecord(filePath),
        };
      } catch (error) {
        return {
          ok: false,
          filePath,
          message: error.message,
        };
      }
    },
    ({ completedCount, totalCount }) => {
      progressReporter?.({
        phase: 'process',
        current: completedCount,
        total: totalCount,
        message: '画像を取り込み中...',
      });
    }
  );

  for (const result of buildResults) {
    if (result?.ok && result.photoRecord) {
      photoRecords.push(result.photoRecord);
      continue;
    }

    failedFiles.push({
      filePath: result?.filePath || '',
      message: result?.message || '不明なエラー',
    });
  }

  const existingPhotoMap = new Map(
    photoDb
      .getPhotosByHashes(photoRecords.map((photoRecord) => photoRecord.fileHash))
      .map((row) => [row.file_hash, row])
  );
  const existingHashSet = new Set(existingPhotoMap.keys());

  progressReporter?.({
    phase: 'save',
    current: photoRecords.length,
    total: uniquePaths.length,
    message: '取り込み結果を保存中...',
  });

  try {
    photoDb.insertOrUpdatePhotos(photoRecords);
  } catch {
    for (const photoRecord of photoRecords) {
      photoDb.insertOrUpdatePhoto(photoRecord);
    }
  }

  const thumbnailRecoveryTargets = photoRecords.filter(
    (photoRecord) => !photoRecord.thumbnailPath
  );

  let recoveredCount = 0;

  for (const photoRecord of thumbnailRecoveryTargets) {
    const savedRow = photoDb.getPhotoByHash(photoRecord.fileHash);

    if (!savedRow) {
      continue;
    }

    await ensureManagedThumbnailForRow(savedRow);
    recoveredCount += 1;

    progressReporter?.({
      phase: 'thumbnail-recovery',
      current: recoveredCount,
      total: thumbnailRecoveryTargets.length,
      message: '不足サムネイルを補完中...',
    });
  }

  let newCount = 0;
  let updatedCount = 0;
  let relocatedCount = 0;

  for (const photoRecord of photoRecords) {
    const existingPhoto = existingPhotoMap.get(photoRecord.fileHash);

    if (existingPhoto) {
      updatedCount += 1;

      const previousFilePath = normalizeKnownFilePath(existingPhoto.file_path);
      const nextFilePath = normalizeKnownFilePath(photoRecord.filePath);

      if (previousFilePath && nextFilePath && previousFilePath !== nextFilePath) {
        relocatedCount += 1;
      }
    } else {
      newCount += 1;
    }
  }

  const latestRow = findLatestPhotoRecord(photoRecords);

  progressReporter?.({
    phase: 'complete',
    current: uniquePaths.length,
    total: uniquePaths.length,
    message: '取り込みが完了しました',
  });

  return createImportSummaryResult({
    totalSelected: uniquePaths.length,
    importedCount: photoRecords.length,
    newCount,
    updatedCount,
    relocatedCount,
    worldMetadataTargets: buildWorldMetadataSyncTargetsFromPhotoRecords(photoRecords),
    failedCount: failedFiles.length,
    failedFiles,
    selectedMonth: buildSelectedMonthFromPhotoRecord(latestRow),
  });
}

function delay(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

// World metadata sync works on unique world IDs so post-import enrichment can
// reuse cache entries instead of re-fetching per photo. Files without a
// resolvable worldId are filtered out before they reach this stage.
function buildWorldMetadataSyncTargetsFromPhotoRecords(photoRecords) {
  const targets = new Map();

  for (const photoRecord of Array.isArray(photoRecords) ? photoRecords : []) {
    const syncTarget = createWorldMetadataSyncTarget({
      worldId: photoRecord?.worldId,
      worldUrl: photoRecord?.worldUrl,
    });

    if (!syncTarget) {
      continue;
    }

    if (!targets.has(syncTarget.worldId)) {
      targets.set(syncTarget.worldId, syncTarget);
      continue;
    }

    const existingTarget = targets.get(syncTarget.worldId);

    if (!existingTarget.worldUrl && syncTarget.worldUrl) {
      existingTarget.worldUrl = syncTarget.worldUrl;
    }
  }

  return Array.from(targets.values());
}

function hasUsableOfficialWorldMetadataRow(row) {
  if (!row) {
    return false;
  }

  const officialName = normalizeOfficialWorldText(row.world_name_official);
  const officialDescription = normalizeOfficialWorldText(row.world_description);

  let officialTags = [];

  try {
    const parsedTags = JSON.parse(row.world_tags_json || '[]');
    officialTags = Array.isArray(parsedTags) ? parsedTags : [];
  } catch {
    officialTags = [];
  }

  return Boolean(officialName || officialDescription || officialTags.length > 0);
}

function addWorldMetadataSyncSubscriber(webContents) {
  if (!webContents || webContents.isDestroyed()) {
    return;
  }

  worldMetadataSyncSubscribers.add(webContents);
}

function broadcastWorldMetadataSyncProgress(payload = {}) {
  for (const webContents of Array.from(worldMetadataSyncSubscribers)) {
    if (!webContents || webContents.isDestroyed()) {
      worldMetadataSyncSubscribers.delete(webContents);
      continue;
    }

    webContents.send(PROCESSING_PROGRESS_CHANNEL, {
      operation: WORLD_METADATA_SYNC_OPERATION,
      ...payload,
    });
  }
}

function broadcastWorldMetadataUpdated(rows) {
  const photos = (Array.isArray(rows) ? rows : [])
    .map((row) => toRendererPhoto(row))
    .filter(Boolean);

  if (photos.length === 0) {
    return;
  }

  for (const webContents of Array.from(worldMetadataSyncSubscribers)) {
    if (!webContents || webContents.isDestroyed()) {
      worldMetadataSyncSubscribers.delete(webContents);
      continue;
    }

    webContents.send(WORLD_METADATA_UPDATED_CHANNEL, {
      photos,
    });
  }
}

async function syncOfficialWorldMetadataForTarget(target) {
  const normalizedTarget = createWorldMetadataSyncTarget(target);

  if (!normalizedTarget) {
    return {
      didFetch: false,
      updatedRows: [],
    };
  }

  const { worldId, worldUrl } = normalizedTarget;
  let metadataRow = photoDb.getWorldMetadataByWorldId(worldId);
  let didFetch = false;

  if (!hasUsableOfficialWorldMetadataRow(metadataRow)) {
    await delay(WORLD_METADATA_SYNC_DELAY_MS);
    metadataRow =
      (await fetchAndCacheOfficialWorldMetadata(worldId, worldUrl)) ||
      photoDb.getWorldMetadataByWorldId(worldId);
    didFetch = true;
  }

  const officialWorldName = normalizeOfficialWorldText(
    metadataRow?.world_name_official
  );

  if (!officialWorldName) {
    return {
      didFetch,
      updatedRows: [],
    };
  }

  return {
    didFetch,
    updatedRows: photoDb.updateAutoWorldInfoByWorldId(worldId, {
      worldName: officialWorldName,
      worldUrl,
    }),
  };
}

function getQueuedWorldMetadataSyncCount() {
  return pendingWorldMetadataSyncTargets.size + (isWorldMetadataSyncRunning ? 1 : 0);
}

// Queue iteration is intentionally split into "dequeue" and "process" so the
// active sync loop reads as state transitions rather than one large mutation.
function dequeueNextWorldMetadataSyncTarget() {
  if (pendingWorldMetadataSyncTargets.size === 0) {
    return null;
  }

  const [worldId, target] = pendingWorldMetadataSyncTargets.entries().next().value;
  pendingWorldMetadataSyncTargets.delete(worldId);
  activeWorldMetadataSyncWorldId = worldId;

  return { worldId, target };
}

async function processQueuedWorldMetadataSyncTarget(target) {
  try {
    const result = await syncOfficialWorldMetadataForTarget(target);

    if (result.updatedRows.length > 0) {
      broadcastWorldMetadataUpdated(result.updatedRows);
    }
  } catch (error) {
    console.warn(
      '[world-metadata-sync] Failed to sync official world metadata',
      target,
      error
    );
  }
}

// Active world metadata processing flows through this clean queue runner.
// The legacy runner below only remains as an encoding-safe quarantine block.
async function runQueuedWorldMetadataSyncInternal() {
  if (isWorldMetadataSyncRunning) {
    return;
  }

  isWorldMetadataSyncRunning = true;
  let processedCount = 0;

  try {
    while (pendingWorldMetadataSyncTargets.size > 0) {
      const nextTarget = dequeueNextWorldMetadataSyncTarget();

      if (!nextTarget) {
        break;
      }

      broadcastQueuedWorldMetadataSyncProgress('process', processedCount);
      await processQueuedWorldMetadataSyncTarget(nextTarget.target);

      processedCount += 1;
      broadcastQueuedWorldMetadataSyncProgress(
        pendingWorldMetadataSyncTargets.size > 0 ? 'process' : 'complete',
        processedCount
      );
    }
  } finally {
    isWorldMetadataSyncRunning = false;
    activeWorldMetadataSyncWorldId = null;
  }
}

// World metadata sync runs as a deduplicated queue keyed by worldId so imports
// can stay fast while official data is fetched in the background.
// Legacy fallback retained only as an encoding-safe quarantine block.
async function runQueuedWorldMetadataSyncLegacy() {
  return runQueuedWorldMetadataSyncInternal();

  // Intentionally unreachable: preserved only so the mojibake block can be
  // compared against the clean implementation during a later UTF-8 cleanup.
  if (isWorldMetadataSyncRunning) {
    return;
  }

  isWorldMetadataSyncRunning = true;
  let processedCount = 0;

  try {
    while (pendingWorldMetadataSyncTargets.size > 0) {
      const [worldId, target] =
        pendingWorldMetadataSyncTargets.entries().next().value;

      pendingWorldMetadataSyncTargets.delete(worldId);
      activeWorldMetadataSyncWorldId = worldId;

      broadcastWorldMetadataSyncProgress({
        phase: 'process',
        current: processedCount,
        total: totalCount,
        message: 'World情報を自動で同期しています...',
      });

      let result = { updatedRows: [] };

      try {
        result = await syncOfficialWorldMetadataForTarget(target);
      } catch (error) {
        console.warn(
          '[world-metadata-sync] Failed to sync official world metadata',
          target,
          error
        );
      }

      if (result.updatedRows.length > 0) {
        broadcastWorldMetadataUpdated(result.updatedRows);
      }

      processedCount += 1;

      broadcastWorldMetadataSyncProgress({
        phase: pendingWorldMetadataSyncTargets.size > 0 ? 'process' : 'complete',
        current: processedCount,
        total: processedCount + pendingWorldMetadataSyncTargets.size,
        message:
          pendingWorldMetadataSyncTargets.size > 0
            ? 'World情報を自動で同期しています...'
            : 'World情報の自動同期が完了しました',
      });
    }
  } finally {
    isWorldMetadataSyncRunning = false;
    activeWorldMetadataSyncWorldId = null;
  }
}

async function runQueuedWorldMetadataSyncActive() {
  return runQueuedWorldMetadataSyncInternal();
}

// Compatibility wrapper retained so older references still route into the
// clean active runner above while quarantine cleanup continues.
async function runQueuedWorldMetadataSync() {
  return runQueuedWorldMetadataSyncActive();
}

// Renderer requests enter the sync queue through this helper so subscriber
// registration, dedupe, and queue start stay in one active entry point.
function startWorldMetadataSyncForSubscriber(targets, webContents) {
  addWorldMetadataSyncSubscriber(webContents);

  let queuedCount = 0;

  for (const target of Array.isArray(targets) ? targets : []) {
    if (queueWorldMetadataSyncTarget(target)) {
      queuedCount += 1;
    }
  }

  if (queuedCount > 0) {
    void runQueuedWorldMetadataSyncActive();
  }

  return {
    ok: true,
    queuedCount,
    pendingCount: getQueuedWorldMetadataSyncCount(),
  };
}

// Drag-and-drop can mix files and folders, so expand them here before import so
// the higher-level import flow only deals with normalized image paths.
async function expandDroppedPathsToImportTargets(droppedPaths) {
  const collected = [];
  const trackedFolderPaths = [];

  for (const droppedPath of droppedPaths) {
    if (typeof droppedPath !== 'string' || droppedPath.trim().length === 0) {
      continue;
    }

    try {
      const stats = await fs.stat(droppedPath);

      if (stats.isDirectory()) {
        trackedFolderPaths.push(droppedPath);
        const nestedFiles = await collectImageFilesFromDirectory(droppedPath);
        collected.push(...nestedFiles);
        continue;
      }

      if (stats.isFile() && isSupportedImageFile(droppedPath)) {
        collected.push(droppedPath);
      }
    } catch {
      // 読み取れないパスはスキップ
    }
  }

  return {
    filePaths: Array.from(new Set(collected)),
    trackedFolderPaths: Array.from(
      new Set(trackedFolderPaths.map(normalizeTrackedFolderPath).filter(Boolean))
    ),
  };
}

function parseWorldTitleAndAuthor(rawTitle) {
  const normalizedTitle = normalizeOfficialWorldText(rawTitle);

  if (!normalizedTitle) {
    return {
      worldNameOfficial: null,
      authorName: null,
    };
  }

  const matched = normalizedTitle.match(/^(.*)\s+by\s+(.+)$/i);

  if (!matched) {
    return {
      worldNameOfficial: normalizedTitle,
      authorName: null,
    };
  }

  return {
    worldNameOfficial: normalizeOfficialWorldText(matched[1]),
    authorName: normalizeOfficialWorldText(matched[2]),
  };
}

async function fetchTextWithTimeout(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      'user-agent': 'VRChatWorldPhotoManager/1.0',
      ...options.headers,
    },
    signal: AbortSignal.timeout(WORLD_METADATA_FETCH_TIMEOUT_MS),
  });

  return {
    response,
    text: await response.text(),
  };
}

async function fetchOfficialWorldMetadataFromApi(worldId) {
  const normalizedWorldId = normalizeWorldId(worldId);

  if (!normalizedWorldId) {
    return null;
  }

  const apiUrl = `${WORLD_METADATA_API_BASE_URL}/${normalizedWorldId}`;

  try {
    const { response, text } = await fetchTextWithTimeout(apiUrl, {
      headers: {
        accept: 'application/json',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        fetchStatus:
          response.status === 401 || response.status === 403
            ? 'unauthorized'
            : response.status === 404
              ? 'not_found'
              : 'error',
        fetchError: `HTTP ${response.status}`,
      };
    }

    const payload = JSON.parse(text);

    return {
      ok: true,
      worldId: normalizedWorldId,
      sourceUrl: buildWorldUrlFromId(normalizedWorldId),
      worldNameOfficial: normalizeOfficialWorldText(payload.name),
      worldDescription: normalizeOfficialWorldText(payload.description),
      worldTags: normalizeOfficialWorldTags(payload.tags),
      authorId:
        typeof payload.authorId === 'string' && payload.authorId.trim().length > 0
          ? payload.authorId.trim()
          : null,
      authorName: normalizeOfficialWorldText(payload.authorName),
      releaseStatus:
        typeof payload.releaseStatus === 'string' && payload.releaseStatus.trim().length > 0
          ? payload.releaseStatus.trim()
          : null,
      imageUrl:
        typeof payload.imageUrl === 'string' && payload.imageUrl.trim().length > 0
          ? payload.imageUrl.trim()
          : null,
      thumbnailImageUrl:
        typeof payload.thumbnailImageUrl === 'string' &&
        payload.thumbnailImageUrl.trim().length > 0
          ? payload.thumbnailImageUrl.trim()
          : null,
      fetchStatus: 'success',
      fetchError: null,
    };
  } catch (error) {
    return {
      ok: false,
      fetchStatus: 'error',
      fetchError: error.message,
    };
  }
}

async function fetchOfficialWorldMetadataFromPage(worldId, worldUrl) {
  const normalizedWorldId = normalizeWorldId(worldId) || parseWorldIdFromUrl(worldUrl);
  const sourceUrl = worldUrl || buildWorldUrlFromId(normalizedWorldId);

  if (!normalizedWorldId || !sourceUrl) {
    return null;
  }

  try {
    const { response, text } = await fetchTextWithTimeout(sourceUrl);

    if (!response.ok) {
      return {
        ok: false,
        fetchStatus:
          response.status === 401 || response.status === 403
            ? 'unauthorized'
            : response.status === 404
              ? 'not_found'
              : 'error',
        fetchError: `HTTP ${response.status}`,
      };
    }

    const { worldNameOfficial, authorName } = parseWorldTitleAndAuthor(
      extractMetaTagContent(text, ['og:title', 'twitter:title'])
    );

    return {
      ok: true,
      worldId: normalizedWorldId,
      sourceUrl,
      worldNameOfficial,
      worldDescription: normalizeOfficialWorldText(
        extractMetaTagContent(text, ['og:description', 'twitter:description', 'description'])
      ),
      worldTags: [],
      authorId: null,
      authorName,
      releaseStatus: null,
      imageUrl: extractMetaTagContent(text, ['og:image', 'twitter:image']),
      thumbnailImageUrl: extractMetaTagContent(text, ['og:image', 'twitter:image']),
      fetchStatus: 'success',
      fetchError: null,
    };
  } catch (error) {
    return {
      ok: false,
      fetchStatus: 'error',
      fetchError: error.message,
    };
  }
}

async function fetchAndCacheOfficialWorldMetadata(worldId, worldUrl) {
  const normalizedWorldId = normalizeWorldId(worldId) || parseWorldIdFromUrl(worldUrl);
  const sourceUrl = worldUrl || buildWorldUrlFromId(normalizedWorldId);

  if (!normalizedWorldId) {
    return null;
  }

  const attemptedAt = new Date().toISOString();
  const apiResult = await fetchOfficialWorldMetadataFromApi(normalizedWorldId);

  if (apiResult?.ok) {
    return photoDb.upsertWorldMetadata({
      ...apiResult,
      sourceUrl: sourceUrl || apiResult.sourceUrl,
      fetchedAt: attemptedAt,
      lastAttemptedAt: attemptedAt,
    });
  }

  const pageResult = await fetchOfficialWorldMetadataFromPage(
    normalizedWorldId,
    sourceUrl
  );

  if (pageResult?.ok) {
    return photoDb.upsertWorldMetadata({
      ...pageResult,
      sourceUrl,
      fetchStatus: apiResult?.fetchStatus === 'success' ? 'success' : 'partial',
      fetchError: apiResult?.fetchError || null,
      fetchedAt: attemptedAt,
      lastAttemptedAt: attemptedAt,
    });
  }

  photoDb.upsertWorldMetadata({
    worldId: normalizedWorldId,
    sourceUrl,
    worldTags: [],
    fetchStatus: apiResult?.fetchStatus || pageResult?.fetchStatus || 'error',
    fetchError:
      apiResult?.fetchError ||
      pageResult?.fetchError ||
      'Failed to fetch official world metadata',
    fetchedAt: null,
    lastAttemptedAt: attemptedAt,
  });

  return null;
}

// World reread can optionally use the URL currently typed in the editor so the
// user does not have to save first just to verify official metadata.
async function rereadWorldInfoFromPhotoId(photoId, options = {}) {
  let row = photoDb.getPhotoById(photoId);

  let pendingWorldUrl = null;
  let pendingWorldId = null;

  if (typeof options.worldUrl === 'string' && options.worldUrl.trim().length > 0) {
    const normalizedUrl = normalizeManualWorldUrl(options.worldUrl, row?.world_url);
    pendingWorldUrl = normalizedUrl.worldUrl;
    pendingWorldId = normalizedUrl.worldId;
  }

  if (!row) {
    throw new Error('対象の写真が見つかりません');
  }

  let localWorldInfo = {
    worldId: pendingWorldId || row.world_id,
    worldName: row.world_name,
    worldUrl: pendingWorldUrl || row.world_url,
  };

  try {
    const accessiblePhoto = await ensureAccessiblePhotoFile({
      photoId,
      filePath: row.file_path,
    });

    if (accessiblePhoto?.photo) {
      row = photoDb.getPhotoById(photoId) || row;
    }

    const fileBuffer = await fs.readFile(accessiblePhoto.filePath);
    const tags = loadExifTagsSafely(fileBuffer);
    localWorldInfo = extractWorldInfo(tags, fileBuffer);
  } catch {
    if (!pendingWorldId && !pendingWorldUrl && !row.world_id && !row.world_url) {
      throw new Error('ワールド情報を再取得できませんでした');
    }
  }

  const resolvedWorldId =
    pendingWorldId ||
    localWorldInfo.worldId ||
    row.world_id ||
    parseWorldIdFromUrl(localWorldInfo.worldUrl) ||
    parseWorldIdFromUrl(row.world_url);
  const resolvedWorldUrl =
    pendingWorldUrl ||
    localWorldInfo.worldUrl ||
    row.world_url ||
    buildWorldUrlFromId(resolvedWorldId);
  const officialMetadata = await fetchAndCacheOfficialWorldMetadata(
    resolvedWorldId,
    resolvedWorldUrl
  );
  const resolvedWorldName =
    normalizeOfficialWorldText(officialMetadata?.world_name_official) ||
    localWorldInfo.worldName ||
    row.world_name;

  const saved = photoDb.updateAutoWorldInfo(photoId, {
    worldId: resolvedWorldId,
    worldName: resolvedWorldName,
    worldUrl: resolvedWorldUrl,
  });

  return saved;
}

// Legacy wrappers retained only as encoding-safe quarantine blocks.
// Active IPC handlers use the clean wrappers defined after these functions.
async function openLocalFileOnDiskLegacy(filePath) {
  // Legacy wrapper kept for compatibility with older call sites. The shared
  // payload-based helper is now the canonical path recovery flow.
  return openRecoveredLocalFileOnDisk(filePath);

  /*

  const resolvedPhoto = await ensureAccessiblePhotoFile(filePath);
  filePath = resolvedPhoto.filePath;

  if (typeof filePath !== 'string' || filePath.trim().length === 0) {
    throw new Error('ファイルパスが不正です');
  }

  await fs.access(filePath);

  const openResult = await shell.openPath(filePath);

  if (openResult) {
    throw new Error(openResult);
  }

  return createResolvedPhotoAccessResult(resolvedPhoto);
  */
}

async function openContainingFolderOnDiskLegacy(filePath) {
  // Legacy wrapper kept for compatibility with older call sites. The shared
  // payload-based helper is now the canonical path recovery flow.
  return openRecoveredContainingFolderOnDisk(filePath);

  /*

  const resolvedPhoto = await ensureAccessiblePhotoFile(filePath);
  filePath = resolvedPhoto.filePath;

  if (typeof filePath !== 'string' || filePath.trim().length === 0) {
    throw new Error('ファイルパスが不正です');
  }

  await fs.access(filePath);
  shell.showItemInFolder(filePath);

  return createResolvedPhotoAccessResult(resolvedPhoto);
  */
}

async function openLocalFileOnDiskActive(filePath) {
  return openRecoveredLocalFileOnDisk(filePath);
}

async function openContainingFolderOnDiskActive(filePath) {
  return openRecoveredContainingFolderOnDisk(filePath);
}

// Compatibility wrappers retained so any older local call sites still route
// through the clean payload-aware access helpers.
async function openLocalFileOnDisk(filePath) {
  return openLocalFileOnDiskActive(filePath);
}

async function openContainingFolderOnDisk(filePath) {
  return openContainingFolderOnDiskActive(filePath);
}

async function openRecoveredLocalFileOnDisk(payload) {
  return withAccessiblePhotoFile(payload, async (resolvedFilePath) => {
    const openResult = await shell.openPath(resolvedFilePath);

    if (openResult) {
      throw new Error(openResult);
    }
  });
}

async function openRecoveredContainingFolderOnDisk(payload) {
  return withAccessiblePhotoFile(payload, async (resolvedFilePath) => {
    shell.showItemInFolder(resolvedFilePath);
  });
}

function createProcessingProgressReporter(webContents, operation) {
  return (payload = {}) => {
    if (!webContents || webContents.isDestroyed()) {
      return;
    }

    webContents.send(PROCESSING_PROGRESS_CHANNEL, {
      operation,
      ...payload,
    });
  };
}

function createWindow() {
  const windowIconPath = fsSync.existsSync(APP_WINDOW_ICON_ICO_PATH)
    ? APP_WINDOW_ICON_ICO_PATH
    : APP_WINDOW_ICON_PNG_PATH;

  const mainWindow = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1200,
    minHeight: 760,
    autoHideMenuBar: true,
    title: APP_TITLE,
    icon: windowIconPath,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  // Hide and detach the stock menu bar so Windows doesn't keep a visible shell.
  mainWindow.setMenuBarVisibility(false);
  mainWindow.autoHideMenuBar = true;
  if (typeof mainWindow.removeMenu === 'function') {
    mainWindow.removeMenu();
  }
  mainWindow.loadFile(path.join(__dirname, 'index.html'));
  mainWindowRef = mainWindow;
  setupAutoUpdater();
  scheduleStartupAutoUpdateCheck(mainWindow);
  mainWindow.on('closed', () => {
    if (mainWindowRef === mainWindow) {
      mainWindowRef = null;
    }
  });
  return mainWindow;
}

app.whenReady().then(async () => {
  // Remove the stock Electron menu so the app ships with a cleaner surface.
  Menu.setApplicationMenu(null);

  const dbDirPath = path.join(app.getPath('userData'), 'data');
  const appStorageRootPath = getThumbnailStorageRootPath(APP_STORAGE_DIRNAME);

  await migrateLegacyThumbnailStorage();

  thumbnailDirPath = path.join(appStorageRootPath, THUMBNAIL_DIRNAME);
  preferencesFilePath = path.join(dbDirPath, 'preferences.json');

  await ensureDir(dbDirPath);
  await ensureDir(appStorageRootPath);
  await ensureDir(thumbnailDirPath);

  photoDb = initDatabase(path.join(dbDirPath, 'app.sqlite'));
  await reconcileManagedThumbnailPaths();

  ipcMain.handle('import-images', async (event) => {
    const result = await dialog.showOpenDialog({
      title: '画像を複数選択',
      properties: ['openFile', 'multiSelections'],
      filters: [
        { name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'webp'] },
      ],
    });
  
    if (result.canceled || result.filePaths.length === 0) {
      return { canceled: true };
    }

    const progressReporter = createProcessingProgressReporter(
      event.sender,
      'import-images'
    );

    return importManyFiles(result.filePaths, progressReporter);
  });
  
  ipcMain.handle('import-folder', async (event) => {
    const result = await dialog.showOpenDialog({
      title: '取り込むフォルダを選択',
      properties: ['openDirectory'],
    });
  
    if (result.canceled || result.filePaths.length === 0) {
      return { canceled: true };
    }
  
    const folderPath = result.filePaths[0];
    const progressReporter = createProcessingProgressReporter(
      event.sender,
      'import-folder'
    );
    registerTrackedFolders([folderPath]);
    const filePaths = await collectImageFilesFromDirectory(folderPath);
  
    if (filePaths.length === 0) {
      return {
        canceled: false,
        totalSelected: 0,
        importedCount: 0,
        newCount: 0,
        updatedCount: 0,
        failedCount: 0,
        failedFiles: [],
        selectedMonth: null,
        emptyFolder: true,
      };
    }
  
    return importManyFiles(filePaths, progressReporter);
  });

  ipcMain.handle('import-dropped-paths', async (event, droppedPaths) => {
    const progressReporter = createProcessingProgressReporter(
      event.sender,
      'import-dropped-paths'
    );
    const { filePaths: importablePaths, trackedFolderPaths } =
      await expandDroppedPathsToImportTargets(
      Array.isArray(droppedPaths) ? droppedPaths : []
      );

    registerTrackedFolders(trackedFolderPaths);

    if (importablePaths.length === 0) {
      return {
        canceled: false,
        totalSelected: 0,
        importedCount: 0,
        newCount: 0,
        updatedCount: 0,
        failedCount: 0,
        failedFiles: [],
        selectedMonth: null,
        emptyDrop: true,
      };
    }

    return importManyFiles(importablePaths, progressReporter);
  });

  ipcMain.handle('regenerate-thumbnails', async (event, payload) => {
    try {
      const progressReporter = createProcessingProgressReporter(
        event.sender,
        'regenerate-thumbnails'
      );

      const targetMonth =
        Number.isInteger(payload?.year) && Number.isInteger(payload?.month)
          ? {
              year: payload.year,
              month: payload.month,
            }
          : null;

      return await regenerateManagedThumbnails(targetMonth, progressReporter);
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        totalCount: 0,
        regeneratedCount: 0,
        skippedCount: 0,
        failedCount: 0,
        failedFiles: [],
      };
    }
  });

  ipcMain.handle('regenerate-missing-thumbnails', async (event) => {
    try {
      const progressReporter = createProcessingProgressReporter(
        event.sender,
        'regenerate-missing-thumbnails'
      );

      return await regenerateMissingThumbnails(progressReporter);
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        totalCount: 0,
        regeneratedCount: 0,
        skippedCount: 0,
        failedCount: 0,
        failedFiles: [],
      };
    }
  });

  ipcMain.handle('clear-thumbnail-cache', async (event) => {
    try {
      const progressReporter = createProcessingProgressReporter(
        event.sender,
        'clear-thumbnail-cache'
      );

      return await clearManagedThumbnailCache(progressReporter);
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        totalCount: 0,
        clearedCount: 0,
        failedCount: 0,
        failedFiles: [],
      };
    }
  });

  ipcMain.handle('reset-database', async (event) => {
    try {
      const progressReporter = createProcessingProgressReporter(
        event.sender,
        'reset-database'
      );
      const automaticBackup = await createAutomaticAppDataBackupFile(
        'before-reset'
      );
      const resetResult = await resetApplicationData(progressReporter);

      return {
        ...resetResult,
        automaticBackupFilePath: automaticBackup.filePath,
        automaticBackupFileName: automaticBackup.fileName,
      };
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        photoCount: 0,
        trackedFolderCount: 0,
        worldCacheCount: 0,
        tagCount: 0,
      };
    }
  });

  ipcMain.handle('check-app-data-health', async () => {
    try {
      return await checkApplicationDataHealth();
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('get-world-metadata-issue-photos', async () => {
    try {
      return await getWorldMetadataIssuePhotos();
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        photoCount: 0,
        photos: [],
      };
    }
  });

  ipcMain.handle('get-health-issue-photos', async (_event, payload) => {
    try {
      return await getHealthIssuePhotos(payload?.issueKind);
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        photoCount: 0,
        photos: [],
      };
    }
  });

  ipcMain.handle('refresh-world-metadata-issues', async (event) => {
    try {
      return await refreshWorldMetadataIssuePhotos(event.sender);
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        photoCount: 0,
        targetCount: 0,
        queuedCount: 0,
        pendingCount: getQueuedWorldMetadataSyncCount(),
      };
    }
  });

  ipcMain.handle('create-app-data-backup', async () => {
    try {
      return await createAppDataBackupFile();
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('restore-app-data-backup', async () => {
    try {
      return await restoreAppDataBackupFile();
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('export-photo-catalog', async (_event, payload) => {
    try {
      return await exportPhotoCatalogFile(payload?.format);
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        format: payload?.format === 'json' ? 'json' : 'csv',
      };
    }
  });

  ipcMain.handle('uninstall-app', async () => {
    return startInternalUninstall({ deleteData: false });
  });

  ipcMain.handle('uninstall-app-and-delete-data', async () => {
    return startInternalUninstall({ deleteData: true });
  });

  ipcMain.handle('get-application-data-summary', async () => {
    return photoDb.getApplicationDataSummary();
  });

  ipcMain.handle('get-sidebar-data', async () => {
    return photoDb.getSidebarTree();
  });

  ipcMain.handle('get-world-sidebar-data', async (_event, sortMode = 'count') => {
    return buildWorldSidebarData(sortMode);
  });

  ipcMain.handle('get-latest-month', async () => {
    return photoDb.getLatestMonth();
  });

  ipcMain.handle('get-photos-by-month', async (_event, year, month) => {
    return buildRendererPhotosFromRows(
      () => photoDb.getPhotosByMonth(year, month),
      'month',
      { year, month }
    );
  });

  ipcMain.handle('get-photos-by-year', async (_event, year) => {
    return buildRendererPhotosFromRows(
      () => photoDb.getPhotosByYear(year),
      'year',
      { year }
    );
  });

  ipcMain.handle('get-photos-by-world-selection', async (_event, selection) => {
    const normalizedWorldId = normalizeWorldId(selection?.worldId);
    const normalizedWorldName =
      normalizeDisplayWorldName(selection?.worldName) ||
      UNKNOWN_WORLD_DISPLAY_NAME;

    return buildRendererPhotosFromRows(
      () =>
        normalizedWorldId
          ? photoDb.getPhotosByWorldId(normalizedWorldId)
          : photoDb.getPhotosByWorldName(
              normalizedWorldName,
              UNKNOWN_WORLD_DISPLAY_NAME
            ),
      'world',
      {
        hasWorldId: Boolean(normalizedWorldId),
        worldKey: selection?.worldKey || '',
      }
    );
  });

  ipcMain.handle('get-world-metadata', async (_event, worldId) => {
    const normalizedWorldId = normalizeWorldId(worldId);

    if (!normalizedWorldId) {
      return null;
    }

    return toRendererWorldMetadata(
      photoDb.getWorldMetadataByWorldId(normalizedWorldId)
    );
  });

  ipcMain.handle('start-world-metadata-sync', async (event, payload) => {
    try {
      return startWorldMetadataSyncForSubscriber(
        payload?.targets,
        event.sender
      );
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        queuedCount: 0,
        pendingCount: pendingWorldMetadataSyncTargets.size,
      };
    }
  });

  ipcMain.handle('get-label-catalog', async () => {
    return photoDb.getAllTags().map(toRendererPhotoLabel).filter(Boolean);
  });

  ipcMain.handle('get-photo-labels', async (_event, photoId) => {
    if (!Number.isInteger(photoId) || photoId <= 0) {
      return [];
    }

    return photoDb.getPhotoTags(photoId).map(toRendererPhotoLabel).filter(Boolean);
  });

  ipcMain.handle('replace-photo-labels', async (_event, payload) => {
    try {
      const photoId = payload?.photoId;

      if (!Number.isInteger(photoId) || photoId <= 0) {
        throw new Error('photoId is required');
      }

      if (!photoDb.getPhotoById(photoId)) {
        throw new Error('対象の写真が見つかりません');
      }

      const normalizedLabels = normalizePhotoLabelPayload(payload?.labels);
      const savedLabels = photoDb.replacePhotoTags(photoId, normalizedLabels);

      return {
        ok: true,
        photoId,
        labels: savedLabels.map(toRendererPhotoLabel).filter(Boolean),
        catalog: photoDb.getAllTags().map(toRendererPhotoLabel).filter(Boolean),
      };
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        labels: [],
        catalog: photoDb.getAllTags().map(toRendererPhotoLabel).filter(Boolean),
      };
    }
  });

  ipcMain.handle('get-tracked-folders', async () => {
    return photoDb.getTrackedFolders();
  });

  ipcMain.handle('get-background-image-preference', async () => {
    return {
      ok: true,
      filePath: await getBackgroundImagePreference(),
    };
  });

  ipcMain.handle('set-background-image-preference', async (_event, payload) => {
    return {
      ok: true,
      filePath: await setBackgroundImagePreference(payload?.filePath),
    };
  });

  ipcMain.handle('select-background-image', async () => {
    const result = await dialog.showOpenDialog({
      title: '背景画像を選択',
      properties: ['openFile'],
      filters: [
        {
          name: 'Image Files',
          extensions: ['png', 'jpg', 'jpeg', 'webp', 'bmp', 'gif'],
        },
      ],
    });

    if (result.canceled || result.filePaths.length === 0) {
      return {
        ok: true,
        canceled: true,
        filePath: '',
      };
    }

    return {
      ok: true,
      canceled: false,
      filePath: result.filePaths[0],
    };
  });

  ipcMain.handle('add-tracked-folder', async () => {
    const result = await dialog.showOpenDialog({
      title: '更新対象フォルダを選択',
      properties: ['openDirectory'],
    });

    if (result.canceled || result.filePaths.length === 0) {
      return {
        ok: true,
        canceled: true,
        folders: photoDb.getTrackedFolders(),
      };
    }

    const folderPath = normalizeTrackedFolderPath(result.filePaths[0]);

    if (!folderPath) {
      return {
        ok: false,
        message: 'フォルダの登録に失敗しました',
        folders: photoDb.getTrackedFolders(),
      };
    }

    registerTrackedFolders([folderPath]);

    return {
      ok: true,
      canceled: false,
      folder: photoDb.getTrackedFolders().find((row) => row.folder_path === folderPath) || null,
      folders: photoDb.getTrackedFolders(),
    };
  });

  ipcMain.handle('remove-tracked-folder', async (_event, folderPath) => {
    const normalizedFolderPath = normalizeTrackedFolderPath(folderPath);

    if (!normalizedFolderPath) {
      return {
        ok: false,
        message: 'フォルダパスが不正です',
        folders: photoDb.getTrackedFolders(),
      };
    }

    const deleted = photoDb.deleteTrackedFolder(normalizedFolderPath);

    return {
      ok: true,
      deleted,
      folders: photoDb.getTrackedFolders(),
    };
  });

  ipcMain.handle('refresh-tracked-folders', async (event, folderPaths) => {
    try {
      const progressReporter = createProcessingProgressReporter(
        event.sender,
        'refresh-tracked-folders'
      );

      return await refreshTrackedFolderImports(folderPaths, progressReporter);
    } catch (error) {
      return createTrackedFolderRefreshResult({
        ok: false,
        message: error.message,
      });
    }
  });

  ipcMain.handle('reimport-registered-photos', async (event, payload) => {
    try {
      const progressReporter = createProcessingProgressReporter(
        event.sender,
        'reimport-registered-photos'
      );

      return await reimportRegisteredPhotos(payload, progressReporter);
    } catch (error) {
      return createRegisteredPhotoReimportResult({
        ok: false,
        message: error.message,
      });
    }
  });

  ipcMain.handle('update-world-name', async (_event, payload) => {
    try {
      const saved = photoDb.updateManualWorldName(
        payload.photoId,
        payload.worldNameManual
      );

      return {
        ok: true,
        photo: toRendererPhoto(saved),
      };
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('update-world-settings', async (_event, payload) => {
    try {
      const currentRow = photoDb.getPhotoById(payload.photoId);

      if (!currentRow) {
        return {
          ok: false,
          message: '対象の写真が見つかりませんでした',
        };
      }

      const normalizedWorld = normalizeManualWorldUrl(
        payload.worldUrl,
        currentRow.world_url
      );

      const saved = photoDb.updateManualWorldSettings(payload.photoId, {
        worldNameManual: payload.worldNameManual,
        worldId:
          normalizedWorld.worldId ||
          currentRow.world_id ||
          null,
        worldUrl:
          normalizedWorld.worldUrl ||
          currentRow.world_url ||
          null,
      });

      return {
        ok: true,
        photo: toRendererPhoto(saved),
      };
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('update-photo-memo', async (_event, payload) => {
    try {
      const saved = photoDb.updatePhotoMemo(
        payload.photoId,
        payload.memoText
      );

      if (!saved) {
        return {
          ok: false,
          message: '対象の写真が見つかりませんでした',
        };
      }

      return {
        ok: true,
        photo: toRendererPhoto(saved),
      };
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('reread-world-name', async (_event, payload) => {
    try {
      const saved = await rereadWorldInfoFromPhotoId(payload.photoId, {
        worldUrl: payload?.worldUrl,
      });
  
      return {
        ok: true,
        photo: toRendererPhoto(saved),
      };
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });
  
  ipcMain.handle('open-external-url', async (_event, url) => {
    try {
      if (typeof url !== 'string' || !/^https?:\/\//i.test(url)) {
        throw new Error('無効なURLです');
      }
  
      await shell.openExternal(url);
  
      return { ok: true };
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('resolve-photo-access', async (_event, payload) => {
    try {
      const resolvedPhoto = await ensureAccessiblePhotoFile(payload);

      return createResolvedPhotoAccessResult(resolvedPhoto, {
        includeFilePath: true,
      });
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('open-local-file', async (_event, filePath) => {
    try {
      return await openLocalFileOnDiskActive(filePath);
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('open-containing-folder', async (_event, filePath) => {
    try {
      return await openContainingFolderOnDiskActive(filePath);
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('save-edited-photo', async (_event, payload) => {
    try {
      return await saveEditedPhotoFile(payload);
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('delete-photo', async (_event, { photoId }) => {
    try {
      return await deletePhotoRegistration(photoId);
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('delete-photos', async (_event, { photoIds }) => {
    try {
      return await deletePhotoRegistrations(photoIds);
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        deletedPhotoIds: [],
        deletedCount: 0,
        failed: [],
        failedCount: 0,
      };
    }
  });

  ipcMain.handle('delete-photos-by-month', async (_event, payload = {}) => {
    try {
      return await deletePhotoRegistrationsByMonth(payload.year, payload.month);
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        deletedPhotoIds: [],
        deletedCount: 0,
        failed: [],
        failedCount: 0,
      };
    }
  });

  ipcMain.handle('delete-all-photos', async () => {
    try {
      return await deleteAllPhotoRegistrations();
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        deletedPhotoIds: [],
        deletedCount: 0,
        failed: [],
        failedCount: 0,
      };
    }
  });

  ipcMain.handle('update-favorite-status', async (_event, payload) => {
    try {
      const updated = photoDb.updateFavoriteStatus(
        payload.photoId,
        payload.isFavorite
      );

      if (!updated) {
        return {
          ok: false,
          message: '対象の写真が見つかりませんでした',
        };
      }

      return {
        ok: true,
        photo: toRendererPhoto(updated),
      };
    } catch (error) {
      return {
        ok: false,
        message: error.message,
      };
    }
  });

  ipcMain.handle('update-favorite-statuses', async (_event, payload) => {
    try {
      const updatedRows = photoDb.updateFavoriteStatuses(
        payload?.photoIds,
        payload?.isFavorite
      );

      return {
        ok: true,
        photos: updatedRows.map(toRendererPhoto),
      };
    } catch (error) {
      return {
        ok: false,
        message: error.message,
        photos: [],
      };
    }
  });

  ipcMain.handle('start-app-update-download', async () => {
    try {
      if (!latestAvailableAppUpdateRelease) {
        return {
          ok: false,
          message: '利用可能なアップデートがありません',
        };
      }

      await startAppUpdateDownload(latestAvailableAppUpdateRelease);
      return {
        ok: true,
      };
    } catch (error) {
      return {
        ok: false,
        message: error instanceof Error ? error.message : '不明なエラー',
      };
    }
  });

  ipcMain.handle('install-downloaded-app-update', async () => {
    try {
      autoUpdater.quitAndInstall();
      return {
        ok: true,
      };
    } catch (error) {
      return {
        ok: false,
        message: error instanceof Error ? error.message : '不明なエラー',
      };
    }
  });


  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

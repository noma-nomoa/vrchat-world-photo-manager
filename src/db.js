const Database = require('better-sqlite3');

const SQLITE_VARIABLE_LIMIT = 900;
const DEFAULT_TAG_COLOR_PALETTE = [
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

// ------------------------------
// Schema migration helpers
// ------------------------------
function ensureColumn(db, tableName, columnDefinition) {
  const columnName = columnDefinition.split(' ')[0];
  const columns = db.prepare(`PRAGMA table_info(${tableName})`).all();

  if (!columns.some((column) => column.name === columnName)) {
    db.exec(`ALTER TABLE ${tableName} ADD COLUMN ${columnDefinition}`);
  }
}

function pickDefaultTagColor(normalizedName = '') {
  const seed = typeof normalizedName === 'string' ? normalizedName : '';
  let hash = 0;

  for (let index = 0; index < seed.length; index += 1) {
    hash = (hash * 31 + seed.charCodeAt(index)) >>> 0;
  }

  return DEFAULT_TAG_COLOR_PALETTE[
    hash % DEFAULT_TAG_COLOR_PALETTE.length
  ];
}

function normalizeTagColor(colorHex, normalizedName = '') {
  if (typeof colorHex === 'string') {
    const trimmed = colorHex.trim();
    const match = trimmed.match(/^#?([0-9a-fA-F]{6})$/);

    if (match) {
      return `#${match[1].toUpperCase()}`;
    }
  }

  return pickDefaultTagColor(normalizedName);
}

function normalizeTagName(tagInput) {
  const rawName =
    typeof tagInput === 'string'
      ? tagInput
      : typeof tagInput?.name === 'string'
        ? tagInput.name
        : '';

  if (typeof rawName !== 'string') {
    return null;
  }

  const name = rawName
    .normalize('NFC')
    .replace(/\s+/g, ' ')
    .trim();

  if (name.length === 0) {
    return null;
  }

  return {
    name,
    normalizedName: name.toLowerCase(),
    colorHex: normalizeTagColor(
      typeof tagInput === 'object'
        ? tagInput.colorHex || tagInput.color_hex
        : null,
      name.toLowerCase()
    ),
  };
}

// ------------------------------
// Database initialization / schema
// ------------------------------
function initDatabase(dbPath) {
  const db = new Database(dbPath);

  db.pragma('journal_mode = WAL');
  db.pragma('synchronous = NORMAL');
  db.pragma('temp_store = MEMORY');
  db.pragma('cache_size = -16384');
  db.pragma('busy_timeout = 5000');
  db.pragma('foreign_keys = ON');

  db.exec(`
    CREATE TABLE IF NOT EXISTS photos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      file_path TEXT NOT NULL,
      file_name TEXT NOT NULL,
      file_hash TEXT NOT NULL UNIQUE,
      taken_at TEXT NOT NULL,
      taken_at_timestamp INTEGER NOT NULL,
      group_date TEXT NOT NULL,
      year INTEGER NOT NULL,
      month INTEGER NOT NULL,
      day INTEGER NOT NULL,
      world_id TEXT,
      world_name TEXT,
      world_name_manual TEXT,
      world_url TEXT,
      thumbnail_path TEXT,
      image_width INTEGER,
      image_height INTEGER,
      resolution_tier TEXT,
      orientation_tier TEXT,
      memo_text TEXT,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE INDEX IF NOT EXISTS idx_photos_year_month
      ON photos (year DESC, month DESC);

    CREATE INDEX IF NOT EXISTS idx_photos_group_date
      ON photos (group_date DESC);

    CREATE INDEX IF NOT EXISTS idx_photos_taken_at_timestamp
      ON photos (taken_at_timestamp DESC);

    CREATE TABLE IF NOT EXISTS tracked_folders (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      folder_path TEXT NOT NULL UNIQUE,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL,
      last_imported_at TEXT NOT NULL
    );

    CREATE INDEX IF NOT EXISTS idx_tracked_folders_last_imported_at
      ON tracked_folders (last_imported_at DESC);

    CREATE TABLE IF NOT EXISTS world_metadata_cache (
      world_id TEXT PRIMARY KEY,
      source_url TEXT,
      world_name_official TEXT,
      world_description TEXT,
      world_tags_json TEXT,
      author_id TEXT,
      author_name TEXT,
      release_status TEXT,
      image_url TEXT,
      thumbnail_image_url TEXT,
      fetch_status TEXT NOT NULL,
      fetch_error TEXT,
      fetched_at TEXT,
      last_attempted_at TEXT NOT NULL,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE INDEX IF NOT EXISTS idx_world_metadata_cache_last_attempted_at
      ON world_metadata_cache (last_attempted_at DESC);

    CREATE TABLE IF NOT EXISTS tags (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      normalized_name TEXT NOT NULL UNIQUE,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE INDEX IF NOT EXISTS idx_tags_name
      ON tags (name COLLATE NOCASE ASC);

    CREATE TABLE IF NOT EXISTS photo_tags (
      photo_id INTEGER NOT NULL,
      tag_id INTEGER NOT NULL,
      created_at TEXT NOT NULL,
      PRIMARY KEY (photo_id, tag_id),
      FOREIGN KEY(photo_id) REFERENCES photos(id) ON DELETE CASCADE,
      FOREIGN KEY(tag_id) REFERENCES tags(id) ON DELETE CASCADE
    );

    CREATE INDEX IF NOT EXISTS idx_photo_tags_tag_id
      ON photo_tags (tag_id, photo_id);
  `);

  ensureColumn(db, 'photos', 'world_name_manual TEXT');
  ensureColumn(db, 'photos', 'is_favorite INTEGER NOT NULL DEFAULT 0');
  ensureColumn(db, 'photos', 'image_width INTEGER');
  ensureColumn(db, 'photos', 'image_height INTEGER');
  ensureColumn(db, 'photos', 'resolution_tier TEXT');
  ensureColumn(db, 'photos', 'orientation_tier TEXT');
  ensureColumn(db, 'photos', 'memo_text TEXT');
  ensureColumn(db, 'tags', 'color_hex TEXT');
  db.exec(`
    UPDATE photos
    SET orientation_tier = CASE
      WHEN image_width IS NULL OR image_height IS NULL THEN orientation_tier
      WHEN image_width = image_height THEN 'square'
      WHEN image_width > image_height THEN 'landscape'
      ELSE 'portrait'
    END
    WHERE (orientation_tier IS NULL OR orientation_tier = '')
      AND image_width IS NOT NULL
      AND image_height IS NOT NULL
  `);

  const getPhotoByHashStmt = db.prepare(`
    SELECT *
    FROM photos
    WHERE file_hash = ?
  `);

  const getPhotoByIdStmt = db.prepare(`
    SELECT *
    FROM photos
    WHERE id = ?
  `);

  const deletePhotoByIdStmt = db.prepare(`
    DELETE FROM photos
    WHERE id = ?
  `);

  const insertOrUpdatePhotoStmt = db.prepare(`
    INSERT INTO photos (
      file_path,
      file_name,
      file_hash,
      taken_at,
      taken_at_timestamp,
      group_date,
      year,
      month,
      day,
      world_id,
      world_name,
      world_name_manual,
      world_url,
      thumbnail_path,
      image_width,
      image_height,
      resolution_tier,
      orientation_tier,
      memo_text,
      created_at,
      updated_at
    ) VALUES (
      @filePath,
      @fileName,
      @fileHash,
      @takenAt,
      @takenAtTimestamp,
      @groupDate,
      @year,
      @month,
      @day,
      @worldId,
      @worldName,
      @worldNameManual,
      @worldUrl,
      @thumbnailPath,
      @imageWidth,
      @imageHeight,
      @resolutionTier,
      @orientationTier,
      @memoText,
      @createdAt,
      @updatedAt
    )
    ON CONFLICT(file_hash) DO UPDATE SET
      file_path = excluded.file_path,
      file_name = excluded.file_name,
      taken_at = excluded.taken_at,
      taken_at_timestamp = excluded.taken_at_timestamp,
      group_date = excluded.group_date,
      year = excluded.year,
      month = excluded.month,
      day = excluded.day,
      world_id = excluded.world_id,
      world_name = excluded.world_name,
      world_url = excluded.world_url,
      thumbnail_path = CASE
        WHEN excluded.thumbnail_path IS NULL OR excluded.thumbnail_path = ''
          THEN thumbnail_path
        ELSE excluded.thumbnail_path
      END,
      image_width = excluded.image_width,
      image_height = excluded.image_height,
      resolution_tier = excluded.resolution_tier,
      orientation_tier = excluded.orientation_tier,
      memo_text = CASE
        WHEN excluded.memo_text IS NULL THEN memo_text
        ELSE excluded.memo_text
      END,
      updated_at = excluded.updated_at
  `);

  const getSidebarRowsStmt = db.prepare(`
    SELECT
      year,
      month,
      COUNT(*) AS count
    FROM photos
    GROUP BY year, month
    ORDER BY year DESC, month DESC
  `);

  const getLatestMonthStmt = db.prepare(`
    SELECT year, month
    FROM photos
    ORDER BY year DESC, month DESC
    LIMIT 1
  `);

  const getPhotosByMonthStmt = db.prepare(`
    SELECT *
    FROM photos
    WHERE year = ? AND month = ?
    ORDER BY taken_at_timestamp DESC, id DESC
  `);

  const getPhotosByYearStmt = db.prepare(`
    SELECT *
    FROM photos
    WHERE year = ?
    ORDER BY taken_at_timestamp DESC, id DESC
  `);

  const getAllPhotosStmt = db.prepare(`
    SELECT id, file_path, file_hash, thumbnail_path, year, month
    FROM photos
    ORDER BY id ASC
  `);

  const getTrackedFolderByPathStmt = db.prepare(`
    SELECT *
    FROM tracked_folders
    WHERE folder_path = ?
  `);

  const getTrackedFoldersStmt = db.prepare(`
    SELECT *
    FROM tracked_folders
    ORDER BY folder_path COLLATE NOCASE ASC
  `);

  const deleteTrackedFolderStmt = db.prepare(`
    DELETE FROM tracked_folders
    WHERE folder_path = ?
  `);

  const upsertTrackedFolderStmt = db.prepare(`
    INSERT INTO tracked_folders (
      folder_path,
      created_at,
      updated_at,
      last_imported_at
    ) VALUES (?, ?, ?, ?)
    ON CONFLICT(folder_path) DO UPDATE SET
      updated_at = excluded.updated_at,
      last_imported_at = excluded.last_imported_at
  `);

  const getWorldMetadataByWorldIdStmt = db.prepare(`
    SELECT *
    FROM world_metadata_cache
    WHERE world_id = ?
  `);

  const upsertWorldMetadataStmt = db.prepare(`
    INSERT INTO world_metadata_cache (
      world_id,
      source_url,
      world_name_official,
      world_description,
      world_tags_json,
      author_id,
      author_name,
      release_status,
      image_url,
      thumbnail_image_url,
      fetch_status,
      fetch_error,
      fetched_at,
      last_attempted_at,
      created_at,
      updated_at
    ) VALUES (
      @worldId,
      @sourceUrl,
      @worldNameOfficial,
      @worldDescription,
      @worldTagsJson,
      @authorId,
      @authorName,
      @releaseStatus,
      @imageUrl,
      @thumbnailImageUrl,
      @fetchStatus,
      @fetchError,
      @fetchedAt,
      @lastAttemptedAt,
      @createdAt,
      @updatedAt
    )
    ON CONFLICT(world_id) DO UPDATE SET
      source_url = excluded.source_url,
      world_name_official = excluded.world_name_official,
      world_description = excluded.world_description,
      world_tags_json = excluded.world_tags_json,
      author_id = excluded.author_id,
      author_name = excluded.author_name,
      release_status = excluded.release_status,
      image_url = excluded.image_url,
      thumbnail_image_url = excluded.thumbnail_image_url,
      fetch_status = excluded.fetch_status,
      fetch_error = excluded.fetch_error,
      fetched_at = excluded.fetched_at,
      last_attempted_at = excluded.last_attempted_at,
      updated_at = excluded.updated_at
  `);

  const getTagByNormalizedNameStmt = db.prepare(`
    SELECT *
    FROM tags
    WHERE normalized_name = ?
  `);

  const getAllTagsStmt = db.prepare(`
    SELECT
      t.id,
      t.name,
      t.normalized_name,
      t.color_hex,
      COUNT(pt.photo_id) AS photo_count
    FROM tags t
    LEFT JOIN photo_tags pt
      ON pt.tag_id = t.id
    GROUP BY t.id
    ORDER BY t.name COLLATE NOCASE ASC, t.id ASC
  `);

  const getPhotoTagsStmt = db.prepare(`
    SELECT
      t.id,
      t.name,
      t.normalized_name,
      t.color_hex
    FROM photo_tags pt
    INNER JOIN tags t
      ON t.id = pt.tag_id
    WHERE pt.photo_id = ?
    ORDER BY t.name COLLATE NOCASE ASC, t.id ASC
  `);

  const upsertTagStmt = db.prepare(`
    INSERT INTO tags (
      name,
      normalized_name,
      color_hex,
      created_at,
      updated_at
    ) VALUES (?, ?, ?, ?, ?)
    ON CONFLICT(normalized_name) DO UPDATE SET
      name = excluded.name,
      color_hex = excluded.color_hex,
      updated_at = excluded.updated_at
  `);

  const clearPhotoTagsStmt = db.prepare(`
    DELETE FROM photo_tags
    WHERE photo_id = ?
  `);

  const insertPhotoTagStmt = db.prepare(`
    INSERT OR IGNORE INTO photo_tags (
      photo_id,
      tag_id,
      created_at
    ) VALUES (?, ?, ?)
  `);

  const deleteUnusedTagsStmt = db.prepare(`
    DELETE FROM tags
    WHERE id NOT IN (
      SELECT DISTINCT tag_id
      FROM photo_tags
    )
  `);

  const updateThumbnailPathStmt = db.prepare(`
    UPDATE photos
    SET
      thumbnail_path = ?,
      updated_at = ?
    WHERE id = ?
  `);

  const updateManualWorldNameStmt = db.prepare(`
    UPDATE photos
    SET
      world_name_manual = ?,
      updated_at = ?
    WHERE id = ?
  `);

  const updateManualWorldSettingsStmt = db.prepare(`
    UPDATE photos
    SET
      world_name_manual = ?,
      world_id = ?,
      world_url = ?,
      updated_at = ?
    WHERE id = ?
  `);

  const updateFavoriteStatusStmt = db.prepare(`
    UPDATE photos
    SET
      is_favorite = ?,
      updated_at = ?
    WHERE id = ?
  `);

  const updateImageMetadataStmt = db.prepare(`
    UPDATE photos
    SET
      image_width = ?,
      image_height = ?,
      resolution_tier = ?,
      orientation_tier = ?,
      updated_at = ?
    WHERE id = ?
  `);

  const updatePhotoFileLocationStmt = db.prepare(`
    UPDATE photos
    SET
      file_path = ?,
      file_name = ?,
      updated_at = ?
    WHERE id = ?
  `);

  const updatePhotoMemoStmt = db.prepare(`
    UPDATE photos
    SET
      memo_text = ?,
      updated_at = ?
    WHERE id = ?
  `);

  const deleteAllPhotosStmt = db.prepare(`
    DELETE FROM photos
  `);

  const deleteAllTrackedFoldersStmt = db.prepare(`
    DELETE FROM tracked_folders
  `);

  const deleteAllWorldMetadataCacheStmt = db.prepare(`
    DELETE FROM world_metadata_cache
  `);

  const deleteAllTagsStmt = db.prepare(`
    DELETE FROM tags
  `);

  const resetSqliteSequenceStmt = db.prepare(`
    DELETE FROM sqlite_sequence
    WHERE name IN ('photos', 'tracked_folders', 'tags')
  `);

  const replacePhotoTagsTxn = db.transaction((photoId, tagNames) => {
    const normalizedTags = Array.from(
      new Map(
        (Array.isArray(tagNames) ? tagNames : [])
          .map(normalizeTagName)
          .filter(Boolean)
          .map((tag) => [tag.normalizedName, tag])
      ).values()
    );

    clearPhotoTagsStmt.run(photoId);

    if (normalizedTags.length === 0) {
      deleteUnusedTagsStmt.run();
      return;
    }

    for (const tag of normalizedTags) {
      const nowIso = new Date().toISOString();

      upsertTagStmt.run(
        tag.name,
        tag.normalizedName,
        tag.colorHex,
        nowIso,
        nowIso
      );

      const savedTag = getTagByNormalizedNameStmt.get(tag.normalizedName);
      insertPhotoTagStmt.run(photoId, savedTag.id, nowIso);
    }

    deleteUnusedTagsStmt.run();
  });

  function updateAutoWorldInfo(photoId, payload) {
    db.prepare(`
      UPDATE photos
      SET
        world_id = ?,
        world_name = ?,
        world_name_manual = NULL,
        world_url = ?,
        updated_at = ?
      WHERE id = ?
    `).run(
      payload.worldId,
      payload.worldName,
      payload.worldUrl,
      new Date().toISOString(),
      photoId
    );

    return getPhotoByIdStmt.get(photoId) || null;
  }

  function insertOrUpdatePhoto(photo) {
    insertOrUpdatePhotoStmt.run(photo);
    return getPhotoByHashStmt.get(photo.fileHash);
  }

  const insertOrUpdatePhotosTxn = db.transaction((photos) => {
    for (const photo of photos) {
      insertOrUpdatePhotoStmt.run(photo);
    }
  });

  function getPhotoByHash(fileHash) {
    return getPhotoByHashStmt.get(fileHash) || null;
  }

  function getExistingPhotoHashes(fileHashes) {
    const normalizedHashes = Array.from(
      new Set(
        (Array.isArray(fileHashes) ? fileHashes : []).filter(
          (fileHash) => typeof fileHash === 'string' && fileHash.trim().length > 0
        )
      )
    );

    if (normalizedHashes.length === 0) {
      return [];
    }

    const rows = [];

    for (
      let startIndex = 0;
      startIndex < normalizedHashes.length;
      startIndex += SQLITE_VARIABLE_LIMIT
    ) {
      const chunk = normalizedHashes.slice(
        startIndex,
        startIndex + SQLITE_VARIABLE_LIMIT
      );
      const placeholders = chunk.map(() => '?').join(', ');
      const stmt = db.prepare(`
        SELECT file_hash
        FROM photos
        WHERE file_hash IN (${placeholders})
      `);

      rows.push(...stmt.all(...chunk));
    }

    return rows;
  }

  function getPhotosByHashes(fileHashes) {
    const normalizedHashes = Array.from(
      new Set(
        (Array.isArray(fileHashes) ? fileHashes : []).filter(
          (fileHash) => typeof fileHash === 'string' && fileHash.trim().length > 0
        )
      )
    );

    if (normalizedHashes.length === 0) {
      return [];
    }

    const rows = [];

    for (
      let startIndex = 0;
      startIndex < normalizedHashes.length;
      startIndex += SQLITE_VARIABLE_LIMIT
    ) {
      const chunk = normalizedHashes.slice(
        startIndex,
        startIndex + SQLITE_VARIABLE_LIMIT
      );
      const placeholders = chunk.map(() => '?').join(', ');
      const stmt = db.prepare(`
        SELECT *
        FROM photos
        WHERE file_hash IN (${placeholders})
      `);

      rows.push(...stmt.all(...chunk));
    }

    return rows;
  }

  function insertOrUpdatePhotos(photos) {
    const normalizedPhotos = Array.isArray(photos)
      ? photos.filter(Boolean)
      : [];

    if (normalizedPhotos.length === 0) {
      return 0;
    }

    insertOrUpdatePhotosTxn(normalizedPhotos);
    return normalizedPhotos.length;
  }

  function getPhotoById(photoId) {
    return getPhotoByIdStmt.get(photoId) || null;
  }

  function deletePhotoById(photoId) {
    const existing = getPhotoByIdStmt.get(photoId);

    if (!existing) {
      return null;
    }

    deletePhotoByIdStmt.run(photoId);
    return existing;
  }

  function getSidebarTree() {
    const rows = getSidebarRowsStmt.all();
    const yearMap = new Map();

    for (const row of rows) {
      if (!yearMap.has(row.year)) {
        yearMap.set(row.year, {
          year: row.year,
          totalCount: 0,
          months: [],
        });
      }

      const yearEntry = yearMap.get(row.year);
      yearEntry.totalCount += row.count;
      yearEntry.months.push({
        month: row.month,
        count: row.count,
      });
    }

    return Array.from(yearMap.values());
  }

  function getLatestMonth() {
    return getLatestMonthStmt.get() || null;
  }

  function getPhotosByMonth(year, month) {
    return getPhotosByMonthStmt.all(year, month);
  }

  function getPhotosByYear(year) {
    return getPhotosByYearStmt.all(year);
  }

  function getAllPhotos() {
    return getAllPhotosStmt.all();
  }

  function getTrackedFolders() {
    return getTrackedFoldersStmt.all();
  }

  function upsertTrackedFolder(folderPath) {
    const nowIso = new Date().toISOString();

    upsertTrackedFolderStmt.run(folderPath, nowIso, nowIso, nowIso);
    return getTrackedFolderByPathStmt.get(folderPath) || null;
  }

  function deleteTrackedFolder(folderPath) {
    const existing = getTrackedFolderByPathStmt.get(folderPath);

    if (!existing) {
      return null;
    }

    deleteTrackedFolderStmt.run(folderPath);
    return existing;
  }

  function getWorldMetadataByWorldId(worldId) {
    if (typeof worldId !== 'string' || worldId.trim().length === 0) {
      return null;
    }

    return getWorldMetadataByWorldIdStmt.get(worldId.trim()) || null;
  }

  function upsertWorldMetadata(payload) {
    if (!payload || typeof payload.worldId !== 'string' || payload.worldId.trim().length === 0) {
      throw new Error('worldId is required');
    }

    const worldId = payload.worldId.trim();
    const nowIso = new Date().toISOString();
    const existing = getWorldMetadataByWorldId(worldId);

    upsertWorldMetadataStmt.run({
      worldId,
      sourceUrl: payload.sourceUrl || null,
      worldNameOfficial: payload.worldNameOfficial || null,
      worldDescription: payload.worldDescription || null,
      worldTagsJson:
        typeof payload.worldTagsJson === 'string'
          ? payload.worldTagsJson
          : JSON.stringify(Array.isArray(payload.worldTags) ? payload.worldTags : []),
      authorId: payload.authorId || null,
      authorName: payload.authorName || null,
      releaseStatus: payload.releaseStatus || null,
      imageUrl: payload.imageUrl || null,
      thumbnailImageUrl: payload.thumbnailImageUrl || null,
      fetchStatus: payload.fetchStatus || 'success',
      fetchError: payload.fetchError || null,
      fetchedAt: payload.fetchedAt || nowIso,
      lastAttemptedAt: payload.lastAttemptedAt || nowIso,
      createdAt: existing?.created_at || nowIso,
      updatedAt: nowIso,
    });

    return getWorldMetadataByWorldId(worldId);
  }

  function getAllTags() {
    return getAllTagsStmt.all();
  }

  function getPhotoTags(photoId) {
    return getPhotoTagsStmt.all(photoId);
  }

  function getPhotoTagsByPhotoIds(photoIds) {
    const normalizedIds = Array.from(
      new Set(
        (Array.isArray(photoIds) ? photoIds : []).filter(
          (photoId) => Number.isInteger(photoId) && photoId > 0
        )
      )
    );

    if (normalizedIds.length === 0) {
      return [];
    }

    const placeholders = normalizedIds.map(() => '?').join(', ');
    const stmt = db.prepare(`
      SELECT
        pt.photo_id,
        t.id,
        t.name,
        t.normalized_name,
        t.color_hex
      FROM photo_tags pt
      INNER JOIN tags t
        ON t.id = pt.tag_id
      WHERE pt.photo_id IN (${placeholders})
      ORDER BY pt.photo_id ASC, t.name COLLATE NOCASE ASC, t.id ASC
    `);

    return stmt.all(...normalizedIds);
  }

  function replacePhotoTags(photoId, tagNames) {
    replacePhotoTagsTxn(photoId, tagNames);
    return getPhotoTags(photoId);
  }

  function updateThumbnailPath(photoId, thumbnailPath) {
    updateThumbnailPathStmt.run(
      thumbnailPath,
      new Date().toISOString(),
      photoId
    );

    return getPhotoByIdStmt.get(photoId) || null;
  }

  function updateManualWorldName(photoId, worldNameManual) {
    const normalizedValue =
      typeof worldNameManual === 'string' && worldNameManual.trim().length > 0
        ? worldNameManual.trim()
        : null;

    updateManualWorldNameStmt.run(
      normalizedValue,
      new Date().toISOString(),
      photoId
    );

    return getPhotoByIdStmt.get(photoId) || null;
  }

  function updateFavoriteStatus(photoId, isFavorite) {
    updateFavoriteStatusStmt.run(
      isFavorite ? 1 : 0,
      new Date().toISOString(),
      photoId
    );

    return getPhotoByIdStmt.get(photoId) || null;
  }

  function updateImageMetadata(photoId, payload) {
    updateImageMetadataStmt.run(
      payload.imageWidth,
      payload.imageHeight,
      payload.resolutionTier,
      payload.orientationTier,
      new Date().toISOString(),
      photoId
    );

    return getPhotoByIdStmt.get(photoId) || null;
  }

  function getPhotosByWorldId(worldId) {
    const normalizedWorldId =
      typeof worldId === 'string' && worldId.trim().length > 0
        ? worldId.trim()
        : null;

    if (!normalizedWorldId) {
      return [];
    }

    return db
      .prepare(
        `
          SELECT *
          FROM photos
          WHERE world_id = ?
          ORDER BY taken_at_timestamp DESC, id DESC
        `
      )
      .all(normalizedWorldId);
  }

  function updateAutoWorldInfoByWorldId(worldId, payload) {
    const normalizedWorldId =
      typeof worldId === 'string' && worldId.trim().length > 0
        ? worldId.trim()
        : null;

    if (!normalizedWorldId) {
      return [];
    }

    db.prepare(`
      UPDATE photos
      SET
        world_name = ?,
        world_url = COALESCE(?, world_url),
        updated_at = ?
      WHERE world_id = ?
        AND (
          world_name_manual IS NULL OR
          TRIM(world_name_manual) = ''
        )
    `).run(
      payload.worldName,
      payload.worldUrl || null,
      new Date().toISOString(),
      normalizedWorldId
    );

    return getPhotosByWorldId(normalizedWorldId);
  }

  function clearThumbnailPaths(photoIds) {
    const normalizedIds = Array.from(
      new Set(
        (Array.isArray(photoIds) ? photoIds : [])
          .map((photoId) => Number.parseInt(photoId, 10))
          .filter((photoId) => Number.isInteger(photoId) && photoId > 0)
      )
    );

    if (normalizedIds.length === 0) {
      return 0;
    }

    const updatedAt = new Date().toISOString();
    const txn = db.transaction((ids) => {
      for (const photoId of ids) {
        updateThumbnailPathStmt.run(null, updatedAt, photoId);
      }
    });

    txn(normalizedIds);
    return normalizedIds.length;
  }

  function updateManualWorldSettings(photoId, payload = {}) {
    const normalizedWorldName =
      typeof payload.worldNameManual === 'string' &&
      payload.worldNameManual.trim().length > 0
        ? payload.worldNameManual.trim()
        : null;
    const normalizedWorldId =
      typeof payload.worldId === 'string' && payload.worldId.trim().length > 0
        ? payload.worldId.trim()
        : null;
    const normalizedWorldUrl =
      typeof payload.worldUrl === 'string' && payload.worldUrl.trim().length > 0
        ? payload.worldUrl.trim()
        : null;

    updateManualWorldSettingsStmt.run(
      normalizedWorldName,
      normalizedWorldId,
      normalizedWorldUrl,
      new Date().toISOString(),
      photoId
    );

    return getPhotoByIdStmt.get(photoId) || null;
  }

  function updateFavoriteStatuses(photoIds, isFavorite) {
    const normalizedIds = Array.from(
      new Set(
        (Array.isArray(photoIds) ? photoIds : []).filter(
          (photoId) => Number.isInteger(photoId) && photoId > 0
        )
      )
    );

    if (normalizedIds.length === 0) {
      return [];
    }

    const nowIso = new Date().toISOString();

    const txn = db.transaction((ids) => {
      for (const photoId of ids) {
        updateFavoriteStatusStmt.run(
          isFavorite ? 1 : 0,
          nowIso,
          photoId
        );
      }
    });

    txn(normalizedIds);

    return normalizedIds
      .map((photoId) => getPhotoByIdStmt.get(photoId) || null)
      .filter(Boolean);
  }

  function updatePhotoFileLocation(photoId, filePath, fileName) {
    updatePhotoFileLocationStmt.run(
      filePath,
      fileName,
      new Date().toISOString(),
      photoId
    );

    return getPhotoByIdStmt.get(photoId) || null;
  }

  function updatePhotoMemo(photoId, memoText) {
    const normalizedMemo =
      typeof memoText === 'string' && memoText.trim().length > 0
        ? memoText.replace(/\r\n/g, '\n').trim()
        : null;

    updatePhotoMemoStmt.run(
      normalizedMemo,
      new Date().toISOString(),
      photoId
    );

    return getPhotoByIdStmt.get(photoId) || null;
  }

  // Settings and maintenance surfaces use this lightweight summary to show
  // the current persisted footprint without triggering heavy scans.
  function getApplicationDataSummary() {
    return {
      photoCount: db.prepare('SELECT COUNT(*) AS count FROM photos').get().count,
      trackedFolderCount: db
        .prepare('SELECT COUNT(*) AS count FROM tracked_folders')
        .get().count,
      worldCacheCount: db
        .prepare('SELECT COUNT(*) AS count FROM world_metadata_cache')
        .get().count,
      tagCount: db.prepare('SELECT COUNT(*) AS count FROM tags').get().count,
    };
  }

  // Reset helpers live in the DB layer so destructive maintenance work has a
  // single place to audit instead of scattering direct DELETEs across the app.
  function resetApplicationData() {
    const counts = getApplicationDataSummary();

    const txn = db.transaction(() => {
      deleteAllPhotosStmt.run();
      deleteAllTrackedFoldersStmt.run();
      deleteAllWorldMetadataCacheStmt.run();
      deleteAllTagsStmt.run();
      resetSqliteSequenceStmt.run();
    });

    txn();
    return counts;
  }

  return {
    insertOrUpdatePhoto,
    insertOrUpdatePhotos,
    getPhotoByHash,
    getExistingPhotoHashes,
    getPhotosByHashes,
    getPhotoById,
    getSidebarTree,
    getLatestMonth,
    getPhotosByMonth,
    getPhotosByYear,
    getAllPhotos,
    getPhotosByWorldId,
    getTrackedFolders,
    upsertTrackedFolder,
    deleteTrackedFolder,
    getWorldMetadataByWorldId,
    upsertWorldMetadata,
    getAllTags,
    getPhotoTags,
    getPhotoTagsByPhotoIds,
    replacePhotoTags,
    updateManualWorldName,
    updateManualWorldSettings,
    updateAutoWorldInfo,
    updateAutoWorldInfoByWorldId,
    updateThumbnailPath,
    clearThumbnailPaths,
    deletePhotoById,
    updateFavoriteStatus,
    updateFavoriteStatuses,
    updateImageMetadata,
    updatePhotoFileLocation,
    updatePhotoMemo,
    getApplicationDataSummary,
    resetApplicationData,
  };
}

module.exports = {
  initDatabase,
};

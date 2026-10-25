'use strict';

const path = require('node:path');

// Cache ownership, migration, encoding, and row recovery are kept behind ports.
function createThumbnailStore({
  fs,
  nativeImage,
  getHomePath,
  getThumbnailDirectoryPath,
  getPhotoDb,
  storageDirname,
  legacyStorageDirname,
  thumbnailDirname = 'thumbnails',
}) {
  const APP_STORAGE_DIRNAME = storageDirname;
  const LEGACY_APP_STORAGE_DIRNAME = legacyStorageDirname;
  const THUMBNAIL_DIRNAME = thumbnailDirname;
  const THUMBNAIL_SIZE = 320;
  const THUMBNAIL_EXTENSION = '.jpg';
  const THUMBNAIL_JPEG_QUALITY = 82;

  async function ensureDir(dirPath) {
    await fs.mkdir(dirPath, { recursive: true });
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

  function getThumbnailStorageRootPath(dirname = APP_STORAGE_DIRNAME) {
    return path.join(getHomePath(), dirname);
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
    if (!getPhotoDb()?.getAllPhotos || !getPhotoDb()?.updateThumbnailPath) {
      return;
    }

    const rows = getPhotoDb().getAllPhotos();

    for (const row of rows) {
      const rebasedThumbnailPath = rebaseManagedThumbnailPath(row?.thumbnail_path);

      if (!rebasedThumbnailPath || rebasedThumbnailPath === row.thumbnail_path) {
        continue;
      }

      try {
        await fs.access(rebasedThumbnailPath);
        getPhotoDb().updateThumbnailPath(row.id, rebasedThumbnailPath);
      } catch {
        // Thumbnail cache is disposable. If a migrated file is missing, the row can
        // be refreshed later by regeneration or maintenance actions.
      }
    }
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
    const directory = getThumbnailDirectoryPath();
    const fileName = `${fileHash}${extension}`;
    // Hashes can also arrive from restored rows. Never let one become a path,
    // including a nested directory that could be a Windows junction.
    if (/[:\\/]/.test(fileName) || path.basename(fileName) !== fileName) {
      return null;
    }
    const target = path.join(directory, fileName);
    const relative = path.relative(path.resolve(directory), path.resolve(target));
    if (!relative || relative === '..' || relative.startsWith(`..${path.sep}`)
      || path.isAbsolute(relative)) {
      return null;
    }
    return target;
  }

  async function createThumbnailFromNativeImage(image, fileHash, options = {}) {
    const { force = false } = options;
    const thumbnailPath = getManagedThumbnailPath(fileHash);

    if (!thumbnailPath) {
      return null;
    }

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

    if (!thumbnailPath) {
      return null;
    }

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

    return getPhotoDb().updateThumbnailPath(row.id, nextThumbnailPath) || {
      ...row,
      thumbnail_path: nextThumbnailPath,
    };
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

    await ensureDir(getThumbnailDirectoryPath());
  }

  return {
    getThumbnailStorageRootPath,
    getManagedThumbnailRootPaths,
    getManagedThumbnailDirectoryPaths,
    migrateLegacyThumbnailStorage,
    rebaseManagedThumbnailPath,
    reconcileManagedThumbnailPaths,
    calculateThumbnailFitSize,
    getManagedThumbnailPath,
    createThumbnailFromNativeImage,
    createThumbnail,
    ensureManagedThumbnailForRow,
    isManagedThumbnailPath,
    resetManagedThumbnailDirectories,
  };
}

module.exports = { createThumbnailStore };

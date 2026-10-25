'use strict';

const { mapWithConcurrencyLimit } = require('./concurrency');

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

// Import, tracked-folder refresh and explicit reimport share this response shape.
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

// Orchestrates imports. File/image preparation and persistence remain explicit
// ports so failure ordering and saved metadata can be verified without Electron.
function createPhotoImportService({ photoDb, buildPhotoRecord,
  ensureManagedThumbnailForRow, isSupportedImageFile, normalizeKnownFilePath,
  buildWorldMetadataSyncTargetsFromPhotoRecords, concurrencyLimit }) {
  async function importFiles(filePaths, progressReporter = null) {
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
      concurrencyLimit,
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
    // Reimport may preserve a user-selected world instead of the embedded one.
    // Enrich the identities actually saved in the DB, keeping names and links aligned.
    const savedWorldRecords = photoDb.getPhotosByHashes(
      photoRecords.map((photoRecord) => photoRecord.fileHash)
    ).map((row) => ({ worldId: row.world_id, worldUrl: row.world_url }));

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
      worldMetadataTargets: buildWorldMetadataSyncTargetsFromPhotoRecords(savedWorldRecords),
      failedCount: failedFiles.length,
      failedFiles,
      selectedMonth: buildSelectedMonthFromPhotoRecord(latestRow),
    });
  }

  return { importFiles };
}

module.exports = { createPhotoImportService, createImportSummaryResult };

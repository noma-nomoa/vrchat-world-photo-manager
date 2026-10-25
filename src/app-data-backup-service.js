const path = require('node:path');

// Application workflow: adapters own native dialogs, file I/O, DB, and settings.
// Keeping the sequence here lets tests exercise failure boundaries without Electron.
function createAppDataBackupService({
  selectRestoreFile, readText, photoDb, createAutomaticBackup,
  writePreferences, onDatabaseRestored,
}) {
  async function restoreBackup() {
    const filePath = await selectRestoreFile();
    if (!filePath) return { ok: true, canceled: true };
    const parsed = JSON.parse(await readText(filePath));
    if (!parsed || typeof parsed !== 'object') {
      throw new Error('バックアップファイルを読み込めませんでした');
    }
    if (parsed.appName && parsed.appName !== 'WorldShot Log') {
      throw new Error('WorldShot Log のバックアップではありません');
    }

    // Reject before rotating backups or changing data; retain the same envelope
    // for both validations so nested metadata cannot change its interpretation.
    photoDb.validateBackupSnapshot(parsed);
    const automaticBackup = await createAutomaticBackup('before-restore');
    const restoreResult = photoDb.restoreBackupSnapshot(parsed);
    onDatabaseRestored();
    if (parsed.preferences && typeof parsed.preferences === 'object') {
      await writePreferences(parsed.preferences);
    }
    return {
      ok: true, canceled: false, filePath, fileName: path.basename(filePath),
      automaticBackupFilePath: automaticBackup.filePath,
      automaticBackupFileName: automaticBackup.fileName,
      ...restoreResult.afterCounts,
      restoredPhotoCount: restoreResult.restoredPhotoCount,
      restoredTrackedFolderCount: restoreResult.restoredTrackedFolderCount,
      restoredWorldCacheCount: restoreResult.restoredWorldCacheCount,
      restoredTagCount: restoreResult.restoredTagCount,
    };
  }
  return { restoreBackup };
}

module.exports = { createAppDataBackupService };

const { contextBridge, ipcRenderer, webUtils } = require('electron');

// Renderer-facing bridge.
// Keep this list grouped by workflow so main/renderer responsibilities stay
// easy to follow when features are added or removed later.
contextBridge.exposeInMainWorld('electronAPI', {
  // Import / maintenance actions.
  importImages: () => ipcRenderer.invoke('import-images'),
  importFolder: () => ipcRenderer.invoke('import-folder'),
  importDroppedFiles: (files) => {
    const fileList = Array.from(files || []);
    const paths = fileList
      .map((file) => {
        try {
          return webUtils.getPathForFile(file);
        } catch {
          return '';
        }
      })
      .filter(
        (filePath) =>
          typeof filePath === 'string' && filePath.trim().length > 0
      );

    return ipcRenderer.invoke('import-dropped-paths', paths);
  },
  regenerateThumbnails: (payload) =>
    ipcRenderer.invoke('regenerate-thumbnails', payload),
  regenerateMissingThumbnails: () =>
    ipcRenderer.invoke('regenerate-missing-thumbnails'),
  reimportRegisteredPhotos: (payload) =>
    ipcRenderer.invoke('reimport-registered-photos', payload),
  deletePhoto: (photoId) => ipcRenderer.invoke('delete-photo', { photoId }),
  deletePhotos: (photoIds) => ipcRenderer.invoke('delete-photos', { photoIds }),
  deletePhotosByMonth: (payload) =>
    ipcRenderer.invoke('delete-photos-by-month', payload),
  deleteAllPhotos: () => ipcRenderer.invoke('delete-all-photos'),
  clearThumbnailCache: () => ipcRenderer.invoke('clear-thumbnail-cache'),
  resetDatabase: () => ipcRenderer.invoke('reset-database'),
  checkAppDataHealth: () => ipcRenderer.invoke('check-app-data-health'),
  getHealthIssuePhotos: (issueKind) =>
    ipcRenderer.invoke('get-health-issue-photos', { issueKind }),
  getWorldMetadataIssuePhotos: () =>
    ipcRenderer.invoke('get-world-metadata-issue-photos'),
  refreshWorldMetadataIssues: () =>
    ipcRenderer.invoke('refresh-world-metadata-issues'),
  createAppDataBackup: () => ipcRenderer.invoke('create-app-data-backup'),
  restoreAppDataBackup: () => ipcRenderer.invoke('restore-app-data-backup'),
  exportPhotoCatalog: (format) =>
    ipcRenderer.invoke('export-photo-catalog', { format }),
  uninstallApp: () => ipcRenderer.invoke('uninstall-app'),
  uninstallAppAndDeleteData: () =>
    ipcRenderer.invoke('uninstall-app-and-delete-data'),

  // Read-only overview / sidebar data.
  getApplicationDataSummary: () =>
    ipcRenderer.invoke('get-application-data-summary'),
  getSidebarData: () => ipcRenderer.invoke('get-sidebar-data'),
  getWorldSidebarData: (sortMode) =>
    ipcRenderer.invoke('get-world-sidebar-data', sortMode),
  getLatestMonth: () => ipcRenderer.invoke('get-latest-month'),
  getTrackedFolders: () => ipcRenderer.invoke('get-tracked-folders'),
  getBackgroundImagePreference: () =>
    ipcRenderer.invoke('get-background-image-preference'),
  selectBackgroundImage: () => ipcRenderer.invoke('select-background-image'),
  setBackgroundImagePreference: (filePath) =>
    ipcRenderer.invoke('set-background-image-preference', { filePath }),
  startAppUpdateDownload: () =>
    ipcRenderer.invoke('start-app-update-download'),
  installDownloadedAppUpdate: () =>
    ipcRenderer.invoke('install-downloaded-app-update'),
  addTrackedFolder: () => ipcRenderer.invoke('add-tracked-folder'),
  removeTrackedFolder: (folderPath) =>
    ipcRenderer.invoke('remove-tracked-folder', folderPath),

  // Photo listing and world metadata.
  getPhotosByMonth: (year, month) =>
    ipcRenderer.invoke('get-photos-by-month', year, month),
  getPhotosByYear: (year) => ipcRenderer.invoke('get-photos-by-year', year),
  getPhotosByWorldSelection: (selection) =>
    ipcRenderer.invoke('get-photos-by-world-selection', selection),
  getWorldMetadata: (worldId) => ipcRenderer.invoke('get-world-metadata', worldId),
  getLabelCatalog: () => ipcRenderer.invoke('get-label-catalog'),
  getPhotoLabels: (photoId) => ipcRenderer.invoke('get-photo-labels', photoId),
  replacePhotoLabels: (photoId, labels) =>
    ipcRenderer.invoke('replace-photo-labels', { photoId, labels }),

  // Mutations from modal/editor flows.
  refreshTrackedFolders: (folderPaths) =>
    ipcRenderer.invoke('refresh-tracked-folders', folderPaths),
  updateWorldName: (photoId, worldNameManual) =>
    ipcRenderer.invoke('update-world-name', { photoId, worldNameManual }),
  updateWorldSettings: (photoId, payload) =>
    ipcRenderer.invoke('update-world-settings', { photoId, ...payload }),
  updatePhotoMemo: (photoId, memoText) =>
    ipcRenderer.invoke('update-photo-memo', { photoId, memoText }),
  rereadWorldName: (payload) =>
    ipcRenderer.invoke('reread-world-name', payload),
  startWorldMetadataSync: (targets) =>
    ipcRenderer.invoke('start-world-metadata-sync', { targets }),
  openExternalUrl: (url) => ipcRenderer.invoke('open-external-url', url),
  resolvePhotoAccess: (payload) =>
    ipcRenderer.invoke('resolve-photo-access', payload),
  openLocalFile: (payload) => ipcRenderer.invoke('open-local-file', payload),
  openContainingFolder: (payload) =>
    ipcRenderer.invoke('open-containing-folder', payload),
  saveEditedPhoto: (payload) =>
    ipcRenderer.invoke('save-edited-photo', payload),
  getPhotoEditorOverlayAssets: () =>
    ipcRenderer.invoke('get-photo-editor-overlay-assets'),
  selectPhotoEditorOverlayImages: () =>
    ipcRenderer.invoke('select-photo-editor-overlay-images'),
  deletePhotoEditorOverlayAsset: (assetId) =>
    ipcRenderer.invoke('delete-photo-editor-overlay-asset', { assetId }),
  getAiSubjectModelStatus: () =>
    ipcRenderer.invoke('get-ai-subject-model-status'),
  downloadAiSubjectModel: (modelId) =>
    ipcRenderer.invoke('download-ai-subject-model', { modelId }),
  runAiSubjectModel: (payload) =>
    ipcRenderer.invoke('run-ai-subject-model', payload),
  deleteAiSubjectModel: (modelId) =>
    ipcRenderer.invoke('delete-ai-subject-model', { modelId }),
  openAiSubjectModelFolder: (modelId) =>
    ipcRenderer.invoke('open-ai-subject-model-folder', { modelId }),
  onAiSubjectModelDownloadProgress: (callback) => {
    if (typeof callback !== 'function') {
      return () => {};
    }

    const listener = (_event, payload) => callback(payload);
    ipcRenderer.on('ai-subject-model-download-progress', listener);
    return () =>
      ipcRenderer.removeListener(
        'ai-subject-model-download-progress',
        listener
      );
  },
  updateFavoriteStatus: (photoId, isFavorite) =>
    ipcRenderer.invoke('update-favorite-status', { photoId, isFavorite }),
  updateFavoriteStatuses: (photoIds, isFavorite) =>
    ipcRenderer.invoke('update-favorite-statuses', { photoIds, isFavorite }),

  // Push-style notifications from main.
  onProcessingProgress: (listener) => {
    if (typeof listener !== 'function') {
      return () => {};
    }

    const wrappedListener = (_event, payload) => {
      listener(payload);
    };

    ipcRenderer.on('processing-progress', wrappedListener);

    return () => {
      ipcRenderer.removeListener('processing-progress', wrappedListener);
    };
  },
  onWorldMetadataUpdated: (listener) => {
    if (typeof listener !== 'function') {
      return () => {};
    }

    const wrappedListener = (_event, payload) => {
      listener(payload);
    };

    ipcRenderer.on('world-metadata-updated', wrappedListener);

    return () => {
      ipcRenderer.removeListener('world-metadata-updated', wrappedListener);
    };
  },
  onAppUpdateStatus: (listener) => {
    if (typeof listener !== 'function') {
      return () => {};
    }

    const wrappedListener = (_event, payload) => {
      listener(payload);
    };

    ipcRenderer.on('app-update-status', wrappedListener);

    return () => {
      ipcRenderer.removeListener('app-update-status', wrappedListener);
    };
  },
  onAppUpdateAction: (listener) => {
    if (typeof listener !== 'function') {
      return () => {};
    }

    const wrappedListener = (_event, payload) => {
      listener(payload);
    };

    ipcRenderer.on('app-update-action', wrappedListener);

    return () => {
      ipcRenderer.removeListener('app-update-action', wrappedListener);
    };
  },
});

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
  deletePhoto: (photoId) => ipcRenderer.invoke('delete-photo', { photoId }),
  deletePhotos: (photoIds) => ipcRenderer.invoke('delete-photos', { photoIds }),
  deletePhotosByMonth: (payload) =>
    ipcRenderer.invoke('delete-photos-by-month', payload),
  deleteAllPhotos: () => ipcRenderer.invoke('delete-all-photos'),
  clearThumbnailCache: () => ipcRenderer.invoke('clear-thumbnail-cache'),
  resetDatabase: () => ipcRenderer.invoke('reset-database'),

  // Read-only overview / sidebar data.
  getApplicationDataSummary: () =>
    ipcRenderer.invoke('get-application-data-summary'),
  getSidebarData: () => ipcRenderer.invoke('get-sidebar-data'),
  getLatestMonth: () => ipcRenderer.invoke('get-latest-month'),
  getTrackedFolders: () => ipcRenderer.invoke('get-tracked-folders'),
  selectBackgroundImage: () => ipcRenderer.invoke('select-background-image'),
  addTrackedFolder: () => ipcRenderer.invoke('add-tracked-folder'),
  removeTrackedFolder: (folderPath) =>
    ipcRenderer.invoke('remove-tracked-folder', folderPath),

  // Photo listing and world metadata.
  getPhotosByMonth: (year, month) =>
    ipcRenderer.invoke('get-photos-by-month', year, month),
  getPhotosByYear: (year) => ipcRenderer.invoke('get-photos-by-year', year),
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
});

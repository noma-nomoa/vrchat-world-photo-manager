const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
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
  getSidebarData: () => ipcRenderer.invoke('get-sidebar-data'),
  getLatestMonth: () => ipcRenderer.invoke('get-latest-month'),
  getTrackedFolders: () => ipcRenderer.invoke('get-tracked-folders'),
  addTrackedFolder: () => ipcRenderer.invoke('add-tracked-folder'),
  removeTrackedFolder: (folderPath) =>
    ipcRenderer.invoke('remove-tracked-folder', folderPath),
  getPhotosByMonth: (year, month) =>
    ipcRenderer.invoke('get-photos-by-month', year, month),
  getWorldMetadata: (worldId) => ipcRenderer.invoke('get-world-metadata', worldId),
  getLabelCatalog: () => ipcRenderer.invoke('get-label-catalog'),
  getPhotoLabels: (photoId) => ipcRenderer.invoke('get-photo-labels', photoId),
  replacePhotoLabels: (photoId, labels) =>
    ipcRenderer.invoke('replace-photo-labels', { photoId, labels }),
  refreshTrackedFolders: (folderPaths) =>
    ipcRenderer.invoke('refresh-tracked-folders', folderPaths),
  updateWorldName: (photoId, worldNameManual) =>
    ipcRenderer.invoke('update-world-name', { photoId, worldNameManual }),
  updateWorldSettings: (photoId, payload) =>
    ipcRenderer.invoke('update-world-settings', { photoId, ...payload }),
  updatePhotoMemo: (photoId, memoText) =>
    ipcRenderer.invoke('update-photo-memo', { photoId, memoText }),
  rereadWorldName: (photoId) =>
    ipcRenderer.invoke('reread-world-name', { photoId }),
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
});

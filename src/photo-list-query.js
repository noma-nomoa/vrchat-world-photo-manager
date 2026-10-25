(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PhotoListQuery = api;
})(globalThis, function () {
  'use strict';

  // Date interpretation deliberately follows the existing local-time parser,
  // including Date's calendar rollover and the groupDate fallback.
  function parsePhotoSortTimestamp(value) {
    if (typeof value !== 'string' || value.trim().length === 0) return null;
    const match = value.trim().match(
      /^(\d{4})[/-](\d{2})[/-](\d{2})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/
    );
    if (!match) return null;
    const [, year, month, day, hour = '00', minute = '00', second = '00'] = match;
    const timestamp = new Date(
      Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second)
    ).getTime();
    return Number.isFinite(timestamp) ? timestamp : null;
  }

  function getPhotoSortTimestamp(photo) {
    return parsePhotoSortTimestamp(photo?.takenAt) ?? parsePhotoSortTimestamp(photo?.groupDate) ?? 0;
  }

  function comparePhotos(leftPhoto, rightPhoto, sortOrder = 'desc') {
    const timestampDiff = getPhotoSortTimestamp(leftPhoto) - getPhotoSortTimestamp(rightPhoto);
    if (timestampDiff !== 0) return sortOrder === 'asc' ? timestampDiff : -timestampDiff;
    const idDiff = Number(leftPhoto?.id || 0) - Number(rightPhoto?.id || 0);
    if (idDiff !== 0) return sortOrder === 'asc' ? idDiff : -idDiff;
    // Filename ties have always been ascending in both date directions.
    return String(leftPhoto?.fileName || '').localeCompare(String(rightPhoto?.fileName || ''), 'ja');
  }

  function sortPhotos(photos, sortOrder = 'desc') {
    return [...(Array.isArray(photos) ? photos : [])].sort((left, right) => comparePhotos(left, right, sortOrder));
  }

  function normalizeSearchText(value) {
    return typeof value === 'string' ? value.trim() : '';
  }

  function normalizeSearchQuery(value) {
    return normalizeSearchText(value).normalize('NFKC').toLocaleLowerCase('ja-JP');
  }

  function getPhotoOrientationTier(photo) {
    if (!photo) return null;
    if (photo.orientationTier) return photo.orientationTier;
    const width = Number(photo.imageWidth);
    const height = Number(photo.imageHeight);
    if (Number.isFinite(width) && Number.isFinite(height)) {
      if (width === height) return 'square';
      return width > height ? 'landscape' : 'portrait';
    }
    return null;
  }

  function getSearchTargetText(photo, searchScope = 'world') {
    if (!photo) return '';
    const fields = searchScope === 'memo' ? [photo.memoText]
      : searchScope === 'printNote' ? [photo.printNoteText]
        : [photo.worldName, photo.rawWorldName, photo.worldId];
    return fields.filter(Boolean).join(' ').normalize('NFKC').toLocaleLowerCase('ja-JP');
  }

  // Capture one immutable set of filter choices for a synchronous list query.
  // Only the search needle is prepared once; photo data is read on each call.
  function createMatcher({ favoriteOnly = false, orientation = 'all', labelFilters = [],
    labelMode = 'or', searchText = '', searchScope = 'world' } = {}) {
    const selectedLabels = [...labelFilters];
    const hasSearch = Boolean(searchText);
    const normalizedSearch = hasSearch ? normalizeSearchQuery(searchText) : '';
    return photo => {
      if (favoriteOnly && !photo.isFavorite) return false;
      if (orientation !== 'all' && getPhotoOrientationTier(photo) !== orientation) return false;
      if (selectedLabels.length > 0) {
        const photoLabels = new Set((Array.isArray(photo.photoLabels) ? photo.photoLabels : [])
          .map(label => label?.normalizedName).filter(Boolean));
        const matchesLabels = labelMode === 'and'
          ? selectedLabels.every(name => photoLabels.has(name))
          : selectedLabels.some(name => photoLabels.has(name));
        if (!matchesLabels) return false;
      }
      return !hasSearch || getSearchTargetText(photo, searchScope).includes(normalizedSearch);
    };
  }

  function matchesPhoto(photo, options) {
    return createMatcher(options)(photo);
  }

  function queryPhotos(photos, options = {}) {
    return sortPhotos((Array.isArray(photos) ? photos : []).filter(createMatcher(options)), options.sortOrder);
  }

  function getLabelCatalog(photos) {
    const labelMap = new Map();
    for (const photo of Array.isArray(photos) ? photos : []) {
      const seenNames = new Set();
      for (const label of Array.isArray(photo.photoLabels) ? photo.photoLabels : []) {
        if (!label?.normalizedName || seenNames.has(label.normalizedName)) continue;
        seenNames.add(label.normalizedName);
        if (!labelMap.has(label.normalizedName)) {
          labelMap.set(label.normalizedName, { normalizedName: label.normalizedName,
            name: label.name || label.normalizedName, colorHex: label.colorHex || '#6D5EF6', photoCount: 0 });
        }
        labelMap.get(label.normalizedName).photoCount += 1;
      }
    }
    return [...labelMap.values()].sort((left, right) => left.name.localeCompare(right.name, 'ja'));
  }

  return Object.freeze({ parsePhotoSortTimestamp, getPhotoSortTimestamp, comparePhotos, sortPhotos,
    normalizeSearchText, normalizeSearchQuery, getPhotoOrientationTier, getSearchTargetText,
    createMatcher, matchesPhoto, queryPhotos, getLabelCatalog });
});

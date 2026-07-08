(function () {
  'use strict';

  const IMAGE_OVERLAY_LIMIT = 24;
  const IMAGE_OVERLAY_MIN_SIZE = 0.018;
  const IMAGE_OVERLAY_DEFAULT_EDGE = 0.28;
  const IMAGE_OVERLAY_BLEND_MODES = Object.freeze([
    'source-over',
    'multiply',
    'screen',
    'overlay',
    'soft-light',
    'hard-light',
    'darken',
    'lighten',
  ]);
  const MASK_MODES = Object.freeze([
    'normal',
    'subject-only',
    'background-only',
  ]);

  function clampNumber(value, min, max, fallback = 0) {
    const number = Number(value);

    if (!Number.isFinite(number)) {
      return fallback;
    }

    return Math.min(max, Math.max(min, number));
  }

  function createImageOverlayId() {
    return `image-${Date.now().toString(36)}-${Math.random()
      .toString(36)
      .slice(2, 8)}`;
  }

  function normalizeMaskMode(value) {
    return MASK_MODES.includes(value) ? value : 'normal';
  }

  function normalizeCoordinateSpace(value) {
    return value === 'output' ? 'output' : 'source';
  }

  function normalizeOverlayAsset(rawAsset = {}) {
    const id =
      typeof rawAsset?.id === 'string' && rawAsset.id.trim()
        ? rawAsset.id.trim()
        : '';
    const fileUrl =
      typeof rawAsset?.fileUrl === 'string' && rawAsset.fileUrl.trim()
        ? rawAsset.fileUrl.trim()
        : '';

    if (!id || !fileUrl) {
      return null;
    }

    return {
      id,
      fileName:
        typeof rawAsset?.fileName === 'string' && rawAsset.fileName.trim()
          ? rawAsset.fileName.trim()
          : id,
      fileUrl,
      width: Math.max(0, Number(rawAsset?.width) || 0),
      height: Math.max(0, Number(rawAsset?.height) || 0),
      sizeBytes: Math.max(0, Number(rawAsset?.sizeBytes) || 0),
      updatedAt: Math.max(0, Number(rawAsset?.updatedAt) || 0),
    };
  }

  function normalizeOverlayAssets(assets = []) {
    return (Array.isArray(assets) ? assets : [])
      .map(normalizeOverlayAsset)
      .filter(Boolean);
  }

  function getImageOverlayAspectRatio(overlay = {}) {
    const width = Number(
      overlay?.naturalWidth || overlay?.width || overlay?.assetWidth
    );
    const height = Number(
      overlay?.naturalHeight || overlay?.height || overlay?.assetHeight
    );

    return width > 0 && height > 0 ? width / height : 1;
  }

  function normalizeImageOverlayBlendMode(value) {
    return IMAGE_OVERLAY_BLEND_MODES.includes(value)
      ? value
      : 'source-over';
  }

  function getDefaultImageOverlayState(asset = {}, overrides = {}) {
    const normalizedAsset = normalizeOverlayAsset(asset) || {};
    const {
      canvasAspectRatio: rawCanvasAspectRatio,
      ...stateOverrides
    } = overrides || {};
    const aspectRatio = getImageOverlayAspectRatio({
      width: normalizedAsset.width,
      height: normalizedAsset.height,
    });
    const safeAspectRatio = Math.max(0.05, aspectRatio);
    const canvasAspectRatio = Number(rawCanvasAspectRatio);
    const safeCanvasAspectRatio =
      Number.isFinite(canvasAspectRatio) && canvasAspectRatio > 0
        ? canvasAspectRatio
        : 1;
    const normalizedAspectRatio = safeAspectRatio / safeCanvasAspectRatio;
    const maxSize = 0.72;
    const minSize = IMAGE_OVERLAY_MIN_SIZE;
    let defaultWidth = IMAGE_OVERLAY_DEFAULT_EDGE;
    let defaultHeight = defaultWidth / normalizedAspectRatio;

    if (defaultHeight > maxSize) {
      defaultHeight = maxSize;
      defaultWidth = defaultHeight * normalizedAspectRatio;
    }

    if (defaultWidth > maxSize) {
      defaultWidth = maxSize;
      defaultHeight = defaultWidth / normalizedAspectRatio;
    }

    if (defaultHeight < minSize) {
      defaultHeight = minSize;
      defaultWidth = defaultHeight * normalizedAspectRatio;
    }

    if (defaultWidth > maxSize) {
      defaultWidth = maxSize;
      defaultHeight = defaultWidth / normalizedAspectRatio;
    }

    if (defaultWidth < minSize) {
      defaultWidth = minSize;
      defaultHeight = defaultWidth / normalizedAspectRatio;
    }

    if (defaultHeight > maxSize) {
      defaultHeight = maxSize;
      defaultWidth = defaultHeight * normalizedAspectRatio;
    }

    return {
      id: createImageOverlayId(),
      assetId: normalizedAsset.id || '',
      fileName: normalizedAsset.fileName || 'overlay',
      fileUrl: normalizedAsset.fileUrl || '',
      x: 0.5,
      y: 0.5,
      width: clampNumber(defaultWidth, minSize, maxSize, IMAGE_OVERLAY_DEFAULT_EDGE),
      height: clampNumber(defaultHeight, minSize, maxSize, defaultHeight),
      opacity: 1,
      blendMode: 'source-over',
      maskMode: 'normal',
      space: 'source',
      naturalWidth: normalizedAsset.width || 0,
      naturalHeight: normalizedAsset.height || 0,
      ...stateOverrides,
    };
  }

  function normalizeImageOverlayState(overlay = {}) {
    const defaults = getDefaultImageOverlayState();
    const id =
      typeof overlay?.id === 'string' && overlay.id.trim()
        ? overlay.id.trim()
        : defaults.id;
    const fileUrl =
      typeof overlay?.fileUrl === 'string' && overlay.fileUrl.trim()
        ? overlay.fileUrl.trim()
        : '';
    const naturalWidth = Math.max(0, Number(overlay?.naturalWidth) || 0);
    const naturalHeight = Math.max(0, Number(overlay?.naturalHeight) || 0);

    return {
      id,
      assetId:
        typeof overlay?.assetId === 'string' && overlay.assetId.trim()
          ? overlay.assetId.trim()
          : '',
      fileName:
        typeof overlay?.fileName === 'string' && overlay.fileName.trim()
          ? overlay.fileName.trim()
          : 'overlay',
      fileUrl,
      x: clampNumber(overlay?.x, -0.5, 1.5, defaults.x),
      y: clampNumber(overlay?.y, -0.5, 1.5, defaults.y),
      width: clampNumber(
        overlay?.width,
        IMAGE_OVERLAY_MIN_SIZE,
        2,
        defaults.width
      ),
      height: clampNumber(
        overlay?.height,
        IMAGE_OVERLAY_MIN_SIZE,
        2,
        defaults.height
      ),
      opacity: clampNumber(overlay?.opacity, 0, 1, defaults.opacity),
      blendMode: normalizeImageOverlayBlendMode(
        overlay?.blendMode || defaults.blendMode
      ),
      maskMode: normalizeMaskMode(overlay?.maskMode),
      space: normalizeCoordinateSpace(overlay?.space ?? defaults.space),
      naturalWidth,
      naturalHeight,
    };
  }

  function normalizeImageOverlays(imageOverlays = []) {
    return (Array.isArray(imageOverlays) ? imageOverlays : [])
      .map(normalizeImageOverlayState)
      .filter((overlay) => Boolean(overlay.fileUrl))
      .slice(0, IMAGE_OVERLAY_LIMIT);
  }

  globalThis.PhotoEditorImageOverlay = {
    createImageOverlayId,
    normalizeOverlayAsset,
    normalizeOverlayAssets,
    getImageOverlayAspectRatio,
    normalizeImageOverlayBlendMode,
    getDefaultImageOverlayState,
    normalizeImageOverlayState,
    normalizeImageOverlays,
  };
})();

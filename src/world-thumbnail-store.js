'use strict';

const path = require('node:path');

// World thumbnails: one small JPEG per world, made once from the image that
// VRChat publishes for the world and then reused by the world card view, so
// the cards never load anything from the network.
//
// The source is always the "image" of the world (imageUrl / og:image). A
// preview video set by the author lives in a different field
// (previewYoutubeId), so it can never be picked here; a response that is not
// an image is rejected as well.
function createWorldThumbnailStore({
  fs,
  nativeImage,
  fetchImpl,
  getDirectoryPath,
  userAgent,
  timeoutMs = 15000,
}) {
  const THUMBNAIL_MAX_WIDTH = 480;
  const THUMBNAIL_JPEG_QUALITY = 80;
  const THUMBNAIL_EXTENSION = '.jpg';
  const MAX_DOWNLOAD_BYTES = 12 * 1024 * 1024;
  const RESIZED_SOURCE_WIDTH = 512;
  // World images are files on vrchat.cloud. The website (vrchat.com) is left
  // out on purpose: for a world that is private or gone, its page only offers
  // the generic VRChat banner, which must not become the thumbnail.
  const ALLOWED_HOST_DOMAINS = ['vrchat.cloud'];

  // Only https URLs on the hosts VRChat serves world images from are requested.
  function normalizeSourceUrl(value) {
    if (typeof value !== 'string' || value.trim().length === 0) {
      return null;
    }

    let parsedUrl;

    try {
      parsedUrl = new URL(value.trim());
    } catch {
      return null;
    }

    const hostName = parsedUrl.hostname.toLowerCase();
    const isAllowedHost = ALLOWED_HOST_DOMAINS.some(
      (domain) => hostName === domain || hostName.endsWith(`.${domain}`)
    );

    return parsedUrl.protocol === 'https:' && isAllowedHost
      ? parsedUrl.href
      : null;
  }

  // VRChat serves a resized copy next to the full file:
  // /api/1/file/<file id>/<version>/file -> /api/1/image/<file id>/<version>/512
  function getResizedSourceUrl(sourceUrl) {
    let parsedUrl;

    try {
      parsedUrl = new URL(sourceUrl);
    } catch {
      return null;
    }

    const segments = parsedUrl.pathname.split('/');
    const isFileUrl =
      parsedUrl.hostname === 'api.vrchat.cloud' &&
      segments.length === 7 &&
      segments[1] === 'api' &&
      segments[2] === '1' &&
      segments[3] === 'file' &&
      segments[6] === 'file';

    if (!isFileUrl) {
      return null;
    }

    segments[3] = 'image';
    segments[6] = String(RESIZED_SOURCE_WIDTH);
    parsedUrl.pathname = segments.join('/');
    return parsedUrl.href;
  }

  function getFileName(worldId) {
    const safeWorldId =
      typeof worldId === 'string' ? worldId.replace(/[^A-Za-z0-9_-]/g, '') : '';

    return safeWorldId.length > 0 ? `${safeWorldId}${THUMBNAIL_EXTENSION}` : null;
  }

  // The database stores the file name only; the directory can move with the
  // user profile without leaving stale absolute paths behind.
  function getFilePath(fileName) {
    if (
      typeof fileName !== 'string' ||
      fileName.length === 0 ||
      path.basename(fileName) !== fileName
    ) {
      return null;
    }

    return path.join(getDirectoryPath(), fileName);
  }

  async function hasFile(fileName) {
    const filePath = getFilePath(fileName);

    if (!filePath) {
      return false;
    }

    try {
      await fs.access(filePath);
      return true;
    } catch {
      return false;
    }
  }

  async function downloadImage(url) {
    const response = await fetchImpl(url, {
      headers: {
        'user-agent': userAgent,
        accept: 'image/*',
      },
      signal: AbortSignal.timeout(timeoutMs),
    });

    if (!response.ok) {
      const error = new Error(`HTTP ${response.status}`);
      error.fetchStatus = response.status === 404 ? 'not_found' : 'error';
      throw error;
    }

    const contentType = String(response.headers.get('content-type') || '')
      .trim()
      .toLowerCase();

    if (!contentType.startsWith('image/')) {
      const error = new Error(`not an image (${contentType || 'unknown type'})`);
      error.fetchStatus = 'not_image';
      throw error;
    }

    const declaredLength = Number(response.headers.get('content-length'));

    if (Number.isFinite(declaredLength) && declaredLength > MAX_DOWNLOAD_BYTES) {
      throw new Error('image is too large');
    }

    const buffer = Buffer.from(await response.arrayBuffer());

    if (buffer.length === 0 || buffer.length > MAX_DOWNLOAD_BYTES) {
      throw new Error('image is empty or too large');
    }

    return buffer;
  }

  function encodeThumbnail(buffer) {
    const image = nativeImage.createFromBuffer(buffer);

    if (image.isEmpty()) {
      const error = new Error('image could not be decoded');
      error.fetchStatus = 'not_image';
      throw error;
    }

    const sourceSize = image.getSize();
    const resizedImage =
      sourceSize.width > THUMBNAIL_MAX_WIDTH
        ? image.resize({ width: THUMBNAIL_MAX_WIDTH, quality: 'best' })
        : image;
    const size = resizedImage.getSize();

    return {
      jpegBuffer: resizedImage.toJPEG(THUMBNAIL_JPEG_QUALITY),
      width: size.width,
      height: size.height,
    };
  }

  // Tries the resized copy first and the full file second. Returns either
  // { ok: true, fileName, sourceUrl, width, height, byteSize } or
  // { ok: false, fetchStatus, fetchError }.
  async function saveFromUrl(worldId, imageUrl) {
    const fileName = getFileName(worldId);
    const sourceUrl = normalizeSourceUrl(imageUrl);

    if (!fileName || !sourceUrl) {
      return {
        ok: false,
        fetchStatus: 'no_source',
        fetchError: 'no usable image URL',
      };
    }

    const candidateUrls = [getResizedSourceUrl(sourceUrl), sourceUrl].filter(
      (url, index, urls) => url && urls.indexOf(url) === index
    );
    let lastError = null;

    for (const candidateUrl of candidateUrls) {
      try {
        const downloaded = await downloadImage(candidateUrl);
        const { jpegBuffer, width, height } = encodeThumbnail(downloaded);
        const directoryPath = getDirectoryPath();
        const filePath = path.join(directoryPath, fileName);
        const temporaryPath = `${filePath}.tmp`;

        await fs.mkdir(directoryPath, { recursive: true });
        await fs.writeFile(temporaryPath, jpegBuffer);
        await fs.rename(temporaryPath, filePath);

        return {
          ok: true,
          fileName,
          sourceUrl,
          width,
          height,
          byteSize: jpegBuffer.length,
        };
      } catch (error) {
        lastError = error;
      }
    }

    return {
      ok: false,
      fetchStatus: lastError?.fetchStatus || 'error',
      fetchError: lastError?.message || 'unknown error',
    };
  }

  async function removeAll() {
    await fs.rm(getDirectoryPath(), {
      recursive: true,
      force: true,
      maxRetries: 2,
      retryDelay: 120,
    });
    await fs.mkdir(getDirectoryPath(), { recursive: true });
  }

  return {
    normalizeSourceUrl,
    getResizedSourceUrl,
    getFileName,
    getFilePath,
    hasFile,
    saveFromUrl,
    removeAll,
  };
}

module.exports = { createWorldThumbnailStore };

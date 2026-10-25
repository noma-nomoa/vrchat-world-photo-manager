'use strict';

const path = require('node:path');
const { createHash } = require('node:crypto');
const {
  extractWorldInfoFromRawBuffer,
  parseVrchatFilenameDate,
  parseImageDimensionsFromFilename,
  loadExifTagsSafely,
  extractTakenAtFromTags,
  extractWorldInfo,
  extractPhotoPrintNote,
  getResolutionTier,
  getOrientationTier,
  formatDateTime,
  formatGroupDate,
} = require('./photo-metadata');

// File access, native decoding, thumbnail generation, and time are explicit ports.
function createPhotoRecordBuilder({
  readFile,
  stat,
  decodeImage,
  createThumbnail,
  now = () => new Date(),
}) {
  return async function buildPhotoRecord(filePath) {
    const fileBuffer = await readFile(filePath);
    const imageDetails = decodeImage(fileBuffer);
    const rawWorldInfo = extractWorldInfoFromRawBuffer(fileBuffer);
    const takenAtFromFileName = parseVrchatFilenameDate(filePath);
    const imageDetailsFromFileName = parseImageDimensionsFromFilename(filePath);
    const tags = loadExifTagsSafely(fileBuffer);

    const fileHash = createHash('sha256').update(fileBuffer).digest('hex');
    const thumbnailPathPromise = createThumbnail(filePath, fileHash, {
      sourceImage: imageDetails.image,
    });
    const takenAtDate =
      takenAtFromFileName ||
      extractTakenAtFromTags(tags) ||
      (await stat(filePath)).mtime;
    const worldInfo = extractWorldInfo(tags, fileBuffer, rawWorldInfo);
    const printNoteText = extractPhotoPrintNote(tags);
    const hasDecodedImageDimensions =
      Number.isFinite(imageDetails.imageWidth) &&
      Number.isFinite(imageDetails.imageHeight) &&
      imageDetails.imageWidth > 0 &&
      imageDetails.imageHeight > 0;
    // Windows nativeImage cannot decode WebP directly, although Chromium and the
    // shell thumbnail provider can. ExifReader already reads its canvas header.
    const webpWidth = tags.FileType?.value === 'webp' ? tags.ImageWidth?.value : null;
    const webpHeight = tags.FileType?.value === 'webp' ? tags.ImageHeight?.value : null;
    const hasWebpHeaderDimensions =
      Number.isInteger(webpWidth) && webpWidth > 0 &&
      Number.isInteger(webpHeight) && webpHeight > 0;
    const resolvedImageDetails = hasDecodedImageDimensions
      ? {
          imageWidth: imageDetails.imageWidth,
          imageHeight: imageDetails.imageHeight,
          resolutionTier: imageDetails.resolutionTier,
          orientationTier: imageDetails.orientationTier,
        }
      : hasWebpHeaderDimensions
        ? {
            imageWidth: webpWidth,
            imageHeight: webpHeight,
            resolutionTier: getResolutionTier(webpWidth, webpHeight),
            orientationTier: getOrientationTier(webpWidth, webpHeight),
          }
        : {
            imageWidth: imageDetailsFromFileName?.imageWidth || null,
            imageHeight: imageDetailsFromFileName?.imageHeight || null,
            resolutionTier: imageDetailsFromFileName?.resolutionTier || null,
            orientationTier: imageDetailsFromFileName?.orientationTier || null,
          };
    const nowIso = now().toISOString();
    const thumbnailPath = await thumbnailPathPromise;

    return {
      filePath,
      fileName: path.basename(filePath),
      fileHash,
      takenAt: formatDateTime(takenAtDate),
      takenAtTimestamp: takenAtDate.getTime(),
      groupDate: formatGroupDate(takenAtDate),
      year: takenAtDate.getFullYear(),
      month: takenAtDate.getMonth() + 1,
      day: takenAtDate.getDate(),
      worldId: worldInfo.worldId,
      worldName: worldInfo.worldName,
      worldNameManual: null,
      worldUrl: worldInfo.worldUrl,
      thumbnailPath,
      imageWidth: resolvedImageDetails.imageWidth,
      imageHeight: resolvedImageDetails.imageHeight,
      resolutionTier: resolvedImageDetails.resolutionTier,
      orientationTier: resolvedImageDetails.orientationTier,
      printNoteText,
      memoText: null,
      createdAt: nowIso,
      updatedAt: nowIso,
    };
  };
}

module.exports = { createPhotoRecordBuilder };

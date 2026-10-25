'use strict';

// Image metadata interpretation is independent of Electron, storage, and UI.
const path = require('node:path');
const ExifReader = require('exifreader');
const iconv = require('iconv-lite');

function countJapaneseChars(value) {
  return (value.match(/[\u3040-\u30ff\u3400-\u9fff]/g) || []).length;
}

function countReplacementChars(value) {
  return (value.match(/�/g) || []).length;
}

function countMojibakeMarkers(value) {
  return (
    value.match(/[ÃÂâåãæçèéêëìíîïðñòóôõöøùúûüýþÿ竄繝縺繧鬘蜿榊]/g) || []
  ).length;
}

function countControlChars(value) {
  return (value.match(/[\u0000-\u001f\u007f]/g) || []).length;
}

function countQuestionMarks(value) {
  return (value.match(/\?/g) || []).length;
}

function textScore(value) {
  if (typeof value !== 'string' || value.length === 0) {
    return -9999;
  }

  return (
    countJapaneseChars(value) * 4 -
    countMojibakeMarkers(value) * 3 -
    countReplacementChars(value) * 8 -
    countControlChars(value) * 10 -
    countQuestionMarks(value)
  );
}

function tryLatin1ToUtf8(value) {
  try {
    return Buffer.from(value, 'latin1').toString('utf8');
  } catch {
    return value;
  }
}

function tryShiftJisToUtf8(value) {
  try {
    return iconv.encode(value, 'shift_jis').toString('utf8');
  } catch {
    return value;
  }
}

function repairUtf8Mojibake(value) {
  if (typeof value !== 'string' || value.length === 0) {
    return value;
  }

  const candidates = new Set([
    value,
    tryLatin1ToUtf8(value),
    tryShiftJisToUtf8(value),
    tryLatin1ToUtf8(tryShiftJisToUtf8(value)),
    tryShiftJisToUtf8(tryLatin1ToUtf8(value)),
  ]);

  let best = value;
  let bestScore = textScore(value);

  for (const candidate of candidates) {
    const score = textScore(candidate);
    if (score > bestScore) {
      best = candidate;
      bestScore = score;
    }
  }

  return best;
}

function isExifReaderTagEnvelope(rawValue) {
  return (
    rawValue &&
    typeof rawValue === 'object' &&
    !Array.isArray(rawValue) &&
    (Object.prototype.hasOwnProperty.call(rawValue, 'value') ||
      Object.prototype.hasOwnProperty.call(rawValue, 'description') ||
      Object.prototype.hasOwnProperty.call(rawValue, 'attributes'))
  );
}

function isScalarTagTextValue(rawValue) {
  return (
    typeof rawValue === 'string' ||
    typeof rawValue === 'number' ||
    typeof rawValue === 'boolean' ||
    typeof rawValue === 'bigint'
  );
}

function collectTagTextCandidates(rawValue, visited = new Set()) {
  if (rawValue == null) {
    return [];
  }

  if (isScalarTagTextValue(rawValue)) {
    return [String(rawValue)];
  }

  if (typeof rawValue !== 'object') {
    return [String(rawValue)];
  }

  if (visited.has(rawValue)) {
    return [];
  }

  visited.add(rawValue);

  if (isExifReaderTagEnvelope(rawValue)) {
    const valueCandidates = collectTagTextCandidates(rawValue.value, visited);

    if (valueCandidates.length > 0) {
      if (isScalarTagTextValue(rawValue.value)) {
        return [
          ...valueCandidates,
          ...collectTagTextCandidates(rawValue.description, visited),
        ];
      }

      return valueCandidates;
    }

    return collectTagTextCandidates(rawValue.description, visited);
  }

  if (Array.isArray(rawValue)) {
    return rawValue.flatMap((item) => collectTagTextCandidates(item, visited));
  }

  const nested = Object.values(rawValue).flatMap((item) =>
    collectTagTextCandidates(item, visited)
  );

  if (nested.length > 0) {
    return nested;
  }

  try {
    return [JSON.stringify(rawValue)];
  } catch {
    return [];
  }
}

function sanitizeExtractedText(value) {
  if (typeof value !== 'string') {
    return null;
  }

  const sanitized = value
    .replace(/\u0000/g, ' ')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();

  if (/^\[object\s+[^\]]+\]$/i.test(sanitized)) {
    return null;
  }

  return sanitized.length > 0 ? sanitized : null;
}

function isEmptyPrintNotePlaceholder(value) {
  if (typeof value !== 'string') {
    return false;
  }

  const trimmed = value.trim();

  if (trimmed === '{}' || trimmed === '[]') {
    return true;
  }

  try {
    const parsed = JSON.parse(trimmed);

    if (Array.isArray(parsed)) {
      return parsed.length === 0;
    }

    if (parsed && typeof parsed === 'object') {
      return Object.values(parsed).every((entry) => {
        if (entry == null) {
          return true;
        }

        if (typeof entry === 'string') {
          return entry.trim().length === 0;
        }

        if (Array.isArray(entry)) {
          return entry.length === 0;
        }

        return false;
      });
    }
  } catch {
    // Not JSON-like, so it can still be a real print note.
  }

  return false;
}

function pickBestTextCandidate(rawValues) {
  const candidates = new Set();

  for (const rawValue of rawValues) {
    for (const stringValue of collectTagTextCandidates(rawValue)) {
      if (!stringValue) {
        continue;
      }

      candidates.add(stringValue);
      candidates.add(repairUtf8Mojibake(stringValue));

      const strippedNulls = stringValue.replace(/\u0000+/g, '');
      candidates.add(strippedNulls);
      candidates.add(repairUtf8Mojibake(strippedNulls));

      const strippedControls = stringValue.replace(/[\u0000-\u001f\u007f]+/g, ' ');
      candidates.add(strippedControls);
      candidates.add(repairUtf8Mojibake(strippedControls));

      const removedControls = stringValue.replace(/[\u0000-\u001f\u007f]+/g, '');
      candidates.add(removedControls);
      candidates.add(repairUtf8Mojibake(removedControls));
    }
  }

  let best = null;
  let bestScore = -9999;

  for (const candidate of candidates) {
    const sanitized = sanitizeExtractedText(candidate);

    if (!sanitized) {
      continue;
    }

    const score = textScore(sanitized);

    if (score > bestScore) {
      best = sanitized;
      bestScore = score;
    }
  }

  return best;
}

function normalizeWorldId(value) {
  if (typeof value !== 'string') {
    return null;
  }

  const trimmed = value.trim();
  return /^wrld_[0-9a-f-]+$/i.test(trimmed) ? trimmed : null;
}

function isNumericTagValueArray(value) {
  return (
    Array.isArray(value) &&
    value.length > 0 &&
    value.every((item) => typeof item === 'number' && Number.isFinite(item))
  );
}

function normalizeTagValue(tag) {
  if (!tag) return null;

  const valueCandidate = isNumericTagValueArray(tag.value)
    ? null
    : pickBestTextCandidate([tag.value]);

  return valueCandidate || pickBestTextCandidate([tag.description]);
}

function normalizePhotoPrintNoteText(value) {
  const sanitized = sanitizeExtractedText(repairUtf8Mojibake(value));

  if (!sanitized || isEmptyPrintNotePlaceholder(sanitized)) {
    return null;
  }

  return sanitized
    .replace(/\u2044/g, '/')
    .replace(/[\u201c\u201d\uFF02]/g, '"')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/\uFFFD/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim()
    .normalize('NFC');
}

function shouldRepairStoredPrintNote(value) {
  if (typeof value !== 'string') {
    return false;
  }

  const trimmed = value.trim();

  if (!trimmed) {
    return false;
  }

  if (/^\[object\s+[^\]]+\]$/i.test(trimmed)) {
    return true;
  }

  if (isEmptyPrintNotePlaceholder(trimmed)) {
    return true;
  }

  const commaSeparatedSegments = trimmed
    .split(/\s*,\s*/)
    .map((segment) => segment.trim())
    .filter(Boolean);

  if (
    commaSeparatedSegments.some((segment) =>
      /^(?:x-default|[a-z]{2,3}(?:-[a-z0-9]{2,8})+)$/i.test(segment)
    )
  ) {
    return true;
  }

  return /^\{.*(?:x-default|xml:lang|lang=).*\}$/i.test(trimmed);
}

// Only treat the exact "{}" placeholder as missing so unusual but real names still render.
function hasMeaningfulWorldNameContent(value) {
  if (typeof value !== 'string') {
    return false;
  }

  const trimmed = value.trim();

  if (!trimmed) {
    return false;
  }

  return trimmed !== '{}';
}

function normalizeDisplayWorldName(value) {
  const sanitized = sanitizeExtractedText(value);

  if (!sanitized) {
    return null;
  }

  const normalized = sanitized
    .replace(/[⁄∕／]/g, '/')
    .replace(/ǃ/g, '!')
    .replace(/‚/g, ',')
    .replace(/․/g, '.')
    .replace(/［/g, '[')
    .replace(/］/g, ']')
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .replace(/＂/g, '"')
    .replace(/\uFFFD/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim()
    .normalize('NFC');

  return hasMeaningfulWorldNameContent(normalized) ? normalized : null;
}

function normalizeOfficialWorldText(value) {
  const sanitized = sanitizeExtractedText(repairUtf8Mojibake(value));

  if (!sanitized) {
    return null;
  }

  return sanitized
    .replace(/\u2044/g, '/')
    .replace(/\u201a/g, ',')
    .replace(/\u2024/g, '.')
    .replace(/[\u201c\u201d\uFF02]/g, '"')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/\uFFFD/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim()
    .normalize('NFC');
}

function normalizeOfficialWorldTags(tags) {
  if (!Array.isArray(tags)) {
    return [];
  }

  return Array.from(
    new Set(
      tags
        .map((tag) => normalizeOfficialWorldText(tag))
        .filter(Boolean)
    )
  );
}

function pad2(value) {
  return String(value).padStart(2, '0');
}

function formatDateTime(date) {
  return `${date.getFullYear()}/${pad2(date.getMonth() + 1)}/${pad2(
    date.getDate()
  )} ${pad2(date.getHours())}:${pad2(date.getMinutes())}:${pad2(
    date.getSeconds()
  )}`;
}

function formatGroupDate(date) {
  return `${date.getFullYear()}/${pad2(date.getMonth() + 1)}/${pad2(
    date.getDate()
  )}`;
}

function parseExifDateString(value) {
  if (!value || typeof value !== 'string') return null;

  const normalized = value
    .trim()
    .replace(/^(\d{4}):(\d{2}):(\d{2})/, '$1-$2-$3')
    .replace(/\.\d+$/, '');

  const parsed = new Date(normalized);

  if (Number.isNaN(parsed.getTime())) {
    return null;
  }

  return parsed;
}

function parseVrchatFilenameDate(filePath) {
  const fileName = path.basename(filePath);

  const match = fileName.match(
    /VRChat_(\d{4})-(\d{2})-(\d{2})_(\d{2})-(\d{2})-(\d{2})(?:\.\d+)?/i
  );

  if (!match) {
    return null;
  }

  const [, year, month, day, hour, minute, second] = match;

  return new Date(
    Number(year),
    Number(month) - 1,
    Number(day),
    Number(hour),
    Number(minute),
    Number(second)
  );
}

function parseImageDimensionsFromFilename(filePath) {
  const fileName = path.basename(filePath);
  const patterns = [
    /(?:^|[_\-. \[(])(\d{2,5})[xX](\d{2,5})(?:[_\-. \])]|$)/,
    /(?:width|w)(\d{2,5})[^0-9]{0,4}(?:height|h)(\d{2,5})/i,
  ];

  for (const pattern of patterns) {
    const match = fileName.match(pattern);

    if (!match) {
      continue;
    }

    const imageWidth = Number(match[1]);
    const imageHeight = Number(match[2]);

    if (
      !Number.isFinite(imageWidth) ||
      !Number.isFinite(imageHeight) ||
      imageWidth <= 0 ||
      imageHeight <= 0
    ) {
      continue;
    }

    return {
      imageWidth,
      imageHeight,
      resolutionTier: getResolutionTier(imageWidth, imageHeight),
      orientationTier: getOrientationTier(imageWidth, imageHeight),
    };
  }

  return null;
}

function extractTakenAtFromTags(tags) {
  const candidateKeys = [
    'DateTimeOriginal',
    'CreateDate',
    'DateCreated',
    'ModifyDate',
    'File Modification Date/Time',
  ];

  for (const key of candidateKeys) {
    const tag = tags[key];
    if (!tag) continue;

    const value = normalizeTagValue(tag);
    const parsed = parseExifDateString(value);

    if (parsed) {
      return parsed;
    }
  }

  return null;
}

function pickPreferredTag(flatTags, preferredKeys, matcher) {
  for (const key of preferredKeys) {
    if (flatTags[key]) {
      return { key, value: flatTags[key] };
    }
  }

  const fallbackKey = Object.keys(flatTags).find((key) =>
    matcher.test(key.toLowerCase())
  );

  if (!fallbackKey) {
    return null;
  }

  return {
    key: fallbackKey,
    value: flatTags[fallbackKey],
  };
}

function decodeXmlEntities(value) {
  if (typeof value !== 'string') {
    return value;
  }

  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) =>
      String.fromCodePoint(parseInt(hex, 16))
    )
    .replace(/&#([0-9]+);/g, (_, num) =>
      String.fromCodePoint(parseInt(num, 10))
    );
}

function extractXmpAttributeFromText(text, attrName) {
  if (typeof text !== 'string' || text.length === 0) {
    return null;
  }

  const patterns = [
    new RegExp(`(?:vrc:)?${attrName}="([^"]*)"`, 'i'),
    new RegExp(`(?:vrc:)?${attrName}='([^']*)'`, 'i'),
    new RegExp(
      `<(?:[a-z0-9_-]+:)?${attrName}>([\\s\\S]*?)<\\/(?:[a-z0-9_-]+:)?${attrName}>`,
      'i'
    ),
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match && match[1] != null) {
      return decodeXmlEntities(match[1]);
    }
  }

  return null;
}

function extractRelevantXmpTextFromBuffer(fileBuffer) {
  const xmpStartMarker = Buffer.from('<x:xmpmeta', 'utf8');
  const xmpEndMarker = Buffer.from('</x:xmpmeta>', 'utf8');
  const worldIdMarker = Buffer.from('WorldID', 'utf8');
  const worldNameMarker = Buffer.from('WorldDisplayName', 'utf8');

  const xmpStartIndex = fileBuffer.indexOf(xmpStartMarker);

  if (xmpStartIndex !== -1) {
    const xmpEndIndex = fileBuffer.indexOf(xmpEndMarker, xmpStartIndex);

    if (xmpEndIndex !== -1) {
      return fileBuffer.toString(
        'utf8',
        xmpStartIndex,
        xmpEndIndex + xmpEndMarker.length
      );
    }
  }

  const anchorIndices = [
    fileBuffer.indexOf(worldIdMarker),
    fileBuffer.indexOf(worldNameMarker),
  ].filter((index) => index !== -1);

  if (anchorIndices.length === 0) {
    return null;
  }

  const anchorIndex = Math.min(...anchorIndices);
  const sliceStart = Math.max(0, anchorIndex - 4096);
  const sliceEnd = Math.min(fileBuffer.length, anchorIndex + 4096);

  return fileBuffer.toString('utf8', sliceStart, sliceEnd);
}

function extractWorldInfoFromRawBuffer(fileBuffer) {
  const xmpText = extractRelevantXmpTextFromBuffer(fileBuffer);

  if (!xmpText) {
    return {
      worldId: null,
      worldName: null,
    };
  }

  const worldId = extractXmpAttributeFromText(xmpText, 'WorldID');
  const worldName = extractXmpAttributeFromText(xmpText, 'WorldDisplayName');

  return {
    worldId,
    worldName,
  };
}

function extractWorldInfo(tags, fileBuffer, rawBufferWorldInfo = null) {
  const flatTags = {};

  for (const [key, tag] of Object.entries(tags || {})) {
    flatTags[key] = normalizeTagValue(tag);
  }

  const worldIdEntry = pickPreferredTag(
    flatTags,
    ['XMP-vrc:WorldID', 'vrc:WorldID', 'World ID', 'WorldID'],
    /worldid|world id/
  );

  const worldNameEntry = pickPreferredTag(
    flatTags,
    [
      'XMP-vrc:WorldDisplayName',
      'vrc:WorldDisplayName',
      'World Display Name',
      'WorldDisplayName',
    ],
    /worlddisplayname|world display name/
  );

  const resolvedRawWorldInfo =
    rawBufferWorldInfo || extractWorldInfoFromRawBuffer(fileBuffer);

  const worldId = normalizeWorldId(
    resolvedRawWorldInfo.worldId ||
      (worldIdEntry ? worldIdEntry.value : null) ||
      null
  );

  const worldName = normalizeDisplayWorldName(
    pickBestTextCandidate([
      resolvedRawWorldInfo.worldName,
      worldNameEntry ? worldNameEntry.value : null,
    ])
  );

  const worldUrl = worldId
    ? `https://vrchat.com/home/world/${worldId}`
    : null;

  return {
    worldId,
    worldName,
    worldUrl,
  };
}

function extractPhotoPrintNote(tags) {
  const flatTags = {};
  const printNoteCandidates = [];
  const preferredKeys = [
    'XMP-dc:Title',
    'dc:Title',
    'Title',
    'XPTitle',
    'Windows XP Title',
    'IPTC:ObjectName',
    'ObjectName',
    'Object Name',
    'XMP-photoshop:Headline',
    'Headline',
  ];
  const fallbackMatcher =
    /(?:^|[: _-])(title|xptitle|objectname|object name|headline)(?:$|[: _-])/;

  for (const [key, tag] of Object.entries(tags || {})) {
    const normalizedValue = normalizeTagValue(tag);
    flatTags[key] = normalizedValue;

    if (
      normalizedValue &&
      (preferredKeys.includes(key) || fallbackMatcher.test(key.toLowerCase()))
    ) {
      printNoteCandidates.push(normalizedValue);
    }
  }

  const printNoteEntry = pickPreferredTag(flatTags, preferredKeys, fallbackMatcher);
  const preferredPrintNoteText = pickBestTextCandidate(
    printNoteCandidates.length > 0
      ? printNoteCandidates
      : [printNoteEntry?.value || null]
  );

  return normalizePhotoPrintNoteText(preferredPrintNoteText);
}

function loadExifTagsSafely(fileBuffer) {
  try {
    return ExifReader.load(fileBuffer);
  } catch {
    return {};
  }
}

function getResolutionTier(width, height) {
  if (!Number.isFinite(width) || !Number.isFinite(height)) {
    return null;
  }

  const longEdge = Math.max(width, height);
  const shortEdge = Math.min(width, height);

  if (longEdge >= 7680 && shortEdge >= 4320) {
    return '8K';
  }

  if (longEdge >= 3840 && shortEdge >= 2160) {
    return '4K';
  }

  if (longEdge >= 2560 && shortEdge >= 1440) {
    return 'QHD';
  }

  if (longEdge >= 1920 && shortEdge >= 1080) {
    return 'FHD';
  }

  if (longEdge >= 1280 && shortEdge >= 720) {
    return 'HD';
  }

  return null;
}

function getOrientationTier(width, height) {
  if (!Number.isFinite(width) || !Number.isFinite(height)) {
    return null;
  }

  if (width === height) {
    return 'square';
  }

  return width > height ? 'landscape' : 'portrait';
}

module.exports = {
  isEmptyPrintNotePlaceholder,
  normalizeWorldId,
  normalizePhotoPrintNoteText,
  shouldRepairStoredPrintNote,
  normalizeDisplayWorldName,
  normalizeOfficialWorldText,
  normalizeOfficialWorldTags,
  pad2,
  formatDateTime,
  formatGroupDate,
  parseVrchatFilenameDate,
  parseImageDimensionsFromFilename,
  extractTakenAtFromTags,
  extractWorldInfoFromRawBuffer,
  extractWorldInfo,
  extractPhotoPrintNote,
  loadExifTagsSafely,
  getResolutionTier,
  getOrientationTier
};

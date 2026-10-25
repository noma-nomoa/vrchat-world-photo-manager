'use strict';

// Synthetic fixtures generated locally; no user photos or metadata are included.
const fs = require('node:fs');
const path = require('node:path');
const readFixture = name => fs.readFileSync(path.join(__dirname, 'photo-metadata', name));
const WORLD_ID = 'wrld_aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
const OTHER_WORLD_ID = 'wrld_bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';
const FIXED_NOW = '2026-09-10T00:00:00.000Z';
const STAT_DATE = '2026-01-02T03:04:05.000Z';

function createMetadataCases() {
  const cases = [];
  const add = (fn, name, ...args) => cases.push({ fn, name: `${fn}/${name}`, args });
  const textValues = [null, undefined, '', '  ', '夜の海 🐈', 'Cafe\u0301',
    'A\u0000B\nC\tD', '{}', '[]', '{"x-default":""}', '[object Object]',
    '［海］／山（夜）ǃ ＂Cafe＂', '“title” ⁄ ‘note’', 'bad�value', 'x-default, text',
    Buffer.from('日本語の写真', 'utf8').toString('latin1'), 0, false];
  for (const fn of ['normalizePhotoPrintNoteText', 'normalizeDisplayWorldName',
    'normalizeOfficialWorldText', 'shouldRepairStoredPrintNote', 'isEmptyPrintNotePlaceholder']) {
    textValues.forEach((value, index) => add(fn, `text-${index}`, value));
  }
  for (const [index, value] of [null, undefined, '', 'wrld_', WORLD_ID,
    ` ${WORLD_ID} `, WORLD_ID.toUpperCase(), 'wrld_abc', 'wrld_123-456', 'wrld_xyz',
    `https://vrchat.com/home/world/${WORLD_ID}`, 42].entries()) {
    add('normalizeWorldId', `id-${index}`, value);
  }
  for (const [index, tags] of [null, [], ['海', '海', 'Cafe\u0301', 'Café', '{}', '', null],
    ['“Moon”', '\u0000', '[object Object]'], 'not an array'].entries()) {
    add('normalizeOfficialWorldTags', `tags-${index}`, tags);
  }
  const dimensions = [[7680, 4320], [4320, 7680], [7679, 4320], [3840, 2160],
    [2560, 1440], [1920, 1080], [1280, 720], [1279, 720], [7680, 100],
    [320, 180], [500, 500], [0, 0], [-1, 20], [NaN, 100], [Infinity, 100],
    ['1920', '1080'], [undefined, null]];
  for (const fn of ['getResolutionTier', 'getOrientationTier']) {
    dimensions.forEach((args, index) => add(fn, `size-${index}`, ...args));
  }
  const fileNames = ['VRChat_2026-06-07_22-20-00.000_320x180.png',
    'C:/Pictures/VRChat_2024-02-29_23-59-59_1920x1080.jpg',
    'C:/Pictures/VRChat_2026-13-32_25-61-61.webp', 'vrchat_1999-01-01_00-00-00.png',
    'generic-2560X1440.png', 'image (3840x2160).jpg', 'photo_w2048-h1440.webp',
    'photo_width1920_height1080.png', 'foo_1x2.png', 'foo_000x000.png',
    'foo_999999x999999.png', 'no-metadata.png', '海_1080x1920.png', 'foo_1920x1080bar.png'];
  for (const fn of ['parseVrchatFilenameDate', 'parseImageDimensionsFromFilename']) {
    fileNames.forEach((value, index) => add(fn, `file-${index}`, value));
  }
  for (const fn of ['formatDateTime', 'formatGroupDate']) {
    ['2026-06-07T13:20:00Z', '2000-02-29T23:59:59Z', 'invalid'].forEach((value, index) =>
      add(fn, `date-${index}`, new Date(value)));
  }
  [0, 3, 10, -1, 123, undefined].forEach((value, index) => add('pad2', `pad-${index}`, value));
  const dateTags = [{}, { DateTimeOriginal: { value: '2026:06:07 22:20:00' } },
    { CreateDate: { value: '2026-06-07T22:20:00+09:00' } },
    { DateTimeOriginal: { value: 'bad' }, ModifyDate: { value: '2025:02:03 04:05:06.123' } },
    { DateTimeOriginal: { value: [1, 2, 3], description: '2025:01:02 03:04:05' } },
    { DateTimeOriginal: { value: 'bad' }, CreateDate: { description: 'bad' } },
    { DateTimeOriginal: { value: '2020:01:01 00:00:00' }, CreateDate: { value: '2021:01:01 00:00:00' } }];
  dateTags.forEach((tags, index) => add('extractTakenAtFromTags', `tags-${index}`, tags));
  const cyclic = { value: {} };
  cyclic.value.self = cyclic;
  const noteTags = [{}, { Title: { value: '思い出の写真' } },
    { 'XMP-dc:Title': { value: { 'rdf:Alt': [{ value: '星の海', attributes: { 'xml:lang': 'x-default' } }] } } },
    { Title: { value: [65, 66, 67], description: 'Numeric title fallback' } },
    { Title: { value: {}, description: '{}' } }, { Title: cyclic },
    { Title: { value: 'English' }, XPTitle: { description: '日本語の写真' } },
    { unknown_headline: { value: 'Legacy headline' } },
    { dcTitle: { value: 'Must not match the boundary expression' } },
    { Title: { value: { first: {}, second: { value: '最初の候補', description: '別候補' } } } }];
  noteTags.forEach((tags, index) => add('extractPhotoPrintNote', `tags-${index}`, tags));
  const empty = Buffer.alloc(0);
  const rawValues = [empty, readFixture('world.xmp'), Buffer.from(`WorldID='${WORLD_ID}' WorldDisplayName='海 &amp; 山'`),
    Buffer.from(`<WorldID>${WORLD_ID}</WorldID><WorldDisplayName>&#x1F408; &#12354;</WorldDisplayName>`),
    Buffer.from(`<x:xmpmeta WorldID="${WORLD_ID}" WorldDisplayName="Incomplete packet">`),
    Buffer.from('<x:xmpmeta WorldDisplayName="&#x110000;"></x:xmpmeta>'),
    Buffer.concat([Buffer.alloc(5000), Buffer.from(`WorldDisplayName="遠いメタデータ" WorldID="${WORLD_ID}"`), Buffer.alloc(5000)])];
  for (const [index, buffer] of rawValues.entries()) {
    add('extractWorldInfoFromRawBuffer', `raw-${index}`, buffer);
    add('extractWorldInfo', `raw-${index}`, {}, buffer);
  }
  for (const [index, tags] of [{},
    { 'XMP-vrc:WorldID': { value: WORLD_ID }, 'XMP-vrc:WorldDisplayName': { value: '現行タグ' } },
    { 'World ID': { description: WORLD_ID }, 'World Display Name': { description: '旧タグ' } },
    { custom_worldid: { value: WORLD_ID }, custom_worlddisplayname: { value: '{}' } },
    { WorldID: { value: WORLD_ID }, 'XMP-vrc:WorldID': { value: OTHER_WORLD_ID } },
    { WorldID: { value: 'invalid' }, WorldDisplayName: { value: '海' } }].entries()) {
    add('extractWorldInfo', `tag-only-${index}`, tags, empty);
    add('extractWorldInfo', `raw-precedence-${index}`, tags, readFixture('world.xmp'));
  }
  add('extractWorldInfo', 'preparsed-invalid-id-remains-preferred',
    { WorldID: { value: WORLD_ID } }, empty, { worldId: 'invalid', worldName: 'Raw name' });
  for (const name of ['plain.png', 'plain.jpg', 'plain.webp', 'world.png', 'world.jpg', 'dated-world.jpg']) {
    const buffer = readFixture(name);
    add('loadExifTagsSafely', name, buffer);
    add('extractWorldInfoFromRawBuffer', name, buffer);
  }
  [Buffer.alloc(0), Buffer.from('not an image'), readFixture('world.png').subarray(0, 20)]
    .forEach((buffer, index) => add('loadExifTagsSafely', `malformed-${index}`, buffer));
  return cases;
}

function createRecordCases() {
  return [
    { name: 'png-filename-capture', fixture: 'plain.png', fileName: 'VRChat_2026-06-07_22-20-00.000_320x180.png', width: 320, height: 180 },
    { name: 'jpeg-world-xmp', fixture: 'world.jpg', fileName: 'world.jpg', width: 320, height: 180 },
    { name: 'jpeg-exif-before-stat', fixture: 'dated-world.jpg', fileName: 'dated.jpg', width: 320, height: 180 },
    { name: 'filename-before-exif', fixture: 'dated-world.jpg', fileName: 'VRChat_2025-01-02_03-04-05.jpg', width: 320, height: 180 },
    { name: 'png-world-xmp', fixture: 'world.png', fileName: 'world.png', width: 320, height: 180 },
    { name: 'webp-header-over-filename', fixture: 'plain.webp', fileName: 'wrong_999x888.webp' },
    { name: 'webp-native-over-header', fixture: 'plain.webp', fileName: 'wrong_999x888.webp', width: 640, height: 360 },
    { name: 'unknown-filename-size-fallback', fixture: 'plain.png', fileName: 'unknown_1920x1080.png' },
    { name: 'unknown-without-size', fixture: 'plain.png', fileName: 'unknown.png' },
    { name: 'decoded-size-over-filename', fixture: 'plain.jpg', fileName: 'VRChat_2025-01-02_03-04-05_999x888.jpg', width: 320, height: 180 },
    { name: 'invalid-decoded-size-falls-back', fixture: 'plain.png', fileName: 'fallback_1080x1920.png', width: NaN, height: 0 },
    { name: 'read-failure', fixture: 'plain.png', fileName: 'photo.png', fail: 'read' },
    { name: 'decode-failure', fixture: 'plain.png', fileName: 'photo.png', fail: 'decode' },
    { name: 'stat-failure', fixture: 'plain.png', fileName: 'photo.png', fail: 'stat' },
    { name: 'thumbnail-sync-failure', fixture: 'plain.png', fileName: 'photo.png', fail: 'thumbnail-sync' },
    { name: 'thumbnail-async-failure', fixture: 'plain.png', fileName: 'VRChat_2026-06-07_22-20-00.png', fail: 'thumbnail-async' },
    { name: 'thumbnail-unavailable', fixture: 'plain.png', fileName: 'photo.png', thumbnail: null },
  ];
}

function normalizeResult(value) {
  if (value instanceof Date || Object.prototype.toString.call(value) === '[object Date]') {
    return { date: Number.isNaN(value.getTime()) ? 'invalid' : value.toISOString() };
  }
  if (value === undefined) return { undefined: true };
  if (typeof value === 'number' && !Number.isFinite(value)) return { number: String(value) };
  if (typeof value === 'bigint') return { bigint: String(value) };
  if (Array.isArray(value)) return Array.from(value, normalizeResult);
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, normalizeResult(entry)]));
  }
  return value;
}

async function runRecordCase(makeBuilder, spec) {
  const trace = [];
  const filePath = `C:/fixture/${spec.fileName}`;
  const buffer = readFixture(spec.fixture);
  const ports = {
    async readFile(requestedPath) {
      trace.push(['read', requestedPath]);
      if (spec.fail === 'read') throw new Error('fixture read failure');
      return buffer;
    },
    async stat(requestedPath) {
      trace.push(['stat', requestedPath]);
      if (spec.fail === 'stat') throw new Error('fixture stat failure');
      return { mtime: new Date(STAT_DATE) };
    },
    decodeImage(bytes) {
      trace.push(['decode', bytes.equals(buffer)]);
      if (spec.fail === 'decode') throw new Error('fixture decode failure');
      return { image: { fixture: spec.fixture }, imageWidth: spec.width ?? null,
        imageHeight: spec.height ?? null, resolutionTier: 'fixture-resolution',
        orientationTier: 'fixture-orientation' };
    },
    createThumbnail(requestedPath, fileHash, options) {
      trace.push(['thumbnail:start', requestedPath, fileHash, options]);
      if (spec.fail === 'thumbnail-sync') throw new Error('fixture thumbnail sync failure');
      return Promise.resolve().then(() => {
        trace.push(['thumbnail:done']);
        if (spec.fail === 'thumbnail-async') throw new Error('fixture thumbnail async failure');
        return spec.thumbnail === null ? null : 'C:/fixture/thumbnail.jpg';
      });
    },
    now: () => new Date(FIXED_NOW),
  };
  try {
    return { value: normalizeResult(await makeBuilder(ports)(filePath)), trace: normalizeResult(trace) };
  } catch (error) {
    return { error: { name: error.name, message: error.message }, trace: normalizeResult(trace) };
  }
}

module.exports = { createMetadataCases, createRecordCases, normalizeResult, readFixture,
  runRecordCase, FIXED_NOW, STAT_DATE, WORLD_ID };

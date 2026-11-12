'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { test } = require('node:test');
const { createWorldThumbnailStore } = require('../src/world-thumbnail-store');

const WORLD_ID = 'wrld_11111111-1111-4111-8111-111111111111';
const FILE_URL =
  'https://api.vrchat.cloud/api/1/file/file_aaaaaaaa-0000-4000-8000-000000000000/7/file';
const RESIZED_URL =
  'https://api.vrchat.cloud/api/1/image/file_aaaaaaaa-0000-4000-8000-000000000000/7/512';

function createFakeNativeImage() {
  const toImage = (width, height) => ({
    isEmpty: () => false,
    getSize: () => ({ width, height }),
    resize: ({ width: nextWidth }) =>
      toImage(nextWidth, Math.round((height * nextWidth) / width)),
    toJPEG: () => Buffer.from(`jpeg ${width}x${height}`),
  });

  return {
    createFromBuffer: (buffer) =>
      buffer.toString() === 'broken'
        ? { isEmpty: () => true }
        : toImage(1200, 900),
  };
}

function createResponse({ status = 200, contentType = 'image/png', body = 'png' } = {}) {
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: {
      get: (name) => (name === 'content-type' ? contentType : null),
    },
    arrayBuffer: async () => Buffer.from(body),
  };
}

async function createStore(responses) {
  const directoryPath = await fs.mkdtemp(
    path.join(os.tmpdir(), 'worldshot-world-thumbnails-')
  );
  const requestedUrls = [];
  const store = createWorldThumbnailStore({
    fs,
    nativeImage: createFakeNativeImage(),
    fetchImpl: async (url) => {
      requestedUrls.push(url);
      return responses[url] || createResponse({ status: 404, contentType: 'text/plain' });
    },
    getDirectoryPath: () => directoryPath,
    userAgent: 'test',
  });

  return { store, directoryPath, requestedUrls };
}

test('only https URLs on the VRChat file hosts are accepted as a source', async () => {
  const { store } = await createStore({});

  assert.equal(store.normalizeSourceUrl(FILE_URL), FILE_URL);
  assert.equal(
    store.normalizeSourceUrl('https://files.vrchat.cloud/thumbnails/a.png'),
    'https://files.vrchat.cloud/thumbnails/a.png'
  );

  for (const rejected of [
    'http://api.vrchat.cloud/api/1/file/file_a/1/file',
    'https://evilvrchat.cloud/a.png',
    'https://assets.vrchat.com/www/images/default_social_image.jpg',
    'https://example.com/a.png',
    'https://www.youtube.com/watch?v=abc',
    'file:///C:/a.png',
    'not a url',
    '',
    null,
  ]) {
    assert.equal(store.normalizeSourceUrl(rejected), null, String(rejected));
  }
});

test('the resized copy is derived only from a VRChat file URL', async () => {
  const { store } = await createStore({});

  assert.equal(store.getResizedSourceUrl(FILE_URL), RESIZED_URL);
  assert.equal(store.getResizedSourceUrl(RESIZED_URL), null);
  assert.equal(
    store.getResizedSourceUrl('https://files.vrchat.cloud/thumbnails/a.png'),
    null
  );
});

test('the thumbnail is saved small, from the resized copy first', async () => {
  const { store, directoryPath, requestedUrls } = await createStore({
    [RESIZED_URL]: createResponse(),
    [FILE_URL]: createResponse(),
  });

  const result = await store.saveFromUrl(WORLD_ID, FILE_URL);

  assert.deepEqual(requestedUrls, [RESIZED_URL]);
  assert.equal(result.ok, true);
  assert.equal(result.fileName, `${WORLD_ID}.jpg`);
  assert.equal(result.sourceUrl, FILE_URL);
  assert.equal(result.width, 480);
  assert.equal(result.height, 360);
  assert.equal(await store.hasFile(result.fileName), true);
  assert.equal(
    (await fs.readFile(path.join(directoryPath, result.fileName))).toString(),
    'jpeg 480x360'
  );
  assert.deepEqual(await fs.readdir(directoryPath), [`${WORLD_ID}.jpg`]);
});

test('the full file is used when the resized copy is missing', async () => {
  const { store, requestedUrls } = await createStore({
    [FILE_URL]: createResponse(),
  });

  const result = await store.saveFromUrl(WORLD_ID, FILE_URL);

  assert.deepEqual(requestedUrls, [RESIZED_URL, FILE_URL]);
  assert.equal(result.ok, true);
});

test('a response that is not an image is never saved', async () => {
  for (const response of [
    createResponse({ contentType: 'text/html; charset=utf-8', body: '<html>' }),
    createResponse({ contentType: 'video/mp4', body: 'video' }),
    createResponse({ contentType: 'image/png', body: 'broken' }),
  ]) {
    const { store, directoryPath } = await createStore({
      [RESIZED_URL]: response,
      [FILE_URL]: response,
    });
    const result = await store.saveFromUrl(WORLD_ID, FILE_URL);

    assert.equal(result.ok, false);
    assert.equal(result.fetchStatus, 'not_image');
    assert.deepEqual(await fs.readdir(directoryPath), []);
  }
});

test('a missing image and an unusable source are reported, not thrown', async () => {
  const { store, requestedUrls } = await createStore({});

  assert.deepEqual(await store.saveFromUrl(WORLD_ID, FILE_URL), {
    ok: false,
    fetchStatus: 'not_found',
    fetchError: 'HTTP 404',
  });
  assert.equal(
    (await store.saveFromUrl(WORLD_ID, 'https://example.com/a.png')).fetchStatus,
    'no_source'
  );
  assert.equal(requestedUrls.length, 2);
});

test('file names stay inside the thumbnail folder', async () => {
  const { store, directoryPath } = await createStore({});

  assert.equal(store.getFileName('../../evil'), 'evil.jpg');
  assert.equal(store.getFileName(''), null);
  assert.equal(store.getFilePath('../outside.jpg'), null);
  assert.equal(store.getFilePath(null), null);
  assert.equal(
    store.getFilePath(`${WORLD_ID}.jpg`),
    path.join(directoryPath, `${WORLD_ID}.jpg`)
  );
  assert.equal(await store.hasFile('../outside.jpg'), false);
});
